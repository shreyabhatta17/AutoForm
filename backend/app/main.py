import uuid
from datetime import datetime
from pathlib import Path
from fastapi import FastAPI, Depends, HTTPException, UploadFile, File, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import select
from .database import Base, engine, get_db, SessionLocal, add_missing_columns
from .models import Profile, Education, Experience, Skill, Project, PreviousAnswer, CVDocument
from .schemas import *
from .config import settings
from .services import classifier, get_value, resolve_value, field_text, detect_open_ended, context_for, generate_local, validate_answer, relevance_score, first_person_ratio, extract_cv_text, cv_sections, question_guidance, probe_model

app=FastAPI(title='Automatic Form Filler',version='1.0.0')
app.mount('/ui', StaticFiles(directory=Path(__file__).parent/'static'), name='ui')
app.add_middleware(CORSMiddleware,allow_origin_regex=r'chrome-extension://.*',allow_methods=['GET','POST','PUT','DELETE'],allow_headers=['*'])
PRIVATE_ROUTE_PREFIXES = ('/profile', '/education', '/experience', '/projects', '/skills', '/answers', '/cv', '/context', '/fields/classify', '/settings')

@app.middleware('http')
async def public_demo_privacy_guard(request: Request, call_next):
    path = request.url.path.rstrip('/') or '/'
    if settings.public_demo and (path == '/' or any(path == prefix or path.startswith(prefix + '/') for prefix in PRIVATE_ROUTE_PREFIXES)):
        from fastapi.responses import JSONResponse
        return JSONResponse(status_code=403, content={'detail': 'This personal-data feature is disabled in public demo mode.'})
    return await call_next(request)
@app.on_event('startup')
def init():
    Base.metadata.create_all(engine)
    add_missing_columns()
@app.get('/', include_in_schema=False)
def dashboard(): return FileResponse(Path(__file__).parent/'static'/'index.html')
@app.get('/health')
def health():
    if settings.public_demo:
        return {'status':'ok','deployment_mode':'public-demo','model_loaded':classifier.model is not None,'answer_generation_available':False}
    try:
        model = probe_model()
    except Exception:
        model = None
    return {'status':'ok','deployment_mode':settings.deployment_mode,'model_loaded':classifier.model is not None,'ollama_url':settings.ollama_base_url,'answer_model':model,'configured_model':settings.ollama_model}
def one(model, db):
    row=db.get(model,1)
    if not row: row=model(id=1); db.add(row); db.commit(); db.refresh(row)
    return row
@app.get('/profile')
def profile(db:Session=Depends(get_db)): return one(Profile,db)
@app.put('/profile')
def update_profile(data:ProfileIn,db:Session=Depends(get_db)):
    row=one(Profile,db)
    for k,v in data.model_dump().items(): setattr(row,k,v)
    db.commit(); db.refresh(row); return row
@app.delete('/profile')
def delete_profile(db:Session=Depends(get_db)):
    row=db.get(Profile,1)
    if row: db.delete(row); db.commit()
    return {'deleted':True}
def crud(prefix,model,schema):
    @app.get('/'+prefix)
    def list_(db:Session=Depends(get_db)): return db.scalars(select(model)).all()
    @app.post('/'+prefix)
    def create(data:schema,db:Session=Depends(get_db)):
        row=model(**data.model_dump()); db.add(row); db.commit(); db.refresh(row); return row
    @app.put('/'+prefix+'/{item_id}')
    def update(item_id:int,data:schema,db:Session=Depends(get_db)):
        row=db.get(model,item_id)
        if not row: raise HTTPException(404,'Item not found')
        for k,v in data.model_dump().items(): setattr(row,k,v)
        db.commit(); db.refresh(row); return row
    @app.delete('/'+prefix+'/{item_id}')
    def delete(item_id:int,db:Session=Depends(get_db)):
        row=db.get(model,item_id)
        if not row: raise HTTPException(404,'Item not found')
        db.delete(row);db.commit();return {'deleted':True}
crud('education',Education,EducationIn); crud('experience',Experience,ExperienceIn); crud('projects',Project,ProjectIn); crud('answers',PreviousAnswer,AnswerIn)
@app.get('/skills')
def skills(db:Session=Depends(get_db)): return db.scalars(select(Skill)).all()
@app.post('/skills')
def add_skill(data:SkillIn,db:Session=Depends(get_db)):
    if db.scalar(select(Skill).where(Skill.name==data.name)): raise HTTPException(409,'Skill already exists')
    x=Skill(**data.model_dump());db.add(x);db.commit();db.refresh(x);return x
@app.delete('/skills/{item_id}')
def delete_skill(item_id:int,db:Session=Depends(get_db)):
    x=db.get(Skill,item_id)
    if not x: raise HTTPException(404,'Skill not found')
    db.delete(x);db.commit();return {'deleted':True}
@app.post('/fields/classify')
def classify(payload:ClassifyRequest,db:Session=Depends(get_db)):
    result=[]
    for f in payload.fields:
        try:
            mapped,confidence,source=classifier.predict(f)
            confidence=float(confidence or 0)
        except Exception:
            mapped,confidence,source=None,0.0,'none'

        matched=mapped is not None and confidence>=settings.suggest_threshold
        value,option_value='',None
        if matched:
            try:
                value,option_value=resolve_value(db,mapped,f.options,f.option_values)
                value=str(value or '')
            except Exception:
                # One unsupported control or bad option list must not cancel
                # classification for the rest of the form.
                value,option_value='',None

        # Existing profile/database data wins over essay generation. Only
        # offer a generated answer when there is no saved value to fill.
        if not value:
            try:
                open_ended,question=detect_open_ended(f)
            except Exception:
                open_ended,question=False,''
            if open_ended:
                result.append({'id':f.id,'mapped_field':None,'confidence':0.0,'value':'','option_value':None,'kind':'open_ended','question':question or f.label or f.placeholder or f.nearby_text,'requires_review':True,'detected_label':f.label or f.name or f.id_attribute})
                continue

        result.append({'id':f.id,'mapped_field':mapped if matched else None,'confidence':confidence,'value':value,'option_value':option_value,'kind':'simple','source':source,'requires_review':confidence<settings.auto_fill_threshold,'detected_label':f.label or f.name or f.id_attribute})
    return {'fields':result}
@app.post('/context')
def context(payload:GenerateRequest,db:Session=Depends(get_db)):
    value,sources,kind,missing=context_for(db,payload.question);return {'context':value,'sources':sources,'question_kind':kind,'missing':missing}
@app.post('/answers/generate')
async def generate(payload:GenerateRequest,db:Session=Depends(get_db)):
    if settings.public_demo:
        raise HTTPException(503, 'Answer generation is unavailable in public demo mode.')
    context,sources,kind,missing=context_for(db,payload.question)
    if not context.strip(): raise HTTPException(422,'Add profile information before generating an answer.')
    _,guidance,_=question_guidance(payload.question)
    try: answer=await generate_local(payload.question,context,payload.length,guidance,missing)
    except RuntimeError as e: raise HTTPException(503,str(e))
    valid,message=validate_answer(answer,context,payload.question)
    relevance=round(min(1.0,relevance_score(answer,payload.question)/2.0),3)
    grounding=1.0 if valid else 0.0
    voice=round(min(1.0,first_person_ratio(answer)*12),3)
    confidence=round(0.2+0.35*relevance+0.3*grounding+0.15*voice,3)
    return {'answer':answer,'confidence':confidence,'relevance':relevance,'first_person':voice>=0.5,'sources':sources,'missing':missing,'question_kind':kind,'valid':valid,'validation_message':message}
@app.post('/answers/validate')
def validate(payload:AnswerIn,db:Session=Depends(get_db)):
    context,_,_,_=context_for(db,payload.question); valid,message=validate_answer(payload.answer,context,payload.question);return {'valid':valid,'message':message}
@app.post('/answers/save')
def save_answer(payload:SaveAnswerRequest,db:Session=Depends(get_db)):
    existing=db.scalar(select(PreviousAnswer).where(PreviousAnswer.question==payload.question))
    if existing:
        existing.answer=payload.answer; existing.updated_at=datetime.utcnow()
    else:
        db.add(PreviousAnswer(question=payload.question,answer=payload.answer))
    db.commit()
    return {'saved':True}
@app.post('/cv/upload')
async def upload_cv(file:UploadFile=File(...),db:Session=Depends(get_db)):
    suffix=Path(file.filename or '').suffix.lower()
    if suffix not in {'.pdf','.docx'}: raise HTTPException(415,'Only PDF and DOCX files are supported.')
    content=await file.read()
    if len(content)>settings.max_cv_bytes: raise HTTPException(413,'CV exceeds the 10 MB limit.')
    try:
        text,sections=extract_cv_text(content,suffix)
        if not text: raise ValueError('The CV contains no readable text.')
    except Exception as e:
        raise HTTPException(422,'The CV could not be read. Check that it is a text-based PDF or a valid DOCX file.') from e
    dest=settings.upload_dir/f'{uuid.uuid4()}{suffix}'
    try:
        dest.write_bytes(content)
        db.query(CVDocument).delete()
        db.add(CVDocument(filename=file.filename or 'cv',extracted_text=text[:100000]))
        db.commit()
    except Exception:
        db.rollback(); dest.unlink(missing_ok=True); raise
    return {'available':True,'filename':file.filename or 'cv','text':text[:100000],'sections':sections,'review_required':True,'saved_for_answer_context':True}

@app.get('/cv')
def latest_cv(db:Session=Depends(get_db)):
    cv=db.scalars(select(CVDocument).order_by(CVDocument.created_at.desc())).first()
    if not cv: return {'available':False}
    return {'available':True,'filename':cv.filename,'text':cv.extracted_text,'sections':cv_sections(cv.extracted_text),'saved_for_answer_context':True}
@app.post('/settings/delete-all-data')
def delete_all_data(payload:ConfirmDelete,db:Session=Depends(get_db)):
    if not payload.confirm: raise HTTPException(422,'Set confirm to true to delete all local data.')
    for model in (Education,Experience,Skill,Project,PreviousAnswer,CVDocument,Profile): db.query(model).delete()
    db.commit()
    for path in settings.upload_dir.iterdir():
        if path.is_file(): path.unlink()
    return {'deleted':True}
