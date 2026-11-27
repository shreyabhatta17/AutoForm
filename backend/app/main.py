from contextlib import asynccontextmanager
from datetime import datetime
import logging
from pathlib import Path
from fastapi import FastAPI, Depends, HTTPException, UploadFile, File, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import select
from .database import Base, engine, get_db, add_missing_columns
from .models import User, Profile, Education, Experience, Skill, Project, PreviousAnswer, CVDocument
from .schemas import *
from .config import settings
from .auth import current_user, create_user, limiter, client_ip
from .services import classifier, resolve_value, detect_open_ended, context_for, generate_local, validate_answer, relevance_score, first_person_ratio, extract_cv_text, cv_sections, question_guidance, probe_model

USER_TABLES = (Education, Experience, Skill, Project, PreviousAnswer, CVDocument, Profile)


@asynccontextmanager
async def lifespan(_app):
    key = settings.llm_api_key
    logging.getLogger('autoform').info(
        'LLM config: key_length=%d key_starts_gsk=%s key_has_whitespace_or_nonprintable=%s base_url=%s',
        len(key), key.startswith('gsk_'), any(c.isspace() or not c.isprintable() for c in key), settings.llm_base_url)
    Base.metadata.create_all(engine)
    add_missing_columns()
    yield

app=FastAPI(title='AutoForm API',version='1.1.0',lifespan=lifespan)
app.mount('/ui', StaticFiles(directory=Path(__file__).parent/'static'), name='ui')
app.add_middleware(CORSMiddleware,
    allow_origin_regex=r'(chrome-extension|moz-extension|safari-web-extension)://.*',
    allow_origins=[o.strip() for o in settings.cors_origins.split(',') if o.strip()],
    allow_methods=['GET','POST','PUT','DELETE'],allow_headers=['Authorization','Content-Type','X-API-Key'])

@app.get('/', include_in_schema=False)
def dashboard(): return FileResponse(Path(__file__).parent/'static'/'index.html')

@app.get('/health')
def health():
    return {'status':'ok','model_loaded':classifier.model is not None,'llm_provider':settings.provider,
            'answer_model':probe_model(),'configured_model':settings.llm_model if settings.provider=='openai' else settings.ollama_model}

# ---- Accounts ----------------------------------------------------------------------

@app.post('/auth/register')
def register(request:Request,db:Session=Depends(get_db)):
    limiter.check(f'register:{client_ip(request)}', settings.register_limit_per_hour)
    return {'api_key':create_user(db)}

@app.get('/auth/me')
def me(uid:str=Depends(current_user)): return {'user_id':uid}

@app.delete('/auth/account')
def delete_account(uid:str=Depends(current_user),db:Session=Depends(get_db)):
    _wipe(db,uid)
    db.info.pop('user_id',None)
    user=db.get(User,uid)
    if user: db.delete(user)
    db.commit()
    return {'deleted':True}

def _wipe(db,uid):
    for model in USER_TABLES: db.query(model).filter(model.user_id==uid).delete(synchronize_session=False)

def owned(db,model,item_id,uid):
    row=db.get(model,item_id)
    if not row or row.user_id!=uid: raise HTTPException(404,'Item not found')
    return row

# ---- Profile -----------------------------------------------------------------------

def one(model,db,uid):
    row=db.scalar(select(model).where(model.user_id==uid))
    if not row: row=model(user_id=uid); db.add(row); db.commit(); db.refresh(row)
    return row
@app.get('/profile')
def profile(uid:str=Depends(current_user),db:Session=Depends(get_db)): return one(Profile,db,uid)
@app.put('/profile')
def update_profile(data:ProfileIn,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    row=one(Profile,db,uid)
    for k,v in data.model_dump().items(): setattr(row,k,v)
    db.commit(); db.refresh(row); return row
@app.delete('/profile')
def delete_profile(uid:str=Depends(current_user),db:Session=Depends(get_db)):
    db.query(Profile).filter(Profile.user_id==uid).delete(synchronize_session=False); db.commit()
    return {'deleted':True}

def crud(prefix,model,schema):
    @app.get('/'+prefix)
    def list_(uid:str=Depends(current_user),db:Session=Depends(get_db)): return db.scalars(select(model).where(model.user_id==uid).order_by(model.id)).all()
    @app.post('/'+prefix)
    def create(data:schema,uid:str=Depends(current_user),db:Session=Depends(get_db)):
        row=model(**data.model_dump(),user_id=uid); db.add(row); db.commit(); db.refresh(row); return row
    @app.put('/'+prefix+'/{item_id}')
    def update(item_id:int,data:schema,uid:str=Depends(current_user),db:Session=Depends(get_db)):
        row=owned(db,model,item_id,uid)
        for k,v in data.model_dump().items(): setattr(row,k,v)
        db.commit(); db.refresh(row); return row
    @app.delete('/'+prefix+'/{item_id}')
    def delete(item_id:int,uid:str=Depends(current_user),db:Session=Depends(get_db)):
        row=owned(db,model,item_id,uid)
        db.delete(row);db.commit();return {'deleted':True}
crud('education',Education,EducationIn); crud('experience',Experience,ExperienceIn); crud('projects',Project,ProjectIn); crud('answers',PreviousAnswer,AnswerIn)

@app.get('/skills')
def skills(uid:str=Depends(current_user),db:Session=Depends(get_db)): return db.scalars(select(Skill).where(Skill.user_id==uid).order_by(Skill.name)).all()
@app.post('/skills')
def add_skill(data:SkillIn,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    if db.scalar(select(Skill).where(Skill.user_id==uid,Skill.name==data.name)): raise HTTPException(409,'Skill already exists')
    x=Skill(**data.model_dump(),user_id=uid);db.add(x);db.commit();db.refresh(x);return x
@app.delete('/skills/{item_id}')
def delete_skill(item_id:int,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    x=owned(db,Skill,item_id,uid)
    db.delete(x);db.commit();return {'deleted':True}

# ---- Form scanning -----------------------------------------------------------------

@app.post('/fields/classify')
def classify(payload:ClassifyRequest,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    if len(payload.fields)>settings.max_fields_per_scan: raise HTTPException(413,f'Too many fields in one scan (max {settings.max_fields_per_scan}).')
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

# ---- Written answers ---------------------------------------------------------------

@app.post('/context')
def context(payload:GenerateRequest,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    value,sources,kind,missing=context_for(db,payload.question);return {'context':value,'sources':sources,'question_kind':kind,'missing':missing}
@app.post('/answers/generate')
async def generate(payload:GenerateRequest,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    context,sources,kind,missing=context_for(db,payload.question)
    if not context.strip(): raise HTTPException(422,'Add profile information before generating an answer.')
    limiter.check(f'generate:{uid}', settings.generate_limit_per_hour)
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
def validate(payload:AnswerIn,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    context,_,_,_=context_for(db,payload.question); valid,message=validate_answer(payload.answer,context,payload.question);return {'valid':valid,'message':message}
@app.post('/answers/save')
def save_answer(payload:SaveAnswerRequest,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    existing=db.scalar(select(PreviousAnswer).where(PreviousAnswer.user_id==uid,PreviousAnswer.question==payload.question))
    if existing:
        existing.answer=payload.answer; existing.updated_at=datetime.utcnow()
    else:
        db.add(PreviousAnswer(question=payload.question,answer=payload.answer,user_id=uid))
    db.commit()
    return {'saved':True}

# ---- CV ----------------------------------------------------------------------------

@app.post('/cv/upload')
async def upload_cv(file:UploadFile=File(...),uid:str=Depends(current_user),db:Session=Depends(get_db)):
    suffix=Path(file.filename or '').suffix.lower()
    if suffix not in {'.pdf','.docx'}: raise HTTPException(415,'Only PDF and DOCX files are supported.')
    content=await file.read(settings.max_cv_bytes+1)
    if len(content)>settings.max_cv_bytes: raise HTTPException(413,'CV exceeds the 10 MB limit.')
    limiter.check(f'cv:{uid}', settings.cv_upload_limit_per_hour)
    try:
        text,sections=extract_cv_text(content,suffix)
        if not text: raise ValueError('The CV contains no readable text.')
    except Exception as e:
        raise HTTPException(422,'The CV could not be read. Check that it is a text-based PDF or a valid DOCX file.') from e
    # Only the extracted text is kept (it is what answer generation uses); the
    # original file is never written to disk.
    db.query(CVDocument).filter(CVDocument.user_id==uid).delete(synchronize_session=False)
    db.add(CVDocument(filename=(file.filename or 'cv')[:500],extracted_text=text[:100000],user_id=uid))
    db.commit()
    return {'available':True,'filename':file.filename or 'cv','text':text[:100000],'sections':sections,'review_required':True,'saved_for_answer_context':True}

@app.get('/cv')
def latest_cv(uid:str=Depends(current_user),db:Session=Depends(get_db)):
    cv=db.scalars(select(CVDocument).where(CVDocument.user_id==uid).order_by(CVDocument.created_at.desc())).first()
    if not cv: return {'available':False}
    return {'available':True,'filename':cv.filename,'text':cv.extracted_text,'sections':cv_sections(cv.extracted_text),'saved_for_answer_context':True}

@app.post('/settings/delete-all-data')
def delete_all_data(payload:ConfirmDelete,uid:str=Depends(current_user),db:Session=Depends(get_db)):
    if not payload.confirm: raise HTTPException(422,'Set confirm to true to delete all of your data.')
    _wipe(db,uid); db.commit()
    return {'deleted':True}
