import os
import asyncio
from io import BytesIO
import pytest
from unittest.mock import patch
from docx import Document

# Set TEST_DATABASE_URL=postgresql://... to run the suite against Postgres (use a throwaway database).
os.environ['DATABASE_URL'] = os.environ.get('TEST_DATABASE_URL', 'sqlite:///./test.db')
os.environ['LLM_PROVIDER'] = 'ollama'
os.environ['LLM_API_KEY'] = ''
if os.path.exists('test.db'):
    os.remove('test.db')
if not os.environ['DATABASE_URL'].startswith('sqlite'):
    from sqlalchemy import create_engine, text
    _url = os.environ['DATABASE_URL'].replace('postgres://', 'postgresql+psycopg://', 1).replace('postgresql://', 'postgresql+psycopg://', 1)
    with create_engine(_url).begin() as _conn:
        _conn.execute(text('DROP SCHEMA public CASCADE; CREATE SCHEMA public'))
from fastapi.testclient import TestClient

from app.main import app
from app.config import settings
from app.schemas import FieldIn
from app.services import (
    generate_local, validate_answer, detect_open_ended, resolve_value,
    context_for, question_guidance, build_prompt, clean_answer,
    unsupported_terms, option_similarity, primary_text, relevance_score, missing_facts,
    off_topic, choose_model,
)

client = TestClient(app)
with client:
    API_KEY = client.post('/auth/register').json()['api_key']
    client.headers.update({'Authorization': f'Bearer {API_KEY}'})
    USER_ID = client.get('/auth/me').json()['user_id']


def new_client():
    """A second, independent extension install."""
    other = TestClient(app)
    key = other.post('/auth/register').json()['api_key']
    other.headers.update({'Authorization': f'Bearer {key}'})
    return other


def field(**overrides):
    payload = dict(
        id='f', tag='input', type='text', name='', id_attribute='', label='', placeholder='',
        aria_label='', nearby_text='', required=False, options=[], section='', question_hint='',
        question='', autocomplete='', option_values=[], current_value='', visible=True)
    payload.update(overrides)
    return FieldIn(**payload)


def test_profile_and_classification():
    with client:
        assert client.put('/profile', json={'full_name': 'Ada Lovelace', 'email': 'ada@example.test', 'phone': '123'}).status_code == 200
        r = client.post('/fields/classify', json={'fields': [{'id': 'a', 'label': 'Email Address'}, {'id': 'b', 'label': 'University Name'}]}).json()['fields']
        assert r[0]['mapped_field'] == 'personal.email' and r[0]['value'] == 'ada@example.test'
        assert r[1]['mapped_field'] == 'education.institution'
        open_ended = client.post('/fields/classify', json={'fields': [{'id': 'essay', 'tag': 'textarea', 'label': 'Tell us about yourself'}]}).json()['fields'][0]
        assert open_ended['kind'] == 'open_ended' and open_ended['question'] == 'Tell us about yourself'
        unfamiliar = client.post('/fields/classify', json={'fields': [{'id': 'fit', 'label': 'What makes you a good fit for this role?'}]}).json()['fields'][0]
        assert unfamiliar['kind'] == 'open_ended' and unfamiliar['question'] == 'What makes you a good fit for this role?'


def test_crud_and_unknown_field():
    with client:
        created = client.post('/education', json={'institution': 'Local University', 'degree': 'BSc'})
        assert created.status_code == 200
        assert client.get('/education').status_code == 200
        r = client.post('/fields/classify', json={'fields': [{'id': 'x', 'label': 'Favourite colour'}]}).json()['fields'][0]
        assert r['mapped_field'] is None and r['confidence'] < settings.suggest_threshold
        assert client.post('/cv/upload', files={'file': ('not-a-cv.txt', b'hello', 'text/plain')}).status_code == 415
        assert client.post('/settings/delete-all-data', json={'confirm': False}).status_code == 422
        assert client.get('/').status_code == 200
        assert client.get('/ui/ui.js').status_code == 200


def test_cv_extracts_docx_tables_and_restores_after_refresh():
    document = Document()
    document.add_heading('Professional Experience', level=1)
    document.add_paragraph('Built reliable data pipelines for a local research team.')
    document.add_heading('Education', level=1)
    table = document.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = 'North College'
    table.rows[0].cells[1].text = 'Bachelor of Science'
    content = BytesIO()
    document.save(content)
    with client:
        response = client.post('/cv/upload', files={'file': ('resume.docx', content.getvalue(), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')})
        assert response.status_code == 200
        extracted = response.json()
        assert 'North College' in extracted['text']
        assert extracted['sections']['experience']
        restored = client.get('/cv').json()
        assert restored['available'] is True
        assert restored['filename'] == 'resume.docx'
        assert restored['text'] == extracted['text']
        context = client.post('/context', json={'question': 'What education have I completed?'}).json()
        assert 'North College' in context['context']
        assert context['question_kind'] == 'education'


def test_new_profile_columns_are_migrated_and_saved():
    with client:
        response = client.put('/profile', json={
            'full_name': 'Ada Lovelace', 'headline': 'Data Analyst', 'current_company': 'Acme Corp',
            'work_authorization': 'Yes', 'notice_period': '1 month', 'summary': 'Analyst focused on forecasting.',
        })
        assert response.status_code == 200
        assert client.get('/profile').json()['work_authorization'] == 'Yes'
        mapped = client.post('/fields/classify', json={'fields': [
            {'id': 'n', 'label': 'Notice Period'},
            {'id': 'a', 'label': 'Work Authorisation'},
            {'id': 'h', 'label': 'Current Job Title'},
        ]}).json()['fields']
        by_id = {item['id']: item for item in mapped}
        assert by_id['n']['mapped_field'] == 'personal.notice_period' and by_id['n']['value'] == '1 month'
        assert by_id['a']['mapped_field'] == 'personal.work_authorization' and by_id['a']['value'] == 'Yes'
        assert by_id['h']['mapped_field'] == 'personal.headline' and by_id['h']['value'] == 'Data Analyst'


def test_scanner_context_disambiguates_repeating_labels():
    with client:
        rows = client.post('/fields/classify', json={'fields': [
            {'id': '1', 'label': 'Start Date', 'section': 'Work Experience'},
            {'id': '2', 'label': 'Institution', 'section': 'Education'},
            {'id': '3', 'label': 'Technologies Used', 'section': 'Projects'},
        ]}).json()['fields']
        by_id = {item['id']: item for item in rows}
        assert by_id['1']['mapped_field'] == 'experience.start_date'
        assert by_id['2']['mapped_field'] == 'education.institution'
        assert by_id['3']['mapped_field'] == 'projects.technologies'


def test_noisy_identifiers_do_not_break_detection():
    with client:
        rows = client.post('/fields/classify', json={'fields': [
            {'id': 'a', 'name': 'ctl00_ContentPlaceHolder_txtUserName', 'label': 'Your full name'},
            {'id': 'b', 'name': 'form_jsp_fld_user_email', 'label': 'E-mail address'},
            {'id': 'c', 'name': 'frmMain_txtPhoneMobile', 'label': ''},
        ]}).json()['fields']
        by_id = {item['id']: item for item in rows}
        assert by_id['a']['mapped_field'] == 'personal.name'
        assert by_id['b']['mapped_field'] == 'personal.email'
        assert by_id['c']['mapped_field'] == 'personal.phone'


def test_dropdown_values_are_matched_to_real_options():
    with client:
        client.post('/settings/delete-all-data', json={'confirm': True})
        client.post('/education', json={'institution': 'North College', 'degree': 'Bachelor of Science'})
        result = client.post('/fields/classify', json={'fields': [{
            'id': 'uni', 'tag': 'select', 'type': 'select-one', 'label': 'University',
            'options': ['Please select', 'North College', 'South Institute'],
            'option_values': ['', 'north', 'south'],
        }]}).json()['fields'][0]
        assert result['mapped_field'] == 'education.institution'
        assert result['value'] == 'North College'
        assert result['option_value'] == 'north'
        assert option_similarity('Bachelor of Science', 'bachelor of science (bsc)') > 0.6


def test_open_ended_detection_is_not_every_textarea():
    assert detect_open_ended(field(tag='textarea', label='Technical Skills')) == (False, '')
    assert detect_open_ended(field(tag='textarea', label='Street Address')) == (False, '')
    assert detect_open_ended(field(tag='textarea', label='Job Description', section='Work Experience'))[0] is False
    kind, question = detect_open_ended(field(tag='textarea', label='Why should we hire you?'))
    assert kind is True and question == 'Why should we hire you?'
    kind, _ = detect_open_ended(field(label='What are your key skills?', tag='textarea'))
    assert kind is True


def test_question_extraction_prefers_the_prompt_over_surrounding_text():
    kind, question = detect_open_ended(field(
        tag='textarea', label='', placeholder='',
        question_hint='Additional comments',
        nearby_text='Application form section 3 of 5 Company name Phone Email Address Submit'))
    assert kind is True
    assert question and 'Application form section' not in question


def test_generation_prompt_is_first_person_and_question_specific():
    context = 'WORK EXPERIENCE\nData Analyst — Acme Corp (2022-01 to present)'
    prompt = build_prompt('Why are you interested in this role?', context, 'standard', question_guidance('Why are you interested in this role?')[1], missing=['a salary expectation'])
    assert 'FIRST PERSON' in prompt
    assert 'Never refer to yourself as "the candidate"' in prompt
    assert 'Answer ONLY the question that is asked' in prompt
    assert 'Focus for this question' in prompt
    assert 'Acme Corp' in prompt
    assert 'Do not summarise the profile' in prompt
    assert 'NOT RECORDED IN THE EVIDENCE' in prompt and 'a salary expectation' in prompt
    assert 'Never invent a figure, date, employer, project or result' in prompt
    assert prompt.count('Hard limit') == 1
    assert question_guidance('Why are you interested in this role?')[0] == 'motivation'
    assert question_guidance('Tell us about yourself.')[0] == 'introduction'
    assert question_guidance('What is your notice period?')[0] == 'availability'
    assert question_guidance('Tell us about your team work experience.')[0] == 'leadership'
    assert question_guidance('What is your expected salary?')[0] == 'compensation'


def test_generation_calls_ollama_with_sane_options():
    captured = {}

    class Response:
        def raise_for_status(self): pass
        def json(self): return {'response': 'I built this feature using Python and SQL, and I am confident because my profile matches the role.'}

    class Client:
        def __init__(self, timeout=None): captured['timeout'] = timeout
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def get(self, url, timeout=None):
            class Tags:
                status_code = 503
            return Tags()
        async def post(self, url, json):
            captured.update(url=url, body=json)
            return Response()

    with patch('app.services.httpx.AsyncClient', Client):
        answer = asyncio.run(generate_local('Why are you a good fit for this role?', 'Skills: Python', 'detailed'))
    assert answer.startswith('I ')
    assert captured['url'].endswith('/api/generate')
    assert captured['body']['model'] == settings.ollama_model
    assert captured['body']['think'] is False
    assert captured['body']['keep_alive'] == '30m'
    assert captured['body']['options']['num_predict'] == 420
    assert captured['body']['options']['num_ctx'] == 8192
    assert 'roughly 180 to 240 words' in captured['body']['prompt']


def test_generation_repairs_third_person_drafts():
    calls = {'n': 0}

    class Response:
        def __init__(self, text): self.text = text
        def raise_for_status(self): pass
        def json(self): return {'response': self.text}

    class Client:
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def get(self, url, timeout=None):
            class Tags:
                status_code = 503
            return Tags()
        async def post(self, url, json):
            calls['n'] += 1
            if calls['n'] == 1:
                return Response('Certainly! Here is the answer: The candidate has strong Python skills and led two projects at Acme Corp.')
            return Response('I have built two projects at Acme Corp using Python and SQL.')

    with patch('app.services.httpx.AsyncClient', lambda timeout=None: Client()):
        answer = asyncio.run(generate_local('What did you build at Acme Corp?', 'Experience: Data Analyst at Acme Corp, built two projects using Python and SQL.', 'standard'))
    assert calls['n'] == 2
    assert answer.startswith('I ')
    assert 'the candidate' not in answer.lower()
    assert 'Certainly' not in answer
    assert 'Here is the answer' not in answer


def test_generation_caps_word_count_and_removes_repeats():
    long_answer = 'I have used Python for dashboards APIs and automation and I also used Python and SQL for reporting. ' * 4

    class Response:
        def raise_for_status(self): pass
        def json(self): return {'response': long_answer}

    class Client:
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def get(self, url, timeout=None):
            class Tags:
                status_code = 503
            return Tags()
        async def post(self, url, json): return Response()

    with patch('app.services.httpx.AsyncClient', lambda timeout=None: Client()):
        answer = asyncio.run(generate_local('What makes you a good fit?', 'Skills: Python, SQL', 'concise'))
    assert len(answer.split()) <= 80
    assert len(clean_answer(long_answer)) < len(long_answer)


def test_context_combines_profile_experience_and_cv():
    with client:
        client.post('/settings/delete-all-data', json={'confirm': True})
        client.put('/profile', json={'full_name': 'Ada Lovelace', 'headline': 'Data Analyst'})
        client.post('/skills', json={'name': 'Python'})
        client.post('/skills', json={'name': 'SQL'})
        client.post('/experience', json={'company': 'Acme Corp', 'position': 'Data Analyst', 'start_date': '2022-01', 'end_date': '', 'description': 'Built a forecasting dashboard in Python.'})
        client.post('/projects', json={'name': 'Forecasting dashboard', 'technologies': 'Python, Streamlit', 'role': 'Developer', 'description': 'Predictive charts for retail sales.'})
        document = Document()
        document.add_paragraph('ada@example.test · +44 7700 900000 · London')
        document.add_heading('Professional Experience', level=1)
        document.add_paragraph('Data Analyst at Acme Corp. I built a forecasting dashboard in Python and owned the reporting pipeline.')
        document.add_heading('Skills', level=1)
        document.add_paragraph('Python, SQL, Streamlit')
        buffer = BytesIO()
        document.save(buffer)
        client.post('/cv/upload', files={'file': ('cv.docx', buffer.getvalue(), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')})

        session = _session()
        context, sources, kind, missing = context_for(session, 'Why are you interested in this role?')
        session.close()
        assert 'Acme Corp' in context and 'Python' in context
        assert set(sources) >= {'profile', 'skills', 'experience', 'cv'}
        assert kind == 'motivation'
        assert isinstance(missing, list)

        response = client.post('/answers/generate', json={'question': 'Why are you interested in this role?', 'length': 'standard'})
        assert response.status_code in (200, 503)


def _session():
    from app.database import SessionLocal
    session = SessionLocal()
    session.info['user_id'] = USER_ID
    return session


def test_approved_answers_are_reused_as_evidence():
    with client:
        client.post('/settings/delete-all-data', json={'confirm': True})
        client.post('/answers/save', json={'question': 'Why do you want this role?', 'answer': 'I want this role because I build forecasting tools in Python.'})
        client.post('/answers/save', json={'question': 'Why do you want this role?', 'answer': 'I want this role because I build forecasting tools in Python.'})
        rows = client.get('/answers').json()
        assert len(rows) == 1
        context = client.post('/context', json={'question': 'Why do you want this role?'}).json()
        assert 'forecast' in context['context'] and 'previous_answers' in context['sources']


def test_validation_catches_voice_relevance_and_invented_detail():
    context = 'Skills: Python, SQL\nExperience: Data Analyst at Acme Corp'
    assert validate_answer('The candidate uses Python daily.', context, 'What is your experience?')[0] is False
    assert validate_answer('I enjoy long walks in the hills and gardening.', context, 'What is your experience with Python?')[0] is False
    valid, message = validate_answer('I used Python and SQL at Acme Corp as a Data Analyst.', context, 'What is your experience with Python?')
    assert valid and 'Passed local checks' in message
    assert unsupported_terms('I worked at Globex and at Acme Corp', context) == ['Globex']
    assert unsupported_terms('The opportunity at [Employer] suits me.', context) == ['Employer']
    assert validate_answer('The role at [Employer] fits my work on the role I did.', context, 'Why this role?')[0] is False
    for placeholder in ('I expect a salary in the range of $X to $Y.', 'My salary is XX per year.', 'Role to be confirmed - TBD.'):
        assert validate_answer(placeholder, context, 'What is your expected salary?')[0] is False, placeholder
    assert relevance_score('I use Python every day', 'What is your experience with Python?') > 0
    assert relevance_score('I use Python every day', 'What is your expected salary?') == 0


def test_question_mark_selects_are_still_classified():
    assert detect_open_ended(field(tag='select', type='select-one', label='Are you authorized to work in the UK?')) == (False, '')
    client_ok = client.post('/fields/classify', json={'fields': [{
        'id': 'auth', 'tag': 'select', 'type': 'select-one', 'label': 'Are you legally authorized to work in the UK?',
        'options': ['Select an option', 'Yes', 'No'], 'option_values': ['', 'yes', 'no'],
    }]}).json()['fields'][0]
    assert client_ok['kind'] == 'simple'
    assert client_ok['mapped_field'] == 'personal.work_authorization'


def test_relevance_ignores_pure_function_words_only_questions():
    assert relevance_score('I am excited about this role and the team.', 'Why do you want to work in this role?') > 0
    assert relevance_score('I am excited about this role and the team.', 'Tell us about your team work experience.') > 0
    assert not off_topic('I am excited about this role and the team.', 'Why do you want to work in this role?')
    assert not off_topic('I tend to be overly meticulous about data quality at Acme Retail.',
                         'What is your greatest weakness?',
                         'Experience: Data Analyst at Acme Retail, forecasting and reporting')
    assert off_topic('I enjoy long walks in the hills.', 'What is your greatest weakness?',
                     'Experience: Data Analyst at Acme Retail')
    assert question_guidance('What is your greatest weakness?')[0] == 'weakness'
    assert question_guidance('What is your greatest strength?')[0] == 'strength'


def test_model_selection_prefers_the_configured_model_then_the_strongest():
    installed = [
        {'name': 'nomic-embed-text:latest', 'size': 2_000_000},
        {'name': 'qwen3:0.6b', 'size': 5_000_000},
        {'name': 'llama3:latest', 'size': 43_000_000},
    ]
    assert choose_model([*installed, {'name': settings.ollama_model, 'size': 1}]) == settings.ollama_model
    assert choose_model(installed) in {'llama3:latest', 'qwen3:0.6b'}
    assert choose_model(installed) != 'nomic-embed-text:latest'
    assert choose_model([]) == settings.ollama_model
    assert client.get('/health').json()['status'] == 'ok'


def test_field_text_prioritises_the_label_over_container_noise():
    noisy = field(label='Email', nearby_text='Company name Phone Address Salary expectations Submit form section 2 email policy')
    assert primary_text(noisy) == 'email'


# --------------------------------------------------------------------------------------
# Hosting: accounts, isolation, limits, hosted LLM
# --------------------------------------------------------------------------------------

def test_requests_without_a_key_are_rejected():
    anonymous = TestClient(app)
    assert anonymous.get('/profile').status_code == 401
    assert anonymous.get('/profile', headers={'Authorization': 'Bearer af_not-a-real-key'}).status_code == 401
    assert anonymous.post('/fields/classify', json={'fields': []}).status_code == 401
    assert anonymous.get('/health').status_code == 200
    assert anonymous.get('/').status_code == 200


def test_users_cannot_see_or_change_each_others_data():
    with client:
        client.put('/profile', json={'full_name': 'Ada Lovelace', 'email': 'ada@example.test'})
        client.post('/skills', json={'name': 'Python'})
        edu_id = client.post('/education', json={'institution': 'North College'}).json()['id']
        other = new_client()
        assert other.get('/profile').json()['full_name'] == ''
        assert other.get('/education').json() == []
        assert other.get('/skills').json() == []
        assert other.post('/skills', json={'name': 'Python'}).status_code == 200  # same name, different user
        assert other.put(f'/education/{edu_id}', json={'institution': 'Hacked'}).status_code == 404
        assert other.delete(f'/education/{edu_id}').status_code == 404
        email = other.post('/fields/classify', json={'fields': [{'id': 'e', 'label': 'Email'}]}).json()['fields'][0]
        assert email['value'] == ''
        context = other.post('/context', json={'question': 'Tell us about your education'}).json()['context']
        assert 'North College' not in context and 'Ada' not in context
        assert other.post('/settings/delete-all-data', json={'confirm': True}).status_code == 200
        assert client.get('/profile').json()['full_name'] == 'Ada Lovelace'
        assert any(row['id'] == edu_id for row in client.get('/education').json())
        assert other.delete('/auth/account').status_code == 200
        assert other.get('/profile').status_code == 401


def test_generation_is_rate_limited_per_user():
    from app.auth import limiter
    limiter.reset()
    other = new_client()
    other.put('/profile', json={'full_name': 'Grace Hopper', 'summary': 'Compiler engineer.'})

    async def fake_generate(*args, **kwargs):
        return 'I build compilers.'

    with patch('app.main.generate_local', fake_generate), patch.object(settings, 'generate_limit_per_hour', 2):
        codes = [other.post('/answers/generate', json={'question': 'Tell us about yourself'}).status_code for _ in range(3)]
        assert codes == [200, 200, 429]
        assert client.post('/answers/generate', json={'question': 'Tell us about yourself'}).status_code == 200
    limiter.reset()


def test_hosted_openai_compatible_provider():
    captured = {}

    class Response:
        def raise_for_status(self): pass
        def json(self): return {'choices': [{'message': {'content': '<think>plan</think>I built forecasting tools in Python at Acme Corp.'}}]}

    class Client:
        def __init__(self, timeout=None): pass
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def post(self, url, json, headers=None):
            captured.update(url=url, body=json, headers=headers)
            return Response()

    with patch.object(settings, 'llm_provider', 'openai'), patch.object(settings, 'llm_api_key', 'sk-test'), \
         patch.object(settings, 'llm_base_url', 'https://api.example.test/v1/'), patch('app.services.httpx.AsyncClient', Client):
        answer = asyncio.run(generate_local('What did you build at Acme Corp?', 'Experience: Acme Corp, Python forecasting tools', 'concise'))
    assert captured['url'] == 'https://api.example.test/v1/chat/completions'
    assert captured['headers']['Authorization'] == 'Bearer sk-test'
    assert captured['body']['model'] == settings.llm_model
    assert captured['body']['max_tokens'] == 140
    assert answer.startswith('I built') and 'think' not in answer


def test_hosted_provider_sends_reasoning_effort_and_extra_tokens():
    captured = {}

    class Response:
        def raise_for_status(self): pass
        def json(self): return {'choices': [{'message': {'content': 'I built forecasting tools in Python at Acme Corp.'}}]}

    class Client:
        def __init__(self, timeout=None): pass
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def post(self, url, json, headers=None):
            captured.update(url=url, body=json)
            return Response()

    with patch.object(settings, 'llm_provider', 'openai'), patch.object(settings, 'llm_api_key', 'sk-test'), \
         patch.object(settings, 'llm_reasoning_effort', 'low'), patch.object(settings, 'llm_max_tokens_extra', 80), \
         patch('app.services.httpx.AsyncClient', Client):
        asyncio.run(generate_local('What did you build at Acme Corp?', 'Experience: Acme Corp, Python forecasting tools', 'concise'))

    assert captured['body']['reasoning_effort'] == 'low'
    assert captured['body']['max_tokens'] == 220


def test_hosted_provider_omits_reasoning_effort_and_keeps_default_tokens():
    captured = {}

    class Response:
        def raise_for_status(self): pass
        def json(self): return {'choices': [{'message': {'content': 'I built forecasting tools in Python at Acme Corp.'}}]}

    class Client:
        def __init__(self, timeout=None): pass
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def post(self, url, json, headers=None):
            captured.update(url=url, body=json)
            return Response()

    with patch.object(settings, 'llm_provider', 'openai'), patch.object(settings, 'llm_api_key', 'sk-test'), \
         patch.object(settings, 'llm_reasoning_effort', ''), patch.object(settings, 'llm_max_tokens_extra', 0), \
         patch('app.services.httpx.AsyncClient', Client):
        asyncio.run(generate_local('What did you build at Acme Corp?', 'Experience: Acme Corp, Python forecasting tools', 'concise'))

    assert 'reasoning_effort' not in captured['body']
    assert captured['body']['max_tokens'] == 140


def test_postgres_urls_are_normalised():
    from app.config import Settings
    assert Settings(database_url='postgres://u:p@h/db').database_url == 'postgresql+psycopg://u:p@h/db'
    assert Settings(database_url='postgresql://u:p@h/db?sslmode=require').database_url == 'postgresql+psycopg://u:p@h/db?sslmode=require'
    assert Settings(llm_provider='auto', llm_api_key='x').provider == 'openai'
    assert Settings(llm_provider='auto', llm_api_key='').provider == 'ollama'


def test_llm_settings_strip_whitespace():
    from app.config import Settings
    settings_with_space = Settings(llm_api_key=' \n fake-key \t', llm_base_url='\n https://api.example.test/v1 \r\n')
    assert settings_with_space.llm_api_key == 'fake-key'
    assert settings_with_space.llm_base_url == 'https://api.example.test/v1'


def test_hosted_failure_log_does_not_include_api_key(caplog):
    import httpx

    fake_key = 'gsk_test_fake'

    class Client:
        def __init__(self, timeout=None): pass
        async def __aenter__(self): return self
        async def __aexit__(self, *args): pass
        async def post(self, *args, **kwargs):
            request = httpx.Request('POST', 'https://api.example.test/v1/chat/completions')
            raise httpx.ConnectError(f'failed with Authorization: Bearer {fake_key}', request=request)

    with patch.object(settings, 'llm_provider', 'openai'), patch.object(settings, 'llm_api_key', fake_key), \
         patch.object(settings, 'llm_base_url', 'https://api.example.test/v1'), \
         patch.object(settings, 'llm_model', 'test-model'), patch('app.services.httpx.AsyncClient', Client):
        with pytest.raises(RuntimeError, match='could not be reached'):
            asyncio.run(generate_local('Tell me about yourself', 'Profile: engineer'))

    captured = caplog.text
    assert 'ConnectError' in captured
    assert 'https://api.example.test/v1' in captured
    assert 'test-model' in captured
    assert fake_key not in captured
