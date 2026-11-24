import re
import time
from io import BytesIO

import httpx
import joblib
import fitz
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from sqlalchemy.orm import Session
from sqlalchemy import select

from .config import settings
from .models import Profile, Education, Experience, Skill, Project, PreviousAnswer, CVDocument

# --------------------------------------------------------------------------------------
# Field vocabulary
# --------------------------------------------------------------------------------------

FIRST_NAME = '__first_name'
LAST_NAME = '__last_name'

CANONICAL_FIELDS = {
    'personal.name': 'full_name',
    'personal.first_name': FIRST_NAME,
    'personal.last_name': LAST_NAME,
    'personal.email': 'email',
    'personal.phone': 'phone',
    'personal.location': 'location',
    'personal.address': 'address',
    'personal.date_of_birth': 'date_of_birth',
    'personal.linkedin': 'linkedin',
    'personal.github': 'github',
    'personal.portfolio': 'portfolio',
    'personal.headline': 'headline',
    'personal.summary': 'summary',
    'personal.languages': 'languages',
    'personal.current_company': 'current_company',
    'personal.graduation_year': 'graduation_year',
    'personal.work_authorization': 'work_authorization',
    'personal.visa_status': 'visa_status',
    'personal.availability': 'availability',
    'personal.notice_period': 'notice_period',
    'personal.salary_expectation': 'salary_expectation',
    'personal.references': 'references',
    'personal.certifications': 'certifications',
    'personal.awards': 'awards',
    'personal.interests': 'interests',
}

COLLECTION_FIELDS = {
    'education.institution': (Education, 'institution'),
    'education.degree': (Education, 'degree'),
    'education.field': (Education, 'field'),
    'education.start_date': (Education, 'start_date'),
    'education.end_date': (Education, 'end_date'),
    'education.grade': (Education, 'grade'),
    'education.description': (Education, 'description'),
    'experience.company': (Experience, 'company'),
    'experience.position': (Experience, 'position'),
    'experience.start_date': (Experience, 'start_date'),
    'experience.end_date': (Experience, 'end_date'),
    'experience.description': (Experience, 'description'),
    'projects.name': (Project, 'name'),
    'projects.description': (Project, 'description'),
    'projects.technologies': (Project, 'technologies'),
    'projects.role': (Project, 'role'),
    'projects.url': (Project, 'url'),
}

RULES = {
    'personal.email': (r'e-?mail(?:\s*(?:address|id|address\s+id))?|contact\s*email|work\s*email|mail\s*id', 100),
    'personal.phone': (r'phone(?:\s*(?:number|no\.?|#))?|mobile(?:\s*(?:number|no\.?|phone))?|telephone(?:\s*(?:number|no\.?))?|\btel\b|cell(?:\s*(?:number|no\.?|phone))?|contact\s*(?:number|phone|telephone)', 100),
    'personal.date_of_birth': (r'date\s+of\s+birth|\bd\.?o\.?b\.?\b|birth\s*date|birthday|date\s+born', 100),
    'personal.linkedin': (r'linked\s*-?in(?:\s*(?:profile|url|address))?', 100),
    'personal.github': (r'git\s*-?hub(?:\s*(?:profile|url|address))?|source\s*code\s*profile', 100),
    'personal.portfolio': (r'portfolio(?:\s*(?:url|site|link))?|personal\s*(?:website|site|url|blog)|behance|dribbble', 100),
    'personal.address': (r'street\s*address|address\s*line\s*\d?|home\s*address|full\s*address|mailing\s*address|postal\s*address|house\s*(?:number|no)|flat\s*(?:number|no)', 95),
    'personal.availability': (r'availab\w*|available\s+to\s+start|when\s+can\s+you\s+start|earliest\s+start(?:\s+date)?|date\s+you\s+can\s+start|relocat\w*', 95),
    'personal.notice_period': (r'notice\s*period|how\s+soon\s+can\s+you\s+start|resign\w*\s*notice', 95),
    'personal.salary_expectation': (r'salary|compensation|remuneration|pay\s*expectation|expected\s*(?:salary|pay)|desired\s*salary|current\s*salary|salary\s*range|\bpay\b', 90),
    'personal.visa_status': (r'visa(?:\s*status)?|work\s*permit(?:ting)?|immigration\s*status|requires?\s*sponsorship|sponsorship\s*(?:required|needed)', 90),
    'personal.work_authorization': (r'work\s*authori[sz]ation|authori[sz]ed\s+to\s+work|eligible\s+to\s+work|right\s+to\s+work|employment\s*(?:eligibility|status)|legally\s+authori[sz]ed', 90),
    'personal.certifications': (r'certificat\w*|licen[sc]e\w*|accreditation|training\s*(?:courses|certificates)?', 88),
    'personal.awards': (r'award\w*|honou?rs?|accomplishments?|achievements?|recognitions?', 85),
    'personal.references': (r'refere?nces?|reference\s*contact|who\s+can\s+we\s+(?:speak|talk)\s+to|supervisor\s*contact', 85),
    'personal.languages': (r'languages?|spoken\s*language|language\s*proficien\w*', 84),
    'personal.headline': (r'current\s*(?:job\s*)?title|professional\s*title|desired\s*(?:job\s*)?title|target\s*role|headline|occupation|profession', 90),
    'personal.current_company': (r'current\s*(?:employer|company|organisation|organization)|present\s*(?:employer|company)|current\s*workplace', 82),
    'personal.graduation_year': (r'graduation\s*year|year\s+of\s+graduation|grad\s*year|expected\s*graduation', 80),
    'personal.summary': (r'professional\s*summary|career\s*summary|professional\s*profile|career\s*objective|short\s*bio|about\s*me|personal\s*summary|tell\s+(?:us|me)\s+about\s+yourself|introduce\s+yourself|professional\s+background|career\s+overview|brief\s+bio', 84),
    'personal.first_name': (r'first\s*name|given\s*name|forename|firstname', 78),
    'personal.last_name': (r'last\s*name|family\s*name|surname|sur\s*name|lastname', 78),
    'personal.location': (r'location|city|town|country|state\s*/\s*province|province|region|current\s*location|based\s*in|where\s+are\s+you\s+based', 60),
    'education.institution': (r'(?:college|universit\w+|school|institute|institution|academy|campus|polytechnic|alma\s+mater)\w*', 95),
    'education.degree': (r'degree|qualification|diploma|bachelor\w*|master\w*|doctora\w*|\bmba\b|\bmsc\b|\bms\b|\bbsc\b|\bbs\b|\bba\b|\bb\.?sc\b|\bm\.?sc\b|\bph\.?d\b|education\s*level|highest\s*qualification|academic\s*qualification', 92),
    'education.field': (r'field\s+of\s*study|area\s+of\s*study|major|speciali[sz]ation|discipline|course\s+of\s*study|\bsubject\b|\bdepartment\b', 92),
    'education.description': (r'educational\s+background|academic\s+background|describe\s+your\s+education|tell\s+(?:us|me)\s+about\s+your\s+education', 86),
    'education.end_date': (r'graduation\s*date|date\s+of\s+graduation|completion\s*date|education\s*end|(?:education|academic|school|university|college)\s*end\s*date', 88),
    'education.start_date': (r'(?:education|academic|school|university|college)\s*start(?:\s*date)?|joining\s*date\s*of\s*school|from\s*year\s*of\s*study', 86),
    'education.grade': (r'\bgrade\b|\bgpa\b|\bcgpa\b|marks?\b|final\s*score|class\s*obtained', 76),
    'experience.company': (r'compan(?:y|ies)|employer|organisation|organization|\bfirm\b|business\s*name|workplace|previous\s*employer|where\s+did\s+you\s+work|company\s+name', 90),
    'experience.position': (r'job\s*title|\btitle\b|\brole\b|\bposition\b|designation|current\s*title|job\s*role', 86),
    'experience.description': (r'job\s*description|role\s*description|responsibilit\w*|\bduties\b|what\s+did\s+you\s+do|describe\s+your\s+(?:work|experience)|tell\s+(?:us|me)\s+about\s+your\s+(?:(?:work|professional|previous)\s+)?experience|tell\s+(?:us|me)\s+about\s+your\s+work', 86),
    'experience.start_date': (r'(?:employment|job|work|role|position)?\s*start\s*date|employment\s*start|job\s*start|joining\s*date|from\s*date|start\s*date', 88),
    'experience.end_date': (r'(?:employment|job|work|role|position)?\s*end\s*date|employment\s*end|job\s*end|to\s*date|leaving\s*date|end\s*date', 88),
    'projects.technologies': (r'technologies\s*(?:used|in)|tools?\s*(?:used|in)|built\s*with|stack\s*used|tech\s*used', 86),
    'projects.role': (r'your\s*role\s+in|role\s+in\s+(?:the|this|that)\s+project|contribution\s*(?:to|in)|what\s+did\s+you\s+build', 84),
    'projects.url': (r'project\s*(?:url|link|repo)|live\s*(?:link|url|demo)|demo\s*(?:link|url)|repository\s*url|github\s*repo', 84),
    'projects.description': (r'project\s*description|project\s*summary|project\s*details|describe\s+(?:the|your|a)\s+project|tell\s+(?:us|me)\s+about\s+(?:a|the|your)\s+project', 84),
    'projects.name': (r'project\s*name|\bprojects?\b|portfolio\s+project|case\s+study|initiative\s+name', 82),
    'skills': (r'skills?|technical\s*skills|core\s*competenc\w*|tech\s*stack|technologies|programming\s*languages|tools|frameworks|libraries|competencies|expertise|software\s+proficien\w*|proficient\s*in', 90),
    'personal.name': (r'full\s*name|your\s*name|candidate\s*name|applicant\s*name|legal\s*name|preferred\s*name|middle\s*name|\bname\b', 20),
}

NEGATIVE_RULES = {
    'personal.name': (r'company\s+name|organisation\s+name|organization\s+name|university\s+name|college\s+name|school\s+name|institution\s+name|project\s+name|file\s+name|user\s*name|team\s+name|account\s+name|product\s+name|full\s+name\s+of\s+(?:the|this)\s+(?:company|school|university|organisation|organization)',),
    'personal.location': (r'work\s*location\s*(?:of|:)|job\s*location\s*of',),
    'personal.summary': (r'web\s*summary|salary\s*summary|summary\s*of\s*(?:the|this)\s+(?:form|application|role|job)',),
    'personal.date_of_birth': (r'start\s*date|end\s*date|joining\s*date|leaving\s*date|graduation|available\s+to\s+start',),
    'personal.availability': (r'(?:employment|job|work|education|academic|school|university|college|project)\s*start\s*date|start\s*date\s*(?:of|from)\s*(?:employment|job|work|education)|end\s*date',),
    'skills': (r'skill\s*level|skills?\s*gap|gap\s*in\s*skills',),
    'education.degree': (r'degree\s*of\s+(?:the\s+)?(?:problem|error|change)',),
    'experience.company': (r'company\s+name\s+of\s+previous',),
    'experience.start_date': (r'(?:education|academic|school|university|college)\s*start|available\s+to\s+start|earliest\s+start',),
    'experience.end_date': (r'(?:education|academic|school|university|college|graduation)\s*end|date\s+of\s+graduation',),
}

FIELD_HINTS = {
    'personal.name': ('full name', 'your name', 'candidate name', 'applicant name', 'legal name', 'name', 'fullname'),
    'personal.first_name': ('first name', 'given name', 'forename', 'firstname'),
    'personal.last_name': ('last name', 'family name', 'surname', 'lastname'),
    'personal.email': ('email', 'e mail', 'email address', 'contact email', 'work email', 'email id'),
    'personal.phone': ('phone', 'mobile', 'telephone', 'cell phone', 'phone number', 'telephone number', 'contact number', 'contact phone', 'primary phone'),
    'personal.location': ('location', 'city', 'country', 'state province', 'current location', 'based in', 'residence'),
    'personal.address': ('street address', 'home address', 'full address', 'mailing address', 'postal address', 'address line', 'address'),
    'personal.date_of_birth': ('date of birth', 'dob', 'birthday', 'birth date'),
    'personal.linkedin': ('linkedin', 'linked in'),
    'personal.github': ('github', 'git hub'),
    'personal.portfolio': ('portfolio', 'personal website', 'personal site', 'behance', 'dribbble'),
    'personal.headline': ('current title', 'current job title', 'professional title', 'desired job title', 'target role', 'headline', 'occupation'),
    'personal.summary': ('professional summary', 'career summary', 'about me', 'about yourself', 'tell us about yourself', 'tell me about yourself', 'introduce yourself', 'professional background', 'career overview', 'short bio', 'professional profile', 'career objective'),
    'personal.languages': ('languages', 'spoken language'),
    'personal.current_company': ('current company', 'current employer', 'present employer'),
    'personal.graduation_year': ('graduation year', 'year of graduation'),
    'personal.work_authorization': ('work authorization', 'work authorisation', 'authorized to work', 'eligible to work', 'right to work'),
    'personal.visa_status': ('visa', 'work permit', 'immigration status', 'sponsorship'),
    'personal.availability': ('availability', 'available to start', 'when can you start', 'date you can start', 'relocation'),
    'personal.salary_expectation': ('salary', 'compensation', 'pay expectation', 'expected salary', 'desired salary'),
    'personal.references': ('reference', 'references', 'referee'),
    'personal.certifications': ('certification', 'certifications', 'licenses', 'accreditations'),
    'personal.awards': ('awards', 'honours', 'honors', 'achievements', 'accomplishments'),
    'personal.interests': ('hobbies', 'interests', 'leisure activities', 'outside of work'),
    'education.institution': ('institution', 'university', 'college', 'school', 'institute', 'academy', 'alma mater', 'higher education', 'educational institution'),
    'education.description': ('educational background', 'academic background', 'describe your education', 'tell us about your education'),
    'education.degree': ('degree', 'qualification', 'diploma', 'bachelor', 'bachelors', 'master', 'masters', 'phd', 'mba'),
    'education.field': ('major', 'field of study', 'specialization', 'specialisation', 'discipline', 'course of study'),
    'education.grade': ('grade', 'gpa', 'cgpa', 'marks'),
    'education.start_date': ('start date', 'start year', 'from year', 'education start'),
    'education.end_date': ('graduation date', 'date of graduation', 'end date', 'end year', 'completion date'),
    'experience.company': ('company', 'employer', 'organisation', 'organization', 'workplace', 'firm'),
    'experience.position': ('job title', 'role', 'position', 'designation', 'title'),
    'experience.description': ('job description', 'role description', 'responsibilities', 'duties', 'describe your work experience', 'tell us about your experience', 'tell us about your work experience'),
    'experience.start_date': ('start date', 'joining date'),
    'experience.end_date': ('end date', 'leaving date'),
    'projects.name': ('project name', 'project', 'portfolio project', 'case study'),
    'projects.description': ('project description', 'project summary', 'describe the project', 'tell us about your project'),
    'projects.technologies': ('technologies used', 'tools used', 'built with', 'tech used'),
    'projects.role': ('your role in the project', 'role in the project', 'contribution'),
    'projects.url': ('project url', 'project link', 'live link', 'demo link', 'repository url'),
    'skills': ('skills', 'technical skills', 'core competencies', 'tech stack', 'technologies', 'tools', 'frameworks', 'languages', 'expertise'),
}

AUTOCOMPLETE_FIELDS = {
    'name': 'personal.name', 'given-name': 'personal.first_name', 'additional-name': 'personal.name',
    'family-name': 'personal.last_name', 'honorific-prefix': None, 'nickname': None,
    'email': 'personal.email', 'tel': 'personal.phone', 'tel-national': 'personal.phone',
    'url': 'personal.portfolio', 'photo': None, 'bday': 'personal.date_of_birth',
    'street-address': 'personal.address', 'address-line1': 'personal.address', 'address-line2': 'personal.address',
    'address-level1': 'personal.location', 'address-level2': 'personal.location',
    'address-level3': 'personal.location', 'address-level4': 'personal.location',
    'country': 'personal.location', 'country-name': 'personal.location', 'postal-code': 'personal.address',
    'organization': 'experience.company', 'organization-title': 'experience.position', 'organization-unit': 'experience.position',
    'current-password': None, 'new-password': None, 'one-time-code': None, 'cc-name': None, 'cc-csc': None, 'transaction-currency': None,
}

TYPE_FIELDS = {
    'email': 'personal.email', 'tel': 'personal.phone', 'url': 'personal.portfolio',
    'date': 'personal.date_of_birth', 'month': 'personal.graduation_year',
    'datetime-local': 'personal.availability',
}

STRUCTURED_LONG_TEXT = {
    'personal.address', 'personal.summary', 'personal.references', 'personal.interests',
    'experience.description', 'projects.description', 'education.description', 'skills',
}

SECTION_GROUPS = {
    'education': ('education', 'academic', 'university', 'college', 'school', 'qualification', 'degree'),
    'experience': ('employment', 'work experience', 'work history', 'professional experience',
                   'employment history', 'current employment', 'job', 'career', 'previous role'),
    'projects': ('project', 'portfolio', 'personal project', 'academic project'),
    'skills': ('skill', 'technical', 'competenc', 'expertise'),
}

# --------------------------------------------------------------------------------------
# Question handling
# --------------------------------------------------------------------------------------

OPEN_ENDED = re.compile(
    r'\b('
    r'why\b|describe\b|explain\b|elaborate\b|'
    r'tell\s+(?:us|me)(?:\s+about)?|talk\s+(?:to\s+us\s+about|about)|'
    r'goals?\b|hiring\s+us|motivat\w*|interest(?:ed|s)?\b|strengths?\b|weakness\w*|'
    r'challenge\w*|achievement\w*|accomplish\w*|contribut\w*|good\s+fit|fit\s+for|'
    r'your\s+approach|approach\s+to|how\s+(?:would|do|did|have|has|could|might)\s+you|'
    r'what\s+(?:would|do|did|have|has|could|might|makes|made|is|are)\s+you|'
    r'when\s+have\s+you|describe\s+(?:a\s+)?(?:time|situation|project|experience)|'
    r'what\s+was\s+your|what\s+are\s+your\s+(?:plans|goals|strengths|weaknesses)|'
    r'share\s+(?:an\s+)?example|walk\s+us|give\s+(?:an?\s+)?example|cover\s+letter|'
    r'personal\s+statement|introduce\s+yourself|about\s+yourself|what\s+can\s+we\s+know\s+about\s+you|'
    r'what\s+are\s+you\s+looking\s+for|tell\s+us\s+about\s+your|'
    r'additional\s+information|anything\s+else\s+you\s+would\s+like|'
    r'why\s+should\s+we|what\s+sets\s+you\s+apart|how\s+do\s+you\s+handle'
    r')', re.I)

QUESTION_GUIDANCE = [
    ('cover_letter', re.compile(r'cover\s+letter|letter\s+of\s+application|supporting\s+statement', re.I),
     'Open by naming the role and why the work interests you, then support that with the two or three strongest pieces of evidence, then close on what you would contribute.',
     ['headline', 'skills', 'experience', 'projects', 'education', 'profile', 'previous_answers', 'cv']),
    ('introduction', re.compile(r'tell\s+us\s+about\s+yourself|introduce\s+yourself|about\s+you\b|who\s+are\s+you|summary\s+of\s+your|your\s+background|describe\s+yourself', re.I),
     'Give a short professional summary: who you are professionally, your current focus, and the experience and skills you would bring to this role.',
     ['headline', 'experience', 'skills', 'education', 'projects', 'profile', 'previous_answers', 'cv']),
    ('motivation', re.compile(r'why\b|motivat\w*|attracted|interested?\s+in|want\s+to\s+work|apply|join\s+(?:us|your)|hiring\s+us|good\s+fit|fit\s+for|inspire|excited|passion', re.I),
     'Answer why you want this work specifically, tying the reason directly to your actual experience, projects and skills. Do not restate your whole CV.',
     ['headline', 'experience', 'projects', 'skills', 'previous_answers', 'profile', 'cv', 'education']),
    ('weakness', re.compile(r'weakness|weak\s+point|areas?\s+to\s+improve|to\s+improve|improve\s+your|gap\s+in|development\s+area|not\s+great\s+at|limitation', re.I),
     'Name the limitation plainly, then explain the concrete step you have taken to work on it. Keep it honest and short, and end on the current state rather than the problem.',
     ['skills', 'experience', 'previous_answers', 'headline', 'profile', 'cv', 'education', 'projects']),
    ('strength', re.compile(r'strength|best\s+at|strongest|what\s+(?:are|is)\s+your\s+best|what\s+makes\s+you\s+(?:good|strong)', re.I),
     'Name one or two concrete strengths and prove each one with a specific piece of evidence from your experience or projects.',
     ['skills', 'experience', 'projects', 'previous_answers', 'headline', 'profile', 'cv', 'education']),
    ('challenge', re.compile(r'challenge|difficult|hardest|tough|failure|failed|mistake|setback|problem\s*[- ]?solving|conflict|crisis|deadline', re.I),
     'Describe the situation, what you personally did about it, and the result, in that order. Keep the example specific and do not blame other people.',
     ['experience', 'projects', 'previous_answers', 'headline', 'skills', 'profile', 'cv', 'education']),
    ('leadership', re.compile(r'leader|team\s*work|teamwork|\bteams?\b|collabo?rat\w*|manage|mentor|communicat\w*|delegat\w*|supervis\w*|influenc\w*', re.I),
     'Show how you worked with other people: what you led or coordinated, how you communicated, and the outcome you produced.',
     ['experience', 'projects', 'previous_answers', 'skills', 'headline', 'profile', 'cv', 'education']),
    ('experience', re.compile(r'experience|background|previous|accomplish|achievement\w*|responsibilit\w*|contribut\w*|employment|job\s+history|what\s+have\s+you\s+done|where\s+did\s+you\s+work', re.I),
     'Describe the relevant roles you have held, what you were responsible for, and the results you produced. Use the specific roles and employers from the evidence.',
     ['experience', 'projects', 'previous_answers', 'headline', 'skills', 'profile', 'cv', 'education']),
    ('project', re.compile(r'project|portfolio|built|develop\w*|creat\w*|research|thesis|application\s+you\s+built|most\s+ambitious', re.I),
     'Describe the project you built: the problem it solved, what you personally did, which technologies you used, and the result. Name the project as it appears in the evidence.',
     ['projects', 'experience', 'previous_answers', 'skills', 'profile', 'cv', 'education']),
    ('skills', re.compile(r'skill|technolog\w*|tool|stack|program\w*|proficien\w*|competenc\w*|software|expertise|certificat\w*|licen[cs]e\w*', re.I),
     'State the specific skills and tools you have and the level of hands-on experience you have with each. Name them exactly as they appear in the evidence.',
     ['skills', 'projects', 'experience', 'previous_answers', 'education', 'profile', 'cv']),
    ('education', re.compile(r'educat\w*|universit\w*|college|school|degree|studied|study|academic|graduat\w*|qualification|curriculum|course', re.I),
     'Answer about your education only: what you studied, where, when, and what it gave you. Do not drift into work experience.',
     ['education', 'skills', 'previous_answers', 'profile', 'cv', 'experience', 'projects']),
    ('availability', re.compile(r'availab\w*|start\s+date|when\s+can\s+you\s+start|notice\s+period|relocat\w*|open\s+to|earliest', re.I),
     'Answer the practical question directly and only. State the notice period or start date you have evidence for and any relocation constraint.',
     ['availability', 'headline', 'previous_answers', 'profile', 'experience', 'cv']),
    ('compensation', re.compile(r'salary|compensation|remuneration|pay\s*expectation|expected\s*salary|desired\s*salary|salary\s+range', re.I),
     'Answer the compensation question directly with the figure or range you have evidence for, in one or two sentences.',
     ['compensation', 'headline', 'previous_answers', 'profile', 'experience', 'cv']),
    ('research', re.compile(r'publication|paper|journal|research|thesis|ph\.?d|patent|citation', re.I),
     'Answer about your research or published work only, naming the specific work and your contribution to it.',
     ['education', 'experience', 'projects', 'previous_answers', 'profile', 'cv']),
]

DEFAULT_GUIDANCE = (
    'Answer the exact question that was asked. Use only the evidence that is relevant to that question and ignore the rest. Do not summarise the whole profile.',
    ['headline', 'experience', 'skills', 'projects', 'previous_answers', 'profile', 'education', 'cv'],
)

CV_HEADINGS = {
    'education': {'education', 'academic background', 'academic history', 'qualifications', 'academic qualifications', 'education background'},
    'experience': {'experience', 'work experience', 'professional experience', 'employment history', 'work history', 'professional background', 'career history'},
    'skills': {'skills', 'technical skills', 'core skills', 'technologies', 'technical expertise', 'competencies', 'key skills', 'it skills', 'tools'},
    'projects': {'projects', 'personal projects', 'selected projects', 'key projects', 'academic projects', 'portfolio'},
    'certifications': {'certifications', 'certificates', 'licenses', 'training', 'courses', 'awards and certifications'},
    'summary': {'summary', 'professional summary', 'profile', 'objective', 'about me', 'career summary', 'personal profile', 'personal details'},
}

CV_STOP_WORDS = {
    'a', 'about', 'above', 'after', 'all', 'also', 'am', 'an', 'and', 'any', 'are', 'as', 'at', 'be', 'because', 'been', 'being', 'but', 'by',
    'can', 'could', 'did', 'do', 'does', 'doing', 'done', 'during', 'each', 'etc', 'for', 'from', 'get', 'had', 'has', 'have', 'having', 'he',
    'her', 'here', 'hers', 'him', 'his', 'how', 'i', 'if', 'in', 'into', 'is', 'it', 'its', 'just', 'me', 'might', 'more', 'most', 'my',
    'no', 'not', 'of', 'off', 'on', 'once', 'one', 'only', 'or', 'other', 'our', 'ours', 'out', 'over', 'own', 'please', 'she', 'should',
    'so', 'some', 'such', 'than', 'that', 'the', 'their', 'theirs', 'them', 'then', 'there', 'these', 'they', 'this', 'those', 'through',
    'to', 'too', 'under', 'until', 'up', 'us', 'very', 'was', 'we', 'were', 'what', 'when', 'where', 'which', 'while', 'who', 'whom', 'why',
    'will', 'with', 'would', 'you', 'your', 'yours', 'role', 'position', 'experience', 'describe', 'tell', 'want', 'like', 'well', 'able',
    'give', 'need', 'work', 'working', 'company', 'team', 'please select',
}

# Deliberately narrower than the chunking list: words like "role", "work" and "team" are the
# subject of many questions, so they must survive when judging whether an answer is on topic.
FUNCTION_WORDS = {
    'a', 'an', 'and', 'are', 'as', 'at', 'be', 'been', 'being', 'but', 'by', 'can', 'could', 'did', 'do', 'does', 'for', 'from',
    'had', 'has', 'have', 'he', 'her', 'here', 'hers', 'him', 'his', 'how', 'i', 'if', 'in', 'into', 'is', 'it', 'its', 'me', 'my',
    'no', 'not', 'of', 'on', 'or', 'our', 'ours', 'she', 'should', 'so', 'some', 'such', 'than', 'that', 'the', 'their', 'them', 'then',
    'there', 'these', 'they', 'this', 'those', 'to', 'us', 'was', 'we', 'were', 'what', 'when', 'where', 'which', 'who', 'whom', 'why',
    'will', 'with', 'would', 'you', 'your', 'tell', 'us', 'please', 'describe', 'give', 'share', 'example', 'also', 'many', 'much',
    'most', 'other', 'own', 'too', 'very', 'about', 'above', 'after', 'again', 'because', 'been', 'both', 'each', 'few', 'more',
    'once', 'only', 'same', 'just', 'like', 'made', 'make', 'over', 'under', 'until', 'up', 'out', 'down', 'off', 'am',
}

CONTROL_PREFIX = re.compile(
    r'^(?:[a-z]{1,4}[_\-.])*(?:input|edit|field|box|boxlabel|labelfield|text|fieldlabel|flabel|item)\d*[_:\-. ]*', re.I)
CONTROL_TOKEN = re.compile(r'(?:^|[_\-. ])(?:ctl|ctrl|spn|sp|ddl|cmb|txt|tb|hdn|hid|grd|row|col|cont|container|holder|placeholder)(?:[_\-. ]|$)', re.I)
NOISE_ALIAS = re.compile(
    r'^(?:yes|no|none|null|undefined|true|false|control|default|auto|anonymous|'
    r'element|node|component|widget|gform|js|jquery|ui|userdefinedinput|form|'
    r'element0|input\d*|item\d*|field\d*|ctl\d*|placeholder)$',
    re.I,
)


# --------------------------------------------------------------------------------------
# Text helpers
# --------------------------------------------------------------------------------------

def clean_cv_text(text):
    text = text.replace('\u00a0', ' ').replace('\u00ad', '')
    text = re.sub(r'(?<=[A-Za-z])-\s*\n\s*(?=[a-z])', '', text)
    lines = [re.sub(r'[ \t]+', ' ', line).strip() for line in text.splitlines()]
    cleaned = '\n'.join(lines)
    return re.sub(r'\n{3,}', '\n\n', cleaned).strip()


def cv_sections(text):
    sections = {name: [] for name in CV_HEADINGS}
    aliases = {re.sub(r'[^a-z]', '', alias): name for name, values in CV_HEADINGS.items() for alias in values}
    current = None
    content = []

    def save():
        value = ' '.join(content).strip()
        if current and value:
            sections[current].append(value[:4000])

    for line in text.splitlines():
        heading = re.sub(r'[^a-z]', '', line.lower().strip(' :.-'))
        matched = aliases.get(heading) if 0 < len(heading) < 64 else None
        if matched:
            save(); current = matched; content = []
        elif line.strip():
            content.append(line.strip())
    save()
    return sections


def extract_cv_text(content, suffix):
    if suffix == '.pdf':
        with fitz.open(stream=content, filetype='pdf') as document:
            text = '\n\n'.join(page.get_text('text', sort=True) for page in document)
    elif suffix == '.docx':
        document = Document(BytesIO(content))
        blocks = []
        for child in document.element.body.iterchildren():
            if isinstance(child, CT_P):
                value = Paragraph(child, document).text.strip()
                if value:
                    blocks.append(value)
            elif isinstance(child, CT_Tbl):
                table = Table(child, document)
                for row in table.rows:
                    values = list(dict.fromkeys(cell.text.strip() for cell in row.cells if cell.text.strip()))
                    if values:
                        blocks.append(' | '.join(values))
        text = '\n'.join(blocks)
    else:
        raise ValueError('Unsupported CV format.')
    text = clean_cv_text(text)
    return text, cv_sections(text)


def _split_identifier(value):
    spaced = re.sub(r'[_\-.\[\]]+', ' ', str(value or ''))
    spaced = re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', spaced)
    spaced = re.sub(r'(?<=[A-Z])(?=[A-Z][a-z])', ' ', spaced)
    spaced = CONTROL_PREFIX.sub('', spaced)
    spaced = CONTROL_TOKEN.sub(' ', spaced)
    spaced = re.sub(r'\d+', ' ', spaced)
    words = [word for word in spaced.split() if len(word) > 1 and not NOISE_ALIAS.match(word)]
    return ' '.join(words)


def normalise_text(value):
    spaced = re.sub(r'[_\-]+', ' ', str(value or ''))
    spaced = re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', spaced)
    spaced = re.sub(r'(?<=[A-Z])(?=[A-Z][a-z])', ' ', spaced)
    return re.sub(r'\s+', ' ', spaced).lower().strip()


def _attr(field, name, default=''):
    return getattr(field, name, default) or default


def primary_text(field):
    """Signal-bearing text: what the field actually calls itself.

    Prefer the visible label/aria/placeholder when present so noisy framework
    name/id attributes cannot override an exact on-page label.
    """
    visible = [
        _attr(field, 'label'), _attr(field, 'question_hint'), _attr(field, 'question'), _attr(field, 'placeholder'),
        _attr(field, 'aria_label'), _attr(field, 'autocomplete').replace('-', ' '),
    ]
    # The scanner may provide the same label as both `label` and
    # `question_hint`; dedupe those values so generic labels stay identifiable.
    unique_visible = list(dict.fromkeys(normalise_text(p) for p in visible if normalise_text(p)))
    visible_text = ' '.join(unique_visible)
    if len(visible_text) >= 3:
        return visible_text
    parts = [
        _split_identifier(_attr(field, 'name')),
        _split_identifier(_attr(field, 'id_attribute')),
        _attr(field, 'autocomplete').replace('-', ' '),
    ]
    return normalise_text(' '.join(p for p in parts if p))


def identifier_text(field):
    """Names and ids are weaker than a real label, but often rescue generic labels."""
    return normalise_text(' '.join(filter(None, [
        _split_identifier(_attr(field, 'name')),
        _split_identifier(_attr(field, 'id_attribute')),
    ])))


def secondary_text(field):
    """Surrounding context, used only when the field names itself poorly."""
    return normalise_text(' '.join(filter(None, [_attr(field, 'section'), _attr(field, 'nearby_text')])))


def field_text(field):
    """Weighted view of a field used by the regex and hint layers."""
    weighted = ' ; '.join(filter(None, [
        _attr(field, 'question'), _attr(field, 'label'), _attr(field, 'question_hint'), _attr(field, 'placeholder'),
        _attr(field, 'aria_label'), _attr(field, 'autocomplete').replace('-', ' '),
        _split_identifier(_attr(field, 'name')), _split_identifier(_attr(field, 'id_attribute')),
        _attr(field, 'section'), _attr(field, 'nearby_text')[:280],
    ]))
    return normalise_text(weighted)


def ml_text(field):
    parts = [
        _attr(field, 'label'), _attr(field, 'question_hint'), _attr(field, 'question'), _attr(field, 'placeholder'),
        _attr(field, 'aria_label'), _attr(field, 'autocomplete').replace('-', ' '),
        _split_identifier(_attr(field, 'name')), _split_identifier(_attr(field, 'id_attribute')),
        _attr(field, 'section'),
    ]
    seen = set()
    ordered = []
    for part in parts:
        for word in normalise_text(part).split():
            if word not in seen:
                seen.add(word); ordered.append(word)
    if not _attr(field, 'nearby_text'):
        return ' '.join(ordered)
    return ' '.join(ordered + normalise_text(_attr(field, 'nearby_text'))[:400].split())


def question_of(field):
    for source in (_attr(field, 'question'), _attr(field, 'label'), _attr(field, 'question_hint'),
                   _attr(field, 'placeholder'), _attr(field, 'aria_label'), _attr(field, 'section'),
                   _attr(field, 'nearby_text')[:240]):
        value = re.sub(r'\s+', ' ', str(source or '')).strip()
        if len(value) >= 3:
            return value[:600]
    return ''


def _hint_score(hint, text):
    pattern = r'(?<![a-z0-9])' + re.escape(hint).replace(r'\ ', r'\s+') + r'(?![a-z0-9])'
    match = re.search(pattern, text)
    if not match:
        return None
    return len(match.group(0)), -match.start()


# --------------------------------------------------------------------------------------
# Classification
# --------------------------------------------------------------------------------------

def _section_group(text):
    best, best_length = None, 0
    for group, keywords in SECTION_GROUPS.items():
        for keyword in keywords:
            if keyword in text and len(keyword) > best_length:
                best, best_length = group, len(keyword)
    return best


class FieldClassifier:
    def __init__(self):
        self.vectorizer = self.model = None
        try:
            self.vectorizer = joblib.load(settings.model_dir / 'tfidf_vectorizer.joblib')
            self.model = joblib.load(settings.model_dir / 'field_classifier.joblib')
        except Exception:
            pass

    def _blocked(self, field, text):
        return any(re.search(pattern, text) for pattern in NEGATIVE_RULES.get(field, ()))

    def _adjust(self, field, group, score):
        """Fields inside a labelled section beat generic personal-data matches."""
        if not group:
            return score
        collection = field.split('.', 1)[0] if '.' in field else ''
        if collection == group:
            return score + 32
        # A bare "Name" in an education/work/project section names that
        # record, not the applicant. Explicit first/last-name labels stay safe.
        if field == 'personal.name' and group in {'education', 'experience', 'projects'}:
            return score - 45
        if collection in {'education', 'experience', 'projects'}:
            return score - 22
        if field in {'personal.availability', 'personal.notice_period'} and group == 'experience':
            return score - 18
        return score

    def _rule_candidates(self, text, group):
        candidates = []
        for field, (pattern, priority) in RULES.items():
            if self._blocked(field, text):
                continue
            match = re.search(pattern, text)
            if not match:
                continue
            covered = len(match.group(0))
            candidates.append((self._adjust(field, group, priority + min(covered, 30) / 100), covered, field))
        return sorted(candidates, reverse=True)

    def _hint_candidates(self, text, group):
        candidates = []
        for field, hints in FIELD_HINTS.items():
            if self._blocked(field, text):
                continue
            best = None
            for hint in hints:
                score = _hint_score(hint, text)
                if score and (best is None or score[0] > best[0]):
                    best = score
            if best is None:
                continue
            priority = RULES.get(field, (None, 50))[1]
            candidates.append((self._adjust(field, group, priority + min(best[0], 30) / 100), best[0], field))
        return sorted(candidates, reverse=True)

    def _type_guess(self, field, group, primary):
        """Map HTML input types, but do not force DOB onto employment/education dates."""
        kind = _attr(field, 'type').strip().lower()
        if kind in {'date', 'month', 'datetime-local'}:
            if re.search(r'\b(start|from|joining)\b', primary):
                if group == 'education':
                    return 'education.start_date'
                return 'experience.start_date'
            if re.search(r'\b(end|to|leaving|until|graduation|completion)\b', primary):
                if group == 'education' or re.search(r'graduat|completion|education|academic', primary):
                    return 'education.end_date'
                return 'experience.end_date'
            if group in {'education', 'experience', 'projects'}:
                return None
            if kind == 'month' and re.search(r'graduat', primary + ' ' + (group or '')):
                return 'personal.graduation_year'
        return TYPE_FIELDS.get(kind)

    def predict(self, field):
        """Return (mapped_field, confidence, source)."""
        primary = primary_text(field)
        identifiers = identifier_text(field)
        context = secondary_text(field)
        if not primary and not context:
            type_field = self._type_guess(field, None, primary)
            if type_field:
                return type_field, 0.97, 'type'
            return None, 0.0, 'none'
        group = _section_group(context)

        auto = AUTOCOMPLETE_FIELDS.get(_attr(field, 'autocomplete').strip().lower())
        if auto:
            return auto, 0.99, 'autocomplete'

        # Generic "Name" labels inherit the meaning of their form section.
        if primary in {'name', 'name of'}:
            contextual_field = {
                'education': 'education.institution',
                'experience': 'experience.company',
                'projects': 'projects.name',
            }.get(group)
            if contextual_field:
                return contextual_field, 0.94, 'section'
            if identifiers:
                identifier_hints = self._hint_candidates(identifiers, group)
                if identifier_hints and identifier_hints[0][2] in {
                    'education.institution', 'experience.company', 'projects.name'
                }:
                    return identifier_hints[0][2], 0.92, 'identifier'

        hint_pool = self._hint_candidates(primary, group)
        if hint_pool:
            best = hint_pool[0]
            runner_up = hint_pool[1][2] if len(hint_pool) > 1 else None
            confidence = 0.985 if runner_up != best[2] else 0.95
            if _attr(field, 'required'):
                confidence = min(0.995, confidence + 0.005)
            return best[2], confidence, 'hint'

        rule_pool = self._rule_candidates(primary, group)
        if rule_pool:
            return rule_pool[0][2], 0.98, 'rule'

        type_field = self._type_guess(field, group, primary)
        if type_field:
            return type_field, 0.97, 'type'

        # Framework generated ids often carry the only useful signal when the
        # visible label is a generic placeholder such as "Enter value".
        if identifiers:
            identifier_hints = self._hint_candidates(identifiers, group)
            if identifier_hints:
                return identifier_hints[0][2], 0.90, 'identifier'
            identifier_rules = self._rule_candidates(identifiers, group)
            if identifier_rules:
                return identifier_rules[0][2], 0.88, 'identifier'

        if self.model is not None:
            sample = ml_text(field)
            if sample:
                probabilities = self.model.predict_proba(self.vectorizer.transform([sample]))[0]
                index = probabilities.argmax()
                field_name = self.model.classes_[index]
                confidence = float(probabilities[index])
                if field_name != 'other' and confidence >= 0.60:
                    return field_name, confidence, 'ml'

        secondary_hints = self._hint_candidates(context, group) if context else []
        secondary_rules = self._rule_candidates(context, group) if context else []
        combined = secondary_hints[:1] + secondary_rules[:1]
        if combined:
            return max(combined)[2], 0.72, 'context'

        fallback = self._type_guess(field, group, primary)
        if fallback:
            return fallback, 0.6, 'type'
        return None, 0.0, 'none'


classifier = FieldClassifier()


def detect_open_ended(field):
    """Decide whether a field needs a written answer rather than a stored value."""
    kind = _attr(field, 'type').lower()
    if kind in {'email', 'tel', 'url', 'date', 'month', 'datetime-local', 'number', 'select-one', 'select-multiple'}:
        return False, ''
    # Decide from the field's own prompt, not nearby section text or generated
    # DOM ids. Those can contain unrelated question words from adjacent fields.
    text = normalise_text(' '.join(filter(None, (
        _attr(field, 'question'), _attr(field, 'label'), _attr(field, 'question_hint'),
        _attr(field, 'placeholder'), _attr(field, 'aria_label'),
    ))))
    multiline = _attr(field, 'tag').lower() == 'textarea' or _attr(field, 'type').lower() == 'contenteditable'
    if OPEN_ENDED.search(text) or (multiline and '?' in text):
        return True, question_of(field)
    if not multiline:
        return False, ''
    mapped, confidence, _ = classifier.predict(field)
    if mapped in STRUCTURED_LONG_TEXT and confidence >= 0.9:
        return False, ''
    return True, question_of(field)


# --------------------------------------------------------------------------------------
# Values
# --------------------------------------------------------------------------------------

def _split_name(full_name):
    parts = [part for part in str(full_name or '').split() if part]
    if not parts:
        return '', ''
    if len(parts) == 1:
        return parts[0], ''
    return ' '.join(parts[:-1]), parts[-1]


def _first_row(db, model):
    return db.scalars(select(model).order_by(model.id)).first()


def get_value(db: Session, field: str):
    if not field:
        return ''
    profile = db.scalars(select(Profile).order_by(Profile.id)).first()
    if field in CANONICAL_FIELDS:
        column = CANONICAL_FIELDS[field]
        if column == FIRST_NAME:
            return _split_name(getattr(profile, 'full_name', '') if profile else '')[0]
        if column == LAST_NAME:
            return _split_name(getattr(profile, 'full_name', '') if profile else '')[1]
        return (getattr(profile, column, '') or '') if profile else ''
    if field == 'skills':
        return ', '.join(skill.name for skill in db.scalars(select(Skill).order_by(Skill.name)).all())
    if field in COLLECTION_FIELDS:
        model, attribute = COLLECTION_FIELDS[field]
        if field == 'education.end_date' and not _first_row_value(db, Education, 'end_date'):
            return (getattr(profile, 'graduation_year', '') or '') if profile else ''
        if field == 'experience.company' and not _first_row_value(db, Experience, 'company'):
            return (getattr(profile, 'current_company', '') or '') if profile else ''
        return _first_row_value(db, model, attribute)
    return ''


def _first_row_value(db, model, attribute):
    row = _first_row(db, model)
    if row is None or not hasattr(row, attribute):
        return ''
    return getattr(row, attribute, '') or ''


IGNORED_OPTION_TOKENS = {'please', 'select', 'choose', 'option', 'other', 'the', 'a', 'an', 'one', 'n', 'not', 'applicable', 'na', ''}


def _normalise_option(value):
    return re.sub(r'[^a-z0-9]+', ' ', str(value or '').lower()).strip()


def _option_tokens(value):
    return [token for token in _normalise_option(value).split() if token not in IGNORED_OPTION_TOKENS]


def option_similarity(target, option):
    left = _normalise_option(target)
    right = _normalise_option(option)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if right in left or left in right:
        return 0.9
    left_tokens = _option_tokens(target)
    right_tokens = _option_tokens(option)
    if not left_tokens or not right_tokens:
        return 0.0
    hits = 0
    for word in left_tokens:
        if any(candidate == word or (len(word) > 3 and candidate.startswith(word)) or (len(candidate) > 3 and word.startswith(candidate)) for candidate in right_tokens):
            hits += 1
    return hits / max(len(left_tokens), len(right_tokens))


def resolve_value(db: Session, field: str, options=None, option_values=None):
    """Return (display value, matching option value or None) for a dropdown."""
    value = get_value(db, field)
    if not value or not options:
        return value, None
    best_index, best_score = None, 0.0
    for index, option in enumerate(options):
        score = max(option_similarity(value, option),
                    option_similarity(value, option_values[index]) if option_values and index < len(option_values) else 0.0)
        if score > best_score:
            best_score, best_index = score, index
    if best_index is None or best_score < 0.34:
        return value, None
    return options[best_index], (option_values[best_index] if option_values and best_index < len(option_values) else None)


# --------------------------------------------------------------------------------------
# Answer context
# --------------------------------------------------------------------------------------

def _stem(word):
    for suffix in ('ingly', 'edly', 'ing', 'ies', 'ied', 'ed', 'es', 's'):
        if len(word) > 4 and word.endswith(suffix):
            return word[:-len(suffix)]
    return word


def _content_words(text, stop_words=CV_STOP_WORDS):
    return {
        _stem(word) for word in re.findall(r"[a-z][a-z0-9+#]{2,}", str(text or '').lower())
        if word not in stop_words
    }


def _question_terms(question):
    return _content_words(question, FUNCTION_WORDS)


def _profile_row(db):
    return db.scalars(select(Profile).order_by(Profile.id)).first()


def _profile_block(db):
    profile = _profile_row(db)
    if profile is None:
        return ''
    labels = [
        ('Name', 'full_name'), ('Email', 'email'), ('Phone', 'phone'), ('Location', 'location'),
        ('Address', 'address'), ('Date of birth', 'date_of_birth'),
        ('LinkedIn', 'linkedin'), ('GitHub', 'github'), ('Website or portfolio', 'portfolio'),
        ('Current role', 'headline'), ('Current company', 'current_company'),
        ('Graduation year', 'graduation_year'), ('Languages', 'languages'),
        ('Work authorisation', 'work_authorization'), ('Visa status', 'visa_status'),
        ('Availability', 'availability'), ('Notice period', 'notice_period'),
        ('Expected salary', 'salary_expectation'), ('References', 'references'),
        ('Certifications', 'certifications'), ('Awards', 'awards'), ('Interests', 'interests'),
    ]
    lines = [f'{label}: {getattr(profile, column, "")}'.strip() for label, column in labels if getattr(profile, column, '')]
    if getattr(profile, 'summary', ''):
        lines.append(f'Professional summary: {profile.summary.strip()}')
    return '\n'.join(lines)


def _skills_block(db):
    skills = [skill.name for skill in db.scalars(select(Skill).order_by(Skill.name)).all()]
    return ', '.join(skills)


def _row_text(row):
    return ' '.join(
        str(getattr(row, column.name, '') or '')
        for column in row.__table__.columns
        if column.name not in {'id', 'created_at', 'updated_at'} and getattr(row, column.name, '')
    )


def _date_range(start, end):
    start, end = (start or '').strip(), (end or '').strip()
    if not start and not end:
        return ''
    return f'({start} to {end or "present"})' if start or end else ''


def _recency_key(start, end):
    """Sort key that treats an open-ended role or course as the most recent one."""
    def rank(value):
        numbers = re.findall(r'\d{4}', str(value or ''))
        return int(numbers[0]) if numbers else 0
    return (rank(end) or rank(start), rank(start))


def _experience_block(db, terms, limit=4):
    rows = list(db.scalars(select(Experience).order_by(Experience.start_date.desc(), Experience.id)).all())
    if not rows:
        return ''
    def score(row):
        words = _content_words(_row_text(row))
        return len(terms & words) / (len(words) ** 0.5) if words else 0.0
    ranked = sorted(rows, key=lambda row: (score(row), row.id), reverse=True)[:limit]
    ranked.sort(key=lambda row: _recency_key(row.start_date, row.end_date), reverse=True)
    lines = []
    for row in ranked:
        headline = ' — '.join(part for part in [row.position, row.company] if part) or 'Role'
        period = _date_range(row.start_date, row.end_date)
        entry = f'{headline}{" " + period if period else ""}'
        if row.description:
            entry += f'\n   {row.description.strip()}'
        lines.append(entry)
    return '\n'.join(lines)


def _education_block(db, terms, limit=3):
    rows = list(db.scalars(select(Education).order_by(Education.end_date.desc(), Education.id)).all())
    if not rows:
        return ''
    def score(row):
        words = _content_words(_row_text(row))
        return len(terms & words) / (len(words) ** 0.5) if words else 0.0
    ranked = sorted(rows, key=lambda row: (score(row), row.id), reverse=True)[:limit]
    ranked.sort(key=lambda row: _recency_key(row.start_date, row.end_date), reverse=True)
    lines = []
    for row in ranked:
        headline = ' — '.join(part for part in [row.degree, row.field, row.institution] if part) or 'Education'
        period = _date_range(row.start_date, row.end_date)
        entry = f'{headline}{" " + period if period else ""}'
        details = ', '.join(part for part in [f'Grade: {row.grade}' if row.grade else '', row.description.strip() if row.description else ''] if part)
        if details:
            entry += f'\n   {details}'
        lines.append(entry)
    return '\n'.join(lines)


def _projects_block(db, terms, limit=3):
    rows = list(db.scalars(select(Project).order_by(Project.id)).all())
    if not rows:
        return ''
    def score(row):
        words = _content_words(_row_text(row))
        return len(terms & words) / (len(words) ** 0.5) if words else 0.0
    ranked = sorted(rows, key=lambda row: (score(row), row.id), reverse=True)[:limit]
    ranked.sort(key=lambda row: row.id)
    lines = []
    for row in ranked:
        entry = f'{row.name}'
        details = ', '.join(part for part in [row.role, f'Technologies: {row.technologies}' if row.technologies else '', row.url] if part)
        if details:
            entry += f' ({details})'
        if row.description:
            entry += f'\n   {row.description.strip()}'
        lines.append(entry)
    return '\n'.join(lines)


def _previous_block(db, terms, limit=3):
    rows = list(db.scalars(select(PreviousAnswer).order_by(PreviousAnswer.updated_at.desc(), PreviousAnswer.id)).all())
    if not rows:
        return ''
    def overlap(row):
        return len(terms & _content_words(row.question, FUNCTION_WORDS))
    best = sorted(rows, key=lambda row: (overlap(row), row.id), reverse=True)[:limit]
    if rows[0] not in best:
        best = [rows[0], *best][:limit]
    return '\n'.join(f'Q: {row.question.strip()}\nA: {trim_to_word_limit(row.answer, 140)}' for row in best)


def _cv_blocks(db, terms, limit=3, char_budget=2200):
    cv = db.scalars(select(CVDocument).order_by(CVDocument.created_at.desc())).first()
    if not cv or not cv.extracted_text:
        return ''
    text = cv.extracted_text
    sections = cv_sections(text)
    header = re.split(r'\n\s*\n', text, maxsplit=1)
    body = header[1] if len(header) > 1 else text

    chunks = []
    for block in [chunk.strip() for chunk in re.split(r'\n\s*\n', body) if len(chunk.strip()) > 40]:
        if '@' in block and re.search(r'\+?\d[\d\s().-]{7,}', block):
            continue
        sentences = [part.strip() for part in re.split(r'(?<=[.!?])\s+', re.sub(r'\n+', ' ', block)) if part.strip()]
        window = []
        for sentence in sentences:
            window.append(sentence)
            if len(' '.join(window)) > 420:
                chunks.append(' '.join(window))
                window = []
        if window:
            chunks.append(' '.join(window))
    for name, values in sections.items():
        for value in values:
            trimmed = value.strip()
            if len(trimmed) > 40:
                chunks.append(f'[{name}] {trimmed[:600]}')
    if not chunks:
        return ''

    def score(chunk):
        words = _content_words(chunk)
        overlap = len(terms & words)
        return overlap / (len(words) ** 0.5) if overlap else 0.0

    ranked = sorted(enumerate(chunks), key=lambda item: (score(item[1]), -item[0]), reverse=True)
    selected, used = [], 0
    for _, chunk in ranked:
        if score(chunk) <= 0 and selected:
            break
        if used + len(chunk) > char_budget:
            break
        selected.append(chunk)
        used += len(chunk)
        if len(selected) >= limit:
            break
    if not selected:
        selected = [chunks[0][:char_budget]]
    return '\n'.join(selected)[:char_budget]


BLOCK_TITLES = {
    'headline': 'ABOUT ME',
    'profile': 'ABOUT ME',
    'skills': 'SKILLS',
    'experience': 'WORK EXPERIENCE',
    'projects': 'PROJECTS',
    'education': 'EDUCATION',
    'previous_answers': 'PREVIOUSLY APPROVED ANSWERS',
    'cv': 'CV EXCERPTS',
    'availability': 'LOGISTICS',
    'compensation': 'COMPENSATION',
}


def question_guidance(question):
    text = str(question or '')
    for name, pattern, instruction, focus in QUESTION_GUIDANCE:
        if pattern.search(text):
            return name, instruction, focus
    return 'general', DEFAULT_GUIDANCE[0], DEFAULT_GUIDANCE[1]


FACT_REQUIREMENTS = {
    'compensation': (('salary_expectation', 'a salary expectation or figure'),),
    'availability': (('availability', 'a start date'), ('notice_period', 'a notice period')),
    'introduction': (('summary', 'a professional summary'),),
    'motivation': (('headline', 'a current job title'),),
    'project': (('headline', 'a current job title'),),
    'leadership': (('headline', 'a current job title'),),
    'weakness': (('headline', 'a current job title'),),
    'strength': (('headline', 'a current job title'),),
}


def missing_facts(db, kind):
    """What the evidence does not contain, so the model can say so instead of inventing it."""
    profile = _profile_row(db)
    gaps = []
    if profile is None:
        return ['any saved personal details']
    for column, label in FACT_REQUIREMENTS.get(kind, ()):
        if not getattr(profile, column, ''):
            gaps.append(label)
    if not db.scalar(select(Education).limit(1)):
        gaps.append('education history')
    if not db.scalar(select(Experience).limit(1)):
        gaps.append('work experience')
    if not db.scalar(select(Skill).limit(1)):
        gaps.append('a list of skills')
    if not db.scalar(select(Project).limit(1)):
        gaps.append('project details')
    if not db.scalar(select(CVDocument).limit(1)):
        gaps.append('a CV')
    return gaps


def context_for(db, question, max_chars=6000):
    name, _, focus = question_guidance(question)
    terms = _question_terms(question)
    profile = _profile_row(db)
    profile_text = _profile_block(db)

    builders = {
        'headline': lambda: (profile.summary.strip() if profile and profile.summary else ''),
        'profile': lambda: profile_text,
        'skills': lambda: _skills_block(db),
        'experience': lambda: _experience_block(db, terms),
        'projects': lambda: _projects_block(db, terms),
        'education': lambda: _education_block(db, terms),
        'previous_answers': lambda: _previous_block(db, terms),
        'cv': lambda: _cv_blocks(db, terms),
        'availability': lambda: (', '.join(
            f'{label}: {getattr(profile, column, "")}' for label, column in
            [('Start date', 'availability'), ('Notice period', 'notice_period')]
            if profile and getattr(profile, column, '')) if profile else ''),
        'compensation': lambda: (f'Expected salary: {profile.salary_expectation}' if profile and profile.salary_expectation else ''),
    }
    titles = {key: BLOCK_TITLES[key] for key in builders}
    order = list(focus) + [key for key in ('headline', 'profile', 'skills', 'experience', 'projects', 'education', 'previous_answers', 'cv') if key not in focus]
    seen = set()
    sections, sources = [], []
    for key in order:
        if key in seen:
            continue
        seen.add(key)
        block = (builders[key]() or '').strip()
        if not block:
            continue
        title = titles[key]
        existing = next((item for item in sections if item[0] == title), None)
        if existing:
            existing[1] = f'{existing[1]}\n{block}'
        else:
            sections.append([title, block])
            sources.append(key)
    rendered = [f'{title}\n{block}' for title, block in sections]
    return '\n\n'.join(rendered)[:max_chars], sources, name, missing_facts(db, name)


# --------------------------------------------------------------------------------------
# Answer generation
# --------------------------------------------------------------------------------------

LENGTH_GUIDANCE = {
    'concise': ('2 to 3 sentences, roughly 45 to 70 words', 140, 70),
    'standard': ('4 to 6 sentences, roughly 100 to 150 words', 260, 140),
    'detailed': ('6 to 8 sentences, roughly 180 to 240 words', 420, 210),
}

SELF_REFERENCES = re.compile(
    r'\b(the\s+candidate|the\s+applicant|this\s+candidate|that\s+candidate|'
    r'the\s+person\s+applying|the\s+individual|the\s+author\s+of\s+this\s+answer)\b', re.I)
LEAD_IN = re.compile(
    r"^\s*(?:sure|certainly|of\s+course|absolutely|okay|here\s+is|here\s+are|here's|"
    r'the\s+answer\s+is|answer|response|a\s+draft|my\s+draft|i\s+would\s+write|i\'d\s+write)\b'
    r'[^\n]{0,60}?[:.—-]\s*', re.I)
AI_PREAMBLE = re.compile(r'^\s*(?:as\s+an\s+ai(?:\s+language\s+model)?|i\s+am\s+an\s+ai)\b[^.]*[.,]\s*', re.I)
THIRD_PARTY = re.compile(r'\b(he|she|they|him|her|them|his|hers|their|theirs)\b', re.I)
PLACEHOLDER_TOKEN = re.compile(
    r'[\[\(<]\s*(?:company|employer|name|date|your|insert|xx|tbd|tbc|n/?a|amount|salary|figure)'
    r'[^\]\)>]{0,24}[\]\)>]'
    r'|[$£€]\s*[Xx]{1,3}(?![A-Za-z])'
    r'|(?<![\w])[Xx]{2,4}(?![\w])'
    r'|\b(?:TBD|TBC)\b', re.I)
FIRST_PERSON = re.compile(r"\b(i|i'm|i've|i'd|my|mine|me|we|our|us)\b", re.I)

_MODEL_CACHE = {'key': None, 'value': None, 'at': 0.0}


def _model_preference():
    preferred = [name.strip() for name in settings.ollama_fallback_models.split(',') if name.strip()]
    return list(dict.fromkeys([settings.ollama_model, *preferred]))


def _is_embedding_model(name):
    return any(token in name.lower() for token in ('embed', 'nomic', 'bge-', 'minilm', 'whisper'))


def choose_model(entries):
    """Pick the configured model when installed, else the best local alternative."""
    configured = settings.ollama_model
    names = {entry.get('name', '') for entry in entries}
    if configured in names:
        return configured
    preference = _model_preference()
    preferred = [name for name in preference if name in names]
    if preferred:
        return preferred[0]
    ranked = sorted(
        (entry for entry in entries if not _is_embedding_model(entry.get('name', ''))),
        key=lambda entry: entry.get('size', 0), reverse=True)
    return ranked[0]['name'] if ranked else configured


async def resolve_model(client):
    """Honour the configured model when it is installed, otherwise pick the strongest local one."""
    key = f'{settings.ollama_base_url}|{settings.ollama_model}'
    now = time.monotonic()
    if _MODEL_CACHE['key'] == key and now - _MODEL_CACHE['at'] < 300:
        return _MODEL_CACHE['value']
    chosen = settings.ollama_model
    try:
        response = await client.get(f'{settings.ollama_base_url}/api/tags', timeout=10)
        if response.status_code == 200:
            chosen = choose_model(response.json().get('models', []))
    except Exception:
        pass
    _MODEL_CACHE.update({'key': key, 'value': chosen, 'at': now})
    return chosen


def probe_model():
    """Synchronous model lookup for the health endpoint."""
    if settings.provider == 'openai':
        return settings.llm_model
    if _MODEL_CACHE['value']:
        return _MODEL_CACHE['value']
    try:
        response = httpx.get(f'{settings.ollama_base_url}/api/tags', timeout=3)
        if response.status_code == 200:
            return choose_model(response.json().get('models', []))
    except Exception:
        pass
    return settings.ollama_model


def trim_to_word_limit(answer, max_words):
    if not answer or max_words <= 0:
        return ''
    words = answer.split()
    if len(words) <= max_words:
        return answer.strip()
    sentences = re.split(r'(?<=[.!?])\s+', answer.strip())
    kept = []
    count = 0
    for sentence in sentences:
        sentence_words = sentence.split()
        if count + len(sentence_words) > max_words:
            remaining = max_words - count
            if remaining <= 0:
                break
            kept.append(' '.join(sentence_words[:remaining]))
            break
        kept.append(sentence)
        count += len(sentence_words)
    if kept:
        return ' '.join(part for part in kept if part).strip()
    return ' '.join(words[:max_words]).strip()


def dedupe_sentences(answer):
    seen, kept = set(), []
    for sentence in re.split(r'(?<=[.!?])\s+', answer.strip()):
        key = re.sub(r'[^a-z0-9 ]', '', sentence.lower())
        key = ' '.join(sorted(set(key.split())))
        if key in seen:
            continue
        seen.add(key)
        kept.append(sentence.strip())
    return ' '.join(part for part in kept if part).strip()


def clean_answer(answer):
    text = str(answer or '').strip()
    # Reasoning models (qwen3, deepseek-r1...) may wrap hidden thoughts in <think> tags.
    text = re.sub(r'<think>.*?(</think>|$)', '', text, flags=re.S | re.I).strip()
    text = re.sub(r'^```[a-z]*\s*|\s*```$', '', text)
    text = AI_PREAMBLE.sub('', text)
    for _ in range(4):
        stripped = LEAD_IN.sub('', text)
        stripped = re.sub(r'^[\s!.,:;*\-]+', '', stripped)
        if stripped == text:
            break
        text = stripped
    text = re.sub(r'^\s*[-*•]\s*', '', text, flags=re.M)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{2,}', '\n', text).strip()
    return dedupe_sentences(text)


def apply_first_person(answer):
    """Last-resort rewrite when the model still answers about itself in third person."""
    if not answer or first_person_ratio(answer) >= 0.02:
        return answer
    text = SELF_REFERENCES.sub('I', answer)
    text = re.sub(r'\bI\s+has\b', 'I have', text)
    text = re.sub(r'\bI\s+does\b', 'I do', text)
    text = re.sub(r'\bI\s+is\s+a\b', 'I am a', text)
    return text


def build_prompt(question, context, answer_length, guidance, missing=(), correction=''):
    length_text, _, max_words = LENGTH_GUIDANCE.get(answer_length, LENGTH_GUIDANCE['standard'])
    lines = [
        'You are helping a person fill in an online job application. Write the answer they will submit.',
        '',
        'Rules:',
        '- Write in the FIRST PERSON, using "I" and "my". Never refer to yourself as "the candidate", "the applicant", "he", "she" or "they".',
        '- Answer ONLY the question that is asked. Do not summarise the profile, do not list the CV, and do not answer a different question.',
        f'- Focus for this question: {guidance}',
        '- Ground every claim in the EVIDENCE below. Copy employer names, school names, project names, technologies and dates exactly as they appear there.',
        '- Never invent a figure, date, employer, project or result. If you need a specific detail and the evidence does not give it, say plainly that you would rather discuss it.',
        '- Never leave bracketed placeholders such as [Company], [Date] or <your name> in the answer, and never write X, XX or TBD for a value you do not know.',
        '- Plain, confident, specific sentences. No bullet points, no headings, no preamble, no sign-off.',
        f'- Length: {length_text}. Hard limit: {max_words} words.',
        '',
        'EVIDENCE',
        context or '(no saved profile or CV information)',
        '',
    ]
    if missing:
        lines += ['NOT RECORDED IN THE EVIDENCE', ', '.join(missing), 'Do not guess at these. Answer with what is recorded and be brief about the rest.', '']
    lines += ['QUESTION', question, '']
    if correction:
        lines += [f'Your previous draft had this problem: {correction}', 'Rewrite it now, fixing only that problem.', '']
    lines += [f'Write only the answer itself, in the first person, in at most {max_words} words.']
    return '\n'.join(lines)


async def _call_ollama(client, model, prompt, token_budget, temperature):
    response = await client.post(
        f'{settings.ollama_base_url}/api/generate',
        json={
            'model': model,
            'prompt': prompt,
            'stream': False,
            'think': False,
            'keep_alive': '30m',
            'options': {'temperature': temperature, 'top_p': 0.9, 'repeat_penalty': 1.05, 'num_predict': token_budget, 'num_ctx': 8192},
        })
    response.raise_for_status()
    return response.json().get('response', '')


async def _call_openai(client, model, prompt, token_budget, temperature):
    """Any OpenAI-compatible chat API: Groq, OpenRouter, Together, Gemini, OpenAI, vLLM..."""
    payload = {
        'model': model,
        'messages': [{'role': 'user', 'content': prompt}],
        'temperature': temperature,
        'top_p': 0.9,
        'max_tokens': token_budget + settings.llm_max_tokens_extra,
        'stream': False,
    }
    if settings.llm_reasoning_effort:
        payload['reasoning_effort'] = settings.llm_reasoning_effort
    response = await client.post(
        f"{settings.llm_base_url.rstrip('/')}/chat/completions",
        headers={'Authorization': f'Bearer {settings.llm_api_key}'},
        json=payload)
    response.raise_for_status()
    choices = response.json().get('choices') or [{}]
    return (choices[0].get('message') or {}).get('content') or ''


async def generate_local(question, context, answer_length='standard', guidance=None, missing=()):
    """Draft an answer with the configured LLM provider (local Ollama or a hosted API)."""
    if guidance is None:
        _, guidance, _ = question_guidance(question)
    length_text, token_budget, max_words = LENGTH_GUIDANCE.get(answer_length, LENGTH_GUIDANCE['standard'])
    prompt = build_prompt(question, context, answer_length, guidance, missing)
    hosted = settings.provider == 'openai'
    call = _call_openai if hosted else _call_ollama

    try:
        async with httpx.AsyncClient(timeout=settings.llm_timeout_seconds if hosted else 90) as client:
            model = settings.llm_model if hosted else await resolve_model(client)
            raw = await call(client, model, prompt, token_budget, 0.35)
            answer = clean_answer(raw)
            correction = _diagnose(answer, question, context)
            if correction:
                retry_prompt = build_prompt(question, context, answer_length, guidance, missing, correction)
                raw = await call(client, model, retry_prompt, token_budget, 0.15)
                retry_answer = clean_answer(raw)
                if not _diagnose(retry_answer, question, context):
                    answer = retry_answer
            return trim_to_word_limit(apply_first_person(answer), max_words)
    except httpx.HTTPStatusError as error:
        if error.response.status_code == 429:
            raise RuntimeError('The answer service is busy right now. Please try again in a minute.') from error
        if hosted:
            raise RuntimeError('The answer service rejected the request. Check LLM_API_KEY and LLM_MODEL on the server.') from error
        raise RuntimeError('Local Ollama is unavailable. Start Ollama and verify the selected model.') from error
    except httpx.HTTPError as error:
        if hosted:
            raise RuntimeError('The answer service could not be reached. Please try again shortly.') from error
        raise RuntimeError('Local Ollama is unavailable. Start Ollama and verify the selected model.') from error


def first_person_ratio(answer):
    words = re.findall(r"[A-Za-z']+", answer)
    if not words:
        return 0.0
    personal = sum(1 for word in words if FIRST_PERSON.fullmatch(word))
    return personal / len(words)


def relevance_score(answer, question):
    terms = _question_terms(question)
    if not terms:
        return 0.0
    return len(terms & _content_words(answer, FUNCTION_WORDS)) / (len(terms) ** 0.5)


def off_topic(answer, question, context=''):
    """True only when an answer shares nothing with the question and nothing with the evidence."""
    if relevance_score(answer, question) > 0:
        return False
    answer_terms = _content_words(answer, FUNCTION_WORDS)
    if not answer_terms:
        return True
    if answer_terms & _content_words(context, FUNCTION_WORDS):
        return False
    return True


def _diagnose(answer, question, context):
    if not answer.strip():
        return 'the draft was empty'
    if first_person_ratio(answer) < 0.02:
        return 'it was written in third person instead of using "I" and "my"'
    if off_topic(answer, question, context):
        return 'it did not address the question that was asked'
    unknown = unsupported_terms(answer, context)
    if unknown:
        return f'it mentioned details that are not in the evidence: {", ".join(unknown[:4])}'
    if PLACEHOLDER_TOKEN.search(answer):
        return 'it left an unfilled placeholder such as [Company] or <name> in the answer'
    return ''


COMMON_CAPS = {
    'I', "I'm", "I've", "I'd", 'The', 'This', 'That', 'These', 'Those', 'My', 'Our', 'We', 'A', 'An', 'And', 'But', 'Or', 'In',
    'On', 'At', 'For', 'To', 'As', 'I’m', 'I’ve', 'I’d', 'It', 'Its', 'If', 'So', 'Not', 'No', 'Yes', 'Then', 'There', 'Their',
    'They', 'He', 'She', 'His', 'Her', 'Him', 'Them', 'When', 'Where', 'What', 'Why', 'How', 'While', 'Although', 'Because',
    'Additionally', 'Moreover', 'Furthermore', 'However', 'Overall', 'Additionally', 'Based', 'Built', 'Also', 'During', 'Through',
    'Across', 'With', 'Without', 'Since', 'Until', 'Between', 'Within', 'Throughout', 'Currently', 'Recently', 'Previously',
    'Working', 'Worked', 'Helped', 'Led', 'Managed', 'Developed', 'Created', 'Designed', 'Implemented', 'Delivered', 'Improved',
    'Increased', 'Reduced', 'Grew', 'Gained', 'Earned', 'Won', 'Completed', 'Started', 'Joined', 'Graduated', 'Studied',
    'Every', 'Each', 'Both', 'Few', 'Many', 'Most', 'Some', 'Such', 'Very', 'Well', 'Please', 'Thank', 'Dear', 'Sincerely',
    'Software', 'Engineering', 'Project', 'Projects', 'Experience', 'Education', 'Skills', 'Team', 'Company', 'Role', 'University',
    'College', 'School', 'Bachelor', 'Master', 'Doctor', 'Full', 'Part', 'Time', 'Remote', 'Hybrid', 'Onsite', 'Salary', 'Location',
    'Email', 'Phone', 'Name', 'Date', 'Years', 'Year', 'Months', 'Month', 'Time', 'Level', 'Field', 'Data', 'System', 'Systems',
    'Application', 'Business', 'Product', 'Products', 'Client', 'Clients', 'Customer', 'Customers', 'Service', 'Services', 'Process',
    'Processes', 'Report', 'Reports', 'Analysis', 'Results', 'Goals', 'Goal', 'Plan', 'Plans', 'Value', 'Values', 'Ability', 'Skills',
}


def unsupported_terms(answer, context):
    """Proper nouns, placeholders and numbers in the answer that the evidence does not support."""
    haystack = str(context or '').lower()
    words = set(re.findall(r'[a-z][a-z0-9\-]{2,}', haystack))
    unknowns = []
    for token in re.findall(r'[\[\(<]?\s*[A-Z][A-Za-z0-9+#.\-]{2,}', answer):
        plain = token.strip('[]()<> ').rstrip('.,;:')
        if plain in COMMON_CAPS or plain in unknowns:
            continue
        lowered = plain.lower()
        if lowered in words or lowered in haystack:
            continue
        if len(lowered) > 3 and any(word.startswith(lowered[:4]) for word in words):
            continue
        unknowns.append(plain)
    for number in set(re.findall(r'\b\d[\d,.]*\b', answer)):
        if number.replace(',', '') and number not in haystack and number.replace(',', '') not in haystack:
            unknowns.append(number)
    return sorted(set(unknowns))


def validate_answer(answer, context, question=''):
    if not answer or not answer.strip():
        return False, 'The generated answer was empty.'
    if len(answer) > 4000:
        return False, 'The generated answer is too long.'
    if first_person_ratio(answer) < 0.02:
        return False, 'The answer is not written in the first person.'
    if off_topic(answer, question, context):
        return False, 'The answer does not address the question that was asked.'
    unknowns = unsupported_terms(answer, context)
    if unknowns:
        return False, f'Contains details that are not in your saved profile or CV: {", ".join(unknowns[:4])}.'
    if PLACEHOLDER_TOKEN.search(answer):
        return False, 'The answer still contains an unfilled placeholder such as [Company] or <name>.'
    return True, 'Passed local checks and matches your saved profile and CV.'
