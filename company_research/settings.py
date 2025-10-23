import os
from decouple import config
import dj_database_url

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SECRET_KEY = config('SECRET_KEY', default='django-insecure-change-me')
DEBUG = config('DEBUG', default=True, cast=bool)
ALLOWED_HOSTS = ['*']

INSTALLED_APPS = [
    'research',  
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'company_research.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'company_research.wsgi.application'

# Database - Neon PostgreSQL
DATABASES = {
    'default': dj_database_url.parse(
        config('DATABASE_URL', default='postgresql://neondb_owner:npg_m7N0XzUxPnMu@ep-spring-hat-a8qw0bvp-pooler.eastus2.azure.neon.tech/neondb?sslmode=require&channel_binding=require'),
        conn_max_age=600,
        conn_health_checks=True,
    )
}

LANGUAGE_CODE = 'ja'
TIME_ZONE = 'Asia/Tokyo'
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATIC_ROOT = os.path.join(BASE_DIR, 'staticfiles')
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Authentication
LOGIN_URL = '/admin/login/'
LOGIN_REDIRECT_URL = '/admin/'

# API Keys
API_KEY = config('API_KEY', default='AIzaSyAjiy9pQMwCwM6zeFsBk0M2dRlE8WhoME8')
SPREADSHEET_ID = config('SPREADSHEET_ID', default='1qoUziLx0uOhFbDwSlEdjs81wkW2ehXgORvINZUxsCuY')
OPEN_AI_API_KEY = config('OPEN_AI_API_KEY', default='')

CREDENTIALS_INFO = {
    "type": "service_account",
    "project_id": config('PROJECT_ID', default='cogent-tract-472405-k8'),
    "private_key_id": config('PRIVATE_KEY_ID', default=''),
    "private_key": config('PRIVATE_KEY', default='').replace('\\n', '\n'),
    "client_email": config('CLIENT_EMAIL', default=''),
    "client_id": config('CLIENT_ID', default=''),
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": config('CLIENT_X509_CERT_URL', default=''),
    "universe_domain": "googleapis.com"
}
