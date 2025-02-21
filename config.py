import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # Flask configuration
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-key-change-in-production'

    # Database configuration with SSL parameters
    db_url = os.environ.get('DATABASE_URL')
    if db_url and 'postgresql' in db_url:
        # Add SSL mode if not present
        if '?' not in db_url:
            db_url += '?sslmode=require'
        elif 'sslmode=' not in db_url:
            db_url += '&sslmode=require'
    SQLALCHEMY_DATABASE_URI = db_url
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Admin configuration
    ADMIN_EMAIL = os.environ.get('ADMIN_EMAIL', 'admin@example.com')
    ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'change-in-production')

    # Subscription configuration
    TRIAL_PERIOD_DAYS = 14