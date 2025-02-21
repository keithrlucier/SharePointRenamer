from flask import Flask
from models import db
from config import Config

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    
    # Initialize Flask extensions
    db.init_app(app)
    
    return app

# Create and configure the app
app = create_app()

# Create all database tables
with app.app_context():
    db.create_all()
