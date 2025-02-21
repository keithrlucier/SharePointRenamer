import logging
from flask import Flask
from models import db, User, Tenant
from config import Config
from flask_migrate import Migrate

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    # Initialize Flask extensions
    db.init_app(app)

    # Initialize Flask-Migrate
    migrate = Migrate(app, db)

    return app

# Create and configure the app
app = create_app()

# Create all database tables
with app.app_context():
    try:
        db.create_all()

        # Create default tenant if it doesn't exist
        default_tenant = Tenant.query.filter_by(name='Default').first()
        if not default_tenant:
            default_tenant = Tenant(
                name='Default',
                subscription_status='active'
            )
            db.session.add(default_tenant)
            db.session.commit()
            logging.info("Created default tenant")

        # Check if admin user exists
        admin = User.query.filter_by(email='admin@example.com').first()
        if not admin:
            admin = User(
                email='admin@example.com',
                is_admin=True,
                is_active=True,
                tenant_id=default_tenant.id
            )
            admin.set_password('admin123')
            db.session.add(admin)
            db.session.commit()
            logging.info("Created initial admin user")

    except Exception as e:
        logging.error(f"Error initializing database: {str(e)}")