from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime
import bcrypt
import pyotp

db = SQLAlchemy()

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.LargeBinary, nullable=False)  # Store as binary
    is_admin = db.Column(db.Boolean, default=False)
    is_active = db.Column(db.Boolean, default=True)
    mfa_secret = db.Column(db.String(32))
    mfa_enabled = db.Column(db.Boolean, default=False)
    tenant_id = db.Column(db.Integer, db.ForeignKey('tenant.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        if isinstance(password, str):
            password = password.encode('utf-8')
        self.password_hash = bcrypt.hashpw(password, bcrypt.gensalt())

    def check_password(self, password):
        if not self.password_hash:
            return False
        if isinstance(password, str):
            password = password.encode('utf-8')
        try:
            return bcrypt.checkpw(password, self.password_hash)
        except Exception:
            return False

    def get_mfa_uri(self):
        if not self.mfa_secret:
            self.mfa_secret = pyotp.random_base32()
        return pyotp.totp.TOTP(self.mfa_secret).provisioning_uri(
            name=self.email,
            issuer_name="SharePoint Manager"
        )

    def verify_mfa(self, code):
        if not self.mfa_enabled or not self.mfa_secret:
            return False
        totp = pyotp.TOTP(self.mfa_secret)
        return totp.verify(code)

class Tenant(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    subscription_status = db.Column(db.String(20), default='trial')  # trial, active, cancelled
    subscription_end = db.Column(db.DateTime)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    users = db.relationship('User', backref='tenant', lazy=True)
    client_credentials = db.relationship('ClientCredential', backref='tenant', lazy=True)

class ClientCredential(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tenant_id = db.Column(db.Integer, db.ForeignKey('tenant.id'))
    client_id = db.Column(db.String(100), nullable=False)
    client_secret = db.Column(db.String(100), nullable=False)
    tenant_id_azure = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_updated = db.Column(db.DateTime, onupdate=datetime.utcnow)

    def decrypt_secret(self):
        # TODO: Implement proper encryption/decryption
        return self.client_secret