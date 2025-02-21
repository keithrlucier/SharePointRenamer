import streamlit as st
import qrcode
import io
from models import User, db
import pyotp
from app import app
import logging

logger = logging.getLogger(__name__)

def show_mfa_setup():
    """Display MFA setup and management interface"""
    st.write("### Two-Factor Authentication Setup")

    if 'user' not in st.session_state:
        st.error("Please log in first")
        return

    # Use Flask app context for database operations
    with app.app_context():
        user = User.query.get(st.session_state['user'])

        if not user:
            st.error("User not found")
            return

        if not user.mfa_enabled:
            st.info("""
            Two-factor authentication adds an extra layer of security to your account.

            Setup Instructions:
            1. Install an authenticator app on your device (like Microsoft Authenticator or Google Authenticator)
            2. Scan the QR code below with your authenticator app
            3. When prompted, select "Work or school account" as this is an enterprise application
            4. Enter the 6-digit code shown in your authenticator app below
            """)

            try:
                # Generate QR code
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=10,
                    border=4,
                )
                qr.add_data(user.get_mfa_uri())
                qr.make(fit=True)

                # Convert QR code to image
                img_buf = io.BytesIO()
                img = qr.make_image(fill_color="black", back_color="white")
                img.save(img_buf, format='PNG')

                # Display QR code
                st.image(img_buf.getvalue(), caption="Scan this QR code with your authenticator app")

                # Manual entry option
                st.write("Or enter this code manually in your authenticator app:")
                st.code(user.mfa_secret)

                # Verification form
                with st.form("mfa_setup_form"):
                    verification_code = st.text_input(
                        "Enter verification code from your authenticator app",
                        help="Enter the 6-digit code shown in your authenticator app"
                    )
                    submit = st.form_submit_button("Verify and Enable 2FA")

                    if submit:
                        if verification_code:
                            if user.verify_mfa(verification_code):
                                user.mfa_enabled = True
                                db.session.commit()
                                logger.info(f"MFA enabled successfully for user {user.email}")
                                st.success("Two-factor authentication enabled successfully!")
                                st.rerun()
                            else:
                                st.error("Invalid verification code. Please make sure your device's time is correct and try again.")
                                logger.warning(f"Failed MFA verification attempt for user {user.email}")
                        else:
                            st.error("Please enter a verification code.")

            except Exception as e:
                logger.error(f"Error in MFA setup: {str(e)}")
                st.error("An error occurred during MFA setup. Please try again.")

        else:
            st.success("Two-factor authentication is enabled")

            if st.button("Disable 2FA", type="secondary"):
                user.mfa_enabled = False
                user.mfa_secret = None
                db.session.commit()
                logger.info(f"MFA disabled for user {user.email}")
                st.success("Two-factor authentication disabled")
                st.rerun()