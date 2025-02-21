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

        # Add navigation buttons at the top
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if user.is_admin:
                if st.button("🔧 Admin Controls"):
                    st.session_state['page'] = 'admin'
                    st.rerun()

        if not user.mfa_enabled:
            st.info("""
            Two-factor authentication adds an extra layer of security to your account.
            Please follow these steps to set up 2FA:
            """)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("""
                1. Install an authenticator app:
                   - Microsoft Authenticator
                   - Google Authenticator
                   - Authy
                """)

            with col2:
                st.markdown("""
                2. Open your authenticator app
                3. Add a new account
                4. Scan the QR code below
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

                # Verification section
                st.markdown("### Verify Setup")
                st.info("Enter the 6-digit code from your authenticator app to complete setup")

                code = st.text_input(
                    "Verification Code",
                    max_chars=6,
                    help="Enter the 6-digit code shown in your authenticator app"
                )

                if st.button("Verify and Enable 2FA"):
                    if not code:
                        st.error("Please enter a verification code")
                    elif not code.isdigit() or len(code) != 6:
                        st.error("Please enter a valid 6-digit code")
                    else:
                        if user.verify_mfa(code):
                            user.mfa_enabled = True
                            db.session.commit()
                            logger.info(f"MFA enabled successfully for user {user.email}")
                            st.success("Two-factor authentication enabled successfully!")
                            st.rerun()
                        else:
                            st.error("Invalid verification code")
                            logger.warning(f"Failed MFA verification attempt for user {user.email}")

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