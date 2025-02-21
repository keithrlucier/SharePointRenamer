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
                1. Install an authenticator app on your device:
                   - Microsoft Authenticator (recommended)
                   - Google Authenticator
                   - Authy

                2. Open your authenticator app and add a new account:
                   - Click '+' or 'Add Account'
                   - Choose 'Scan QR Code'
                   - When prompted, select "Work or school account"

                3. Scan the QR code below
                4. Enter the 6-digit code shown in your authenticator app
                """)

                try:
                    # Generate QR code with smaller size
                    qr = qrcode.QRCode(
                        version=1,
                        error_correction=qrcode.constants.ERROR_CORRECT_L,
                        box_size=6,  # Reduced from 10
                        border=2,    # Reduced from 4
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
                        st.warning("⚠️ Important: Please follow these steps carefully:")
                        st.markdown("""
                        1. Wait for a new code to appear in your authenticator
                        2. Enter the code immediately when it appears
                        3. Make sure your device's time is correctly synchronized
                        4. Only use digits 0-9 (no spaces or special characters)
                        """)

                        # Add account type confirmation
                        st.info("🔍 Important: Make sure you've selected 'Work or school account' in your authenticator app")

                        verification_code = st.text_input(
                            "Enter verification code from your authenticator app",
                            help="Enter the 6-digit code shown in your authenticator app",
                            max_chars=6
                        )
                        submit = st.form_submit_button("Verify and Enable 2FA")

                        if submit:
                            if not verification_code:
                                st.error("Please enter a verification code.")
                            elif not verification_code.isdigit() or len(verification_code) != 6:
                                st.error("Please enter exactly 6 digits (0-9 only).")
                            else:
                                if user.verify_mfa(verification_code):
                                    user.mfa_enabled = True
                                    db.session.commit()
                                    logger.info(f"MFA enabled successfully for user {user.email}")
                                    st.success("Two-factor authentication enabled successfully!")
                                    st.rerun()
                                else:
                                    st.error("Invalid verification code. Please ensure:")
                                    st.markdown("""
                                    1. Your device's time is correctly synchronized
                                    2. You're using a fresh code from your authenticator
                                    3. You selected "Work or school account" during setup
                                    4. You entered the code immediately after it appeared
                                    """)
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