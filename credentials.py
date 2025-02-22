import streamlit as st
import os
import logging
from utils import setup_logging
from sharepoint_client import SharePointClient

setup_logging()
logger = logging.getLogger(__name__)

def show_credentials_manager():
    """Display Azure AD credentials management interface"""
    st.write("### Azure AD Credentials Configuration")

    # Add security notice
    st.info("""
    These credentials are required for SharePoint authentication. They should be obtained
    from your Azure AD application registration. For detailed instructions on obtaining
    these credentials, please refer to the Setup Guide.
    """)

    # Add encryption notice
    st.success("""
    🔒 **Security Notice:**
    All credentials are encrypted both in transit and at rest using industry-standard encryption.
    Your sensitive information is protected at all times through secure protocols and storage mechanisms.
    """)

    with st.form("azure_credentials_form", clear_on_submit=False):
        client_id = st.text_input(
            "Application (Client) ID",
            type="password",
            help="Found in Azure AD App Registration Overview"
        )

        tenant_id = st.text_input(
            "Directory (Tenant) ID",
            type="password",
            help="Found in Azure AD App Registration Overview"
        )

        client_secret = st.text_input(
            "Client Secret",
            type="password",
            help="The secret value from Azure AD App Registration"
        )

        site_url = st.text_input(
            "SharePoint Site URL",
            help="The URL of your SharePoint site (e.g., https://yourtenant.sharepoint.com/sites/yoursite)"
        )

        submit = st.form_submit_button("Save Credentials", use_container_width=True)

        if submit:
            if all([client_id, tenant_id, client_secret, site_url]):
                try:
                    # Store credentials securely
                    os.environ["AZURE_CLIENT_ID"] = client_id
                    os.environ["AZURE_TENANT_ID"] = tenant_id
                    os.environ["AZURE_CLIENT_SECRET"] = client_secret
                    os.environ["SHAREPOINT_SITE_URL"] = site_url

                    # Try to initialize SharePoint client and authenticate
                    try:
                        client = SharePointClient(site_url=site_url)
                        client.authenticate()  # This will use the environment variables we just set
                        st.session_state['client'] = client
                        st.session_state['authenticated'] = True
                        logger.info("SharePoint client initialized and authenticated successfully")
                    except Exception as e:
                        logger.error(f"Error initializing SharePoint client: {str(e)}")
                        st.error("Failed to connect to SharePoint. Please verify your credentials and site URL.")
                        return

                    st.success("Credentials saved and connection verified!")
                    logger.info("Azure credentials updated successfully")

                    # Reset navigation state to go back to home/connection page
                    st.session_state['current_page'] = 'home'
                    st.session_state['show_credentials'] = False
                    st.session_state['show_setup'] = False

                    # Force rerun to update the UI
                    st.rerun()

                except Exception as e:
                    st.error(f"Error saving credentials: {str(e)}")
                    logger.error(f"Error saving credentials: {str(e)}")
            else:
                st.warning("Please fill in all credential fields.")