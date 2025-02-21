import streamlit as st
import os
import logging
from utils import setup_logging

setup_logging()
logger = logging.getLogger(__name__)

def show_credentials_manager():
    """Display Azure AD credentials management interface"""
    st.write("### Azure AD Credentials Configuration")
    st.info("""
    These credentials are required for SharePoint authentication. They should be obtained
    from your Azure AD application registration. For detailed instructions on obtaining
    these credentials, please refer to the Setup Guide.
    """)

    with st.form("azure_credentials_form", clear_on_submit=False):
        client_id = st.text_input(
            "Application (Client) ID",
            type="password",
            value=os.environ.get("AZURE_CLIENT_ID", ""),
            help="Found in Azure AD App Registration Overview"
        )

        tenant_id = st.text_input(
            "Directory (Tenant) ID",
            type="password",
            value=os.environ.get("AZURE_TENANT_ID", ""),
            help="Found in Azure AD App Registration Overview"
        )

        client_secret = st.text_input(
            "Client Secret",
            type="password",
            value=os.environ.get("AZURE_CLIENT_SECRET", ""),
            help="The secret value from Azure AD App Registration"
        )

        submit = st.form_submit_button("Save Credentials", use_container_width=True)

        if submit:
            if all([client_id, tenant_id, client_secret]):
                try:
                    # Store credentials securely
                    os.environ["AZURE_CLIENT_ID"] = client_id
                    os.environ["AZURE_TENANT_ID"] = tenant_id
                    os.environ["AZURE_CLIENT_SECRET"] = client_secret

                    st.success("Credentials saved successfully!")
                    logger.info("Azure credentials updated successfully")

                except Exception as e:
                    st.error(f"Error saving credentials: {str(e)}")
                    logger.error(f"Error saving credentials: {str(e)}")
            else:
                st.warning("Please fill in all credential fields.")