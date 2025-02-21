import streamlit as st
import os
import logging
from utils import setup_logging

setup_logging()
logger = logging.getLogger(__name__)

def show_setup_guide():
    st.title("SharePoint File Manager Setup Guide")
    
    st.write("### Step 1: Register Azure AD Application")
    st.markdown("""
    1. Go to [Azure Portal](https://portal.azure.com)
    2. Navigate to Azure Active Directory → App registrations
    3. Click "New registration"
    4. Fill in the application details:
        - Name: SharePoint File Manager (or your preferred name)
        - Supported account types: "Accounts in this organizational directory only"
        - Redirect URI: Leave blank (Not required for this app)
    5. Click "Register"
    """)
    
    st.write("### Step 2: Configure API Permissions")
    st.markdown("""
    1. In your newly registered app, go to "API permissions"
    2. Click "Add a permission"
    3. Select "Microsoft Graph"
    4. Choose "Application permissions"
    5. Add these permissions:
        - Sites.Read.All (Required for basic file listing and reading)
        - Sites.ReadWrite.All (Required for file operations and test library creation)
        - Sites.Manage.All (Required for creating and managing test libraries)
    6. Click "Grant admin consent" (requires admin privileges)

    Note: The test library feature requires both Sites.ReadWrite.All and Sites.Manage.All permissions to create and populate sample data.
    """)

    st.write("### Step 3: Create Client Secret")
    st.markdown("""
    1. Go to "Certificates & secrets"
    2. Click "New client secret"
    3. Add a description and choose an expiration
    4. Click "Add"
    5. **Important:** Copy the secret value immediately - you won't be able to see it again!
    """)
    
    st.write("### Step 4: Configure Application")
    
    # Create columns for a cleaner layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("#### Required Information")
        st.markdown("""
        You'll need these values from Azure:
        - Application (client) ID
        - Directory (tenant) ID
        - Client secret
        """)
    
    with col2:
        st.write("#### Where to Find")
        st.markdown("""
        - Client ID: Overview page
        - Tenant ID: Overview page
        - Client Secret: Value copied in Step 3
        """)
    
    st.write("### Step 5: Enter Credentials")
    
    with st.form("azure_credentials_form"):
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
            help="The secret value you copied in Step 3"
        )
        
        save_credentials = st.form_submit_button("Save Credentials")
        
        if save_credentials:
            if all([client_id, tenant_id, client_secret]):
                try:
                    # Store credentials securely
                    os.environ["AZURE_CLIENT_ID"] = client_id
                    os.environ["AZURE_TENANT_ID"] = tenant_id
                    os.environ["AZURE_CLIENT_SECRET"] = client_secret
                    
                    st.success("Credentials saved successfully!")
                    st.info("You can now return to the main application and connect to SharePoint.")
                    logger.info("Azure credentials configured successfully")
                    
                except Exception as e:
                    st.error(f"Error saving credentials: {str(e)}")
                    logger.error(f"Error saving credentials: {str(e)}")
            else:
                st.warning("Please fill in all credential fields.")
    
    st.write("### Common Issues")
    with st.expander("Troubleshooting"):
        st.markdown("""
        #### Authentication Failed
        - Verify all credentials are correct
        - Ensure admin consent is granted for API permissions
        - Check if the client secret hasn't expired
        
        #### Access Denied
        - Verify the Azure AD app has the correct API permissions
        - Ensure admin consent is granted
        - Check if your SharePoint site URL is correct
        
        #### Connection Issues
        - Verify your internet connection
        - Check if SharePoint is accessible
        - Ensure your tenant name is correct in the site URL
        """)
    
    st.write("### Need Help?")
    st.info("""
    For additional assistance:
    1. Check the [Azure AD documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
    2. Contact your Azure AD administrator
    3. Review the [Microsoft Graph permissions reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
    """)

if __name__ == "__main__":
    show_setup_guide()