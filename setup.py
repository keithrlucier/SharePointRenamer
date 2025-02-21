import streamlit as st
import logging
from utils import setup_logging

setup_logging()
logger = logging.getLogger(__name__)

def show_setup_guide():
    st.write("### Setup Guide")

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
    4. Choose "Application permissions" (NOT Delegated permissions)
       - Application permissions are required because this app runs without user context
       - Delegated permissions will not work for this application's background operations
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

    st.write("### Step 4: Set Up Two-Factor Authentication")
    st.markdown("""
    For enhanced security, we support two-factor authentication using authenticator apps.

    #### Authenticator App Setup:
    1. Install a supported authenticator app on your mobile device:
        - Microsoft Authenticator (recommended for SharePoint integration)
        - Google Authenticator
        - Authy

    2. When adding this application to your authenticator:
        - Select "Work or school account" when prompted
        - This is important because our app integrates with SharePoint Enterprise
        - Selecting the wrong account type may cause verification issues

    3. Tips for successful 2FA setup:
        - Ensure your device's time is synchronized correctly
        - Wait for the code to refresh before using it
        - Enter the code promptly when verifying
    """)

    st.write("### Step 5: Configure Application")

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

    st.write("### Common Issues")
    with st.expander("Troubleshooting"):
        st.markdown("""
        #### Authentication Failed
        - Verify all credentials are correct
        - Ensure admin consent is granted for API permissions
        - Check if the client secret hasn't expired

        #### 2FA Issues
        - Verify you selected "Work or school account" in authenticator
        - Ensure your device's time is correctly synchronized
        - Try waiting for a new code if verification fails
        - Enter the code as soon as it appears in your authenticator

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

    # Add button to return to credentials manager
    if st.button("⚙️ Configure Credentials"):
        st.session_state['show_credentials'] = True
        st.session_state['show_setup'] = False
        st.rerun()

if __name__ == "__main__":
    show_setup_guide()