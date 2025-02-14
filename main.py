import streamlit as st
import logging
from sharepoint_client import SharePointClient
from utils import setup_logging, validate_filename, sanitize_filename
import time

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

def initialize_session_state():
    """Initialize session state variables"""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'client' not in st.session_state:
        st.session_state.client = None
    if 'auth_in_progress' not in st.session_state:
        st.session_state.auth_in_progress = False

def authenticate():
    """Handle SharePoint authentication with MFA support"""
    st.write("### SharePoint Authentication")
    st.write("""
    Please enter your SharePoint site URL and follow the authentication prompt in your browser.
    You will be asked to sign in with your Microsoft account and complete MFA if required.

    Note: A new browser window will open for secure authentication.
    """)

    with st.form("authentication_form"):
        site_url = st.text_input("SharePoint Site URL", 
                                help="Enter the full SharePoint site URL (e.g., https://your-tenant.sharepoint.com/sites/your-site)")
        submit = st.form_submit_button("Connect")

        if submit and site_url:
            try:
                st.session_state.auth_in_progress = True
                with st.spinner("Initiating authentication process... Please complete the sign-in in your browser window."):
                    client = SharePointClient(site_url)
                    if client.authenticate():
                        st.session_state.client = client
                        st.session_state.authenticated = True
                        st.session_state.auth_in_progress = False
                        st.success("Successfully connected to SharePoint!")
                        logger.info("User authenticated successfully with MFA")
                        time.sleep(2)  # Give user time to see success message
                        st.experimental_rerun()
            except Exception as e:
                st.session_state.auth_in_progress = False
                st.error(f"Authentication failed: {str(e)}")
                logger.error(f"Authentication failed: {str(e)}")
                if "token" in str(e).lower():
                    st.info("""
                    Please ensure you:
                    1. Allow pop-ups in your browser
                    2. Complete the sign-in process in the browser window
                    3. If no browser window opened, try again
                    """)

def show_library_selector():
    """Display SharePoint library selector"""
    try:
        with st.spinner("Loading SharePoint libraries..."):
            libraries = st.session_state.client.get_libraries()
            if not libraries:
                st.warning("No document libraries found in this SharePoint site.")
                return

        selected_library = st.selectbox("Select SharePoint Library", libraries)
        if selected_library:
            show_file_manager(selected_library)
    except Exception as e:
        st.error(f"Error loading libraries: {str(e)}")
        logger.error(f"Error loading libraries: {str(e)}")
        if "authentication" in str(e).lower():
            st.session_state.authenticated = False
            st.experimental_rerun()

def show_file_manager(library_name):
    """Display file management interface"""
    try:
        with st.spinner("Loading files..."):
            files = st.session_state.client.get_files(library_name)

        st.write("### Files in Library")
        if not files:
            st.info("No files found in this library.")
            return

        for file in files:
            col1, col2, col3 = st.columns([3, 1, 1])

            with col1:
                st.write(file['Name'])

            with col2:
                if len(file['Name']) > 128:  # Warning for long filenames
                    st.warning("Long filename!")

            with col3:
                if st.button(f"Rename {file['Name']}", key=file['Id']):
                    show_rename_form(library_name, file)

    except Exception as e:
        st.error(f"Error loading files: {str(e)}")
        logger.error(f"Error loading files: {str(e)}")
        if "authentication" in str(e).lower():
            st.session_state.authenticated = False
            st.experimental_rerun()

def show_rename_form(library_name, file):
    """Display rename form for a file"""
    st.write(f"### Rename File: {file['Name']}")

    new_name = st.text_input(
        "New filename",
        value=file['Name'],
        key=f"new_name_{file['Id']}"
    )

    if st.button("Save", key=f"save_{file['Id']}"):
        if validate_filename(new_name):
            try:
                with st.spinner("Renaming file..."):
                    sanitized_name = sanitize_filename(new_name)
                    st.session_state.client.rename_file(
                        library_name,
                        file['Name'],
                        sanitized_name
                    )
                    time.sleep(1)  # Allow SharePoint to process
                    st.success(f"File renamed to: {sanitized_name}")
                    logger.info(f"File renamed from {file['Name']} to {sanitized_name}")
                    st.experimental_rerun()
            except Exception as e:
                st.error(f"Error renaming file: {str(e)}")
                logger.error(f"Error renaming file: {str(e)}")
        else:
            st.error("Invalid filename. Please try again.")

def main():
    st.title("SharePoint File Name Manager")

    initialize_session_state()

    if not st.session_state.authenticated:
        authenticate()
    else:
        if st.sidebar.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.experimental_rerun()

        show_library_selector()

if __name__ == "__main__":
    main()