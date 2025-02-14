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

def authenticate():
    """Handle SharePoint authentication"""
    with st.form("authentication_form"):
        site_url = st.text_input("SharePoint Site URL")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Connect")

        if submit:
            try:
                st.session_state.client = SharePointClient(site_url, username, password)
                st.session_state.authenticated = True
                st.success("Successfully connected to SharePoint!")
                logger.info(f"User {username} authenticated successfully")
            except Exception as e:
                st.error(f"Authentication failed: {str(e)}")
                logger.error(f"Authentication failed: {str(e)}")

def show_library_selector():
    """Display SharePoint library selector"""
    libraries = st.session_state.client.get_libraries()
    selected_library = st.selectbox("Select SharePoint Library", libraries)
    if selected_library:
        show_file_manager(selected_library)

def show_file_manager(library_name):
    """Display file management interface"""
    try:
        files = st.session_state.client.get_files(library_name)
        
        st.write("### Files in Library")
        
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
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.experimental_rerun()
        
        show_library_selector()

if __name__ == "__main__":
    main()
