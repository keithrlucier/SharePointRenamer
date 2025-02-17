import streamlit as st
import logging
from sharepoint_client import SharePointClient
from utils import setup_logging, validate_filename, sanitize_filename
import time
import re
import os

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

def initialize_session_state():
    """Initialize session state variables"""
    if 'authenticated' not in st.session_state:
        st.session_state['authenticated'] = False
    if 'client' not in st.session_state:
        st.session_state['client'] = None
    if 'selected_files' not in st.session_state:
        st.session_state['selected_files'] = {}
    if 'rename_pattern' not in st.session_state:
        st.session_state['rename_pattern'] = ""
    if 'preview_renames' not in st.session_state:
        st.session_state['preview_renames'] = []

def apply_rename_pattern(filename, pattern):
    """Apply rename pattern to filename"""
    # Extract the file extension
    name, ext = os.path.splitext(filename)

    # Replace placeholders in pattern
    new_name = pattern

    # {name} - original filename without extension
    new_name = new_name.replace('{name}', name)

    # {ext} - original extension including dot
    new_name = new_name.replace('{ext}', ext)

    # Add more pattern replacements here

    return new_name

def show_file_manager(library_name):
    """Display file management interface"""
    try:
        with st.spinner("Loading files..."):
            files = st.session_state.client.get_files(library_name)

        st.write("### Files in Library")
        if not files:
            st.info("No files found in this library.")
            return

        # Group files by parent path
        files_by_path = {}
        for file in files:
            parent_path = file.get('ParentPath', '').split('/')[-1] or 'Root'
            if parent_path not in files_by_path:
                files_by_path[parent_path] = []
            files_by_path[parent_path].append(file)

        # Bulk rename controls
        st.sidebar.write("### Bulk Rename")
        st.sidebar.text_input(
            "Rename Pattern",
            value=st.session_state.rename_pattern,
            help="Use patterns like: prefix_{name}_{ext}",
            key="rename_pattern_input",
            on_change=lambda: setattr(st.session_state, 'rename_pattern', st.session_state.rename_pattern_input)
        )

        if st.sidebar.button("Preview Rename"):
            preview_renames = []
            for file_id in st.session_state.selected_files:
                file = next((f for f in sum(files_by_path.values(), []) if f['Id'] == file_id), None)
                if file:
                    new_name = apply_rename_pattern(file['Name'], st.session_state.rename_pattern)
                    preview_renames.append({
                        'old_name': file['Name'],
                        'new_name': new_name,
                        'file_id': file_id
                    })
            st.session_state.preview_renames = preview_renames

        if st.session_state.preview_renames:
            st.sidebar.write("### Preview Changes")
            for rename in st.session_state.preview_renames:
                st.sidebar.text(f"{rename['old_name']} → {rename['new_name']}")

            if st.sidebar.button("Apply Rename"):
                with st.spinner("Renaming files..."):
                    results = st.session_state.client.bulk_rename_files(
                        library_name,
                        st.session_state.preview_renames
                    )

                    # Show results
                    success_count = sum(1 for r in results if r['success'])
                    if success_count == len(results):
                        st.success(f"Successfully renamed {success_count} files!")
                    else:
                        st.warning(f"Renamed {success_count} out of {len(results)} files.")
                        for result in results:
                            if not result['success']:
                                st.error(f"Failed to rename {result['old_name']}: {result.get('error', 'Unknown error')}")

                    # Clear selections and preview
                    st.session_state.selected_files = {}
                    st.session_state.preview_renames = []
                    time.sleep(1)
                    st.rerun()

        # Display files grouped by folders
        for folder_name, folder_files in files_by_path.items():
            with st.expander(f"📁 {folder_name}", expanded=folder_name == 'Root'):
                for file in folder_files:
                    col1, col2, col3 = st.columns([0.5, 2.5, 1])

                    with col1:
                        # Checkbox for bulk selection
                        is_selected = st.checkbox("", key=f"select_{file['Id']}", 
                                               value=file['Id'] in st.session_state.selected_files)
                        if is_selected:
                            st.session_state.selected_files[file['Id']] = file
                        elif file['Id'] in st.session_state.selected_files:
                            del st.session_state.selected_files[file['Id']]

                    with col2:
                        st.write(f"📄 {file['Name']}")
                        if len(file['Name']) > 128:
                            st.warning("Long filename!")

                    with col3:
                        if st.button(f"Rename", key=f"rename_{file['Id']}"):
                            show_rename_form(library_name, file)

    except Exception as e:
        st.error(f"Error loading files: {str(e)}")
        logger.error(f"Error loading files: {str(e)}")
        if "authentication" in str(e).lower():
            st.session_state.authenticated = False
            st.rerun()

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
                    st.rerun()
            except Exception as e:
                st.error(f"Error renaming file: {str(e)}")
                logger.error(f"Error renaming file: {str(e)}")
        else:
            st.error("Invalid filename. Please try again.")

def main():
    st.title("SharePoint File Name Manager")

    # Initialize session state first
    initialize_session_state()

    if not st.session_state.authenticated:
        authenticate()
    else:
        if st.sidebar.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.rerun()

        show_library_selector()

def authenticate():
    """Handle SharePoint authentication"""
    st.write("### SharePoint Authentication")

    with st.form("authentication_form"):
        site_url = st.text_input("SharePoint Site URL", 
                                 help="Enter the full SharePoint site URL (e.g., https://your-tenant.sharepoint.com/sites/your-site)")
        submit = st.form_submit_button("Connect")

        if submit and site_url:
            try:
                with st.spinner("Connecting to SharePoint..."):
                    client = SharePointClient(site_url)
                    if client.authenticate():
                        st.session_state.client = client
                        st.session_state.authenticated = True
                        st.success("Successfully connected to SharePoint!")
                        time.sleep(2)
                        st.rerun()
            except Exception as e:
                st.error(f"Authentication failed: {str(e)}")
                logger.error(f"Authentication failed: {str(e)}")

                if "AADSTS7000229" in str(e):
                    st.warning("""
                    ### Admin Consent Required
                    This application needs admin consent in Azure AD. Please contact your Azure AD administrator to:
                    1. Go to Azure Portal -> Azure Active Directory -> App registrations
                    2. Find the application
                    3. Click on 'API permissions'
                    4. Click on 'Grant admin consent for [tenant]'

                    Once admin consent is granted, try connecting again.
                    """)
                else:
                    st.info("""
                    Please ensure you:
                    1. Enter a valid SharePoint site URL
                    2. Check your internet connection
                    3. Verify your Azure AD app registration settings
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
            st.rerun()


if __name__ == "__main__":
    main()