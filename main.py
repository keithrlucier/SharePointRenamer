import streamlit as st
import logging
from sharepoint_client import SharePointClient
from utils import setup_logging, validate_filename, sanitize_filename
import time
import re
import os
from setup import show_setup_guide
from credentials import show_credentials_manager

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

# App version
APP_VERSION = "1.0.0"

def show_navigation():
    """Display consistent navigation header"""
    st.markdown("""
    <style>
    .nav-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 1rem 0;
        border-bottom: 1px solid #e5e5e5;
        margin-bottom: 2rem;
    }
    .nav-title {
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .nav-version {
        font-size: 0.8rem;
        color: #666;
    }
    .logo-img {
        height: 40px;
        width: auto;
        vertical-align: middle;
        margin-right: 1rem;
    }
    .connection-status {
        font-size: 0.9rem;
        padding: 4px 8px;
        border-radius: 4px;
        margin-left: 10px;
        display: inline-block;
    }
    .status-connected {
        background-color: #D1FAE5;
        color: #065F46;
    }
    .status-disconnected {
        background-color: #FEE2E2;
        color: #991B1B;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3, col4, col5 = st.columns([2, 1, 1, 1, 1])

    with col1:
        st.markdown(
            f"""
            <div class="nav-title">
                <img src="https://pro-source-demo.s3.us-east-1.amazonaws.com/ProSource+brandmark.png" class="logo-img" alt="ProSource Logo">
                <span>SharePoint File Name Manager <span class="nav-version">v{APP_VERSION}</span></span>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Show connection status more prominently
    connection_status = "🟢 Connected to SharePoint" if st.session_state.get('authenticated', False) else "🔴 Not Connected to SharePoint"
    status_class = "status-connected" if st.session_state.get('authenticated', False) else "status-disconnected"
    st.markdown(f'<span class="connection-status {status_class}">{connection_status}</span>', unsafe_allow_html=True)

    with col2:
        # Home button always visible
        if st.button("🏠 Libraries", key="nav_home"):
            if st.session_state.get('authenticated', False):
                st.session_state['show_setup'] = False
                st.session_state['show_credentials'] = False
                st.rerun()
            else:
                st.warning("Please connect to SharePoint first to access libraries.")
                st.session_state['show_setup'] = False
                st.session_state['show_credentials'] = False
                st.rerun()

    with col3:
        if st.button("📚 Setup Guide", key="nav_setup"):
            st.session_state['show_setup'] = True
            st.session_state['show_credentials'] = False
            st.rerun()

    with col4:
        if st.button("⚙️ Credentials", key="nav_credentials"):
            st.session_state['show_credentials'] = True
            st.session_state['show_setup'] = False
            st.rerun()

    with col5:
        if st.session_state.get('authenticated', False):
            if st.button("🔄 Rename Files", key="nav_rename"):
                st.session_state['show_setup'] = False
                st.session_state['show_credentials'] = False
                st.rerun()

def initialize_session_state():
    """Initialize session state variables"""
    if 'authenticated' not in st.session_state:
        st.session_state['authenticated'] = False
    if 'client' not in st.session_state:
        st.session_state['client'] = None
    if 'selected_files' not in st.session_state:
        st.session_state['selected_files'] = {}
    if 'rename_pattern' not in st.session_state:
        st.session_state['rename_pattern'] = "{name}{ext}"
    if 'preview_renames' not in st.session_state:
        st.session_state['preview_renames'] = []
    if 'show_setup' not in st.session_state:
        st.session_state['show_setup'] = False
    if 'show_credentials' not in st.session_state:
        st.session_state['show_credentials'] = False
    if 'problematic_files' not in st.session_state:
        st.session_state['problematic_files'] = []

def apply_rename_pattern(filename, pattern):
    """Apply rename pattern to filename"""
    try:
        # Handle "Extract Last Part" pattern type
        if pattern == "__extract_last_part__":
            # Split by various delimiters and get meaningful parts
            delimiters = [' - ', '-', '_', 'IN THE MATTER OF', 'ORDER NO', 'MOTION TO']
            name = filename
            name_without_ext, ext = os.path.splitext(name)

            # Try each delimiter
            for delimiter in delimiters:
                if delimiter in name_without_ext:
                    parts = name_without_ext.split(delimiter)
                    # Get the last meaningful part
                    last_part = parts[-1].strip()
                    # If last part is still too long, take first 60 chars
                    if len(last_part) > 60:
                        last_part = last_part[:57] + "..."
                    name = last_part + ext
                    break

            # If still too long after extraction, force truncate
            if len(name) > 128:
                name_part, ext_part = os.path.splitext(name)
                name = name_part[:120] + "..." + ext_part

            return sanitize_filename(name)

        # Extract the file extension
        name, ext = os.path.splitext(filename)

        # If filename is too long (over 128 chars), intelligently truncate it
        if len(name) > 128:
            # Calculate maximum name length (reserve space for extension)
            max_name_length = 128 - len(ext)
            name = name[:max_name_length - 3] + "..."

        # Replace placeholders in pattern
        new_name = pattern.replace('{name}', name).replace('{ext}', ext)

        # Final length check and truncation if needed
        if len(new_name) > 128:
            name_part, ext_part = os.path.splitext(new_name)
            new_name = name_part[:120] + "..." + ext_part

        # Sanitize the new name
        return sanitize_filename(new_name)
    except Exception as e:
        logger.error(f"Error in apply_rename_pattern: {str(e)}")
        return filename

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

        # Add scan button in sidebar with clear description
        if st.sidebar.button("🔍 Scan for Long Names & Paths"):
            with st.spinner("Scanning for problematic files..."):
                problematic_files = st.session_state.client.scan_for_long_paths(library_name)
                st.session_state.problematic_files = problematic_files

                if problematic_files:
                    st.sidebar.warning(f"Found {len(problematic_files)} problematic files")
                    with st.sidebar.expander("View Problematic Files"):
                        for item in problematic_files:
                            st.write(f"📄 {item['name']}")
                            if item['filename_length'] > 128:
                                st.write(f"⚠️ Filename too long: {item['filename_length']} characters")
                            if item['full_path_length'] > 256:
                                st.write(f"⚠️ Path too long: {item['full_path_length']} characters")
                            st.write(f"Full path: {item['full_path']}")
                            st.write("---")
                else:
                    st.sidebar.success("No problematic files found")
                    st.session_state.problematic_files = []

        # Add button to select all problematic files
        if st.session_state.problematic_files:
            if st.sidebar.button("Select All Problematic Files"):
                st.session_state.selected_files = {
                    item['id']: {
                        'Id': item['id'],
                        'Name': item['name'],
                        'Path': item['path']
                    } for item in st.session_state.problematic_files
                }
                st.rerun()

        # Bulk rename controls in sidebar
        st.sidebar.write("### Bulk Rename")

        # Documentation for rename patterns
        with st.sidebar.expander("📖 Pattern Type Documentation", expanded=False):
            st.markdown("""
            ### Rename Pattern Types

            #### 1. Extract Last Part
            Extracts the meaningful last part of a long filename, useful for cleaning up verbose legal document names.

            **Logic:**
            - Splits filename by common legal document delimiters
            - Takes the last meaningful part
            - Truncates to 60 characters if still too long
            - Preserves file extension

            **Examples:**
            - Before: `IN THE MATTER OF CASE 123 - MOTION TO DISMISS - EXHIBIT A.pdf`
            - After: `EXHIBIT A.pdf`

            - Before: `NOTICE OF FILING - DEFENDANT RESPONSE - CONFIDENTIAL.docx`
            - After: `CONFIDENTIAL.docx`

            #### 2. No Pattern (Keep Original)
            Maintains the original filename while ensuring it meets SharePoint's requirements.

            **Logic:**
            - Validates filename length (max 128 chars)
            - Sanitizes invalid characters
            - Preserves extension

            **Example:**
            - Before: `Original File Name.pdf`
            - After: `Original File Name.pdf`

            #### 3. Custom Pattern
            Allows custom naming patterns using placeholders.

            **Placeholders:**
            - {name} = original filename without extension
            - {ext} = original extension (including dot)

            **Examples:**
            - Pattern: `CASE123_{name}{ext}`
              - Before: `motion.pdf`
              - After: `CASE123_motion.pdf`

            - Pattern: `{name}_v1{ext}`
              - Before: `document.docx`
              - After: `document_v1.docx`

            #### 4. Add Prefix
            Prepends a specified prefix to all filenames.

            **Logic:**
            - Adds prefix before original filename
            - Preserves original name and extension
            - Validates total length

            **Examples:**
            - Prefix: `DOC_`
              - Before: `contract.pdf`
              - After: `DOC_contract.pdf`

            #### 5. Add Case Number
            Prepends a case number identifier to filenames.

            **Logic:**
            - Adds case number prefix
            - Maintains consistent format
            - Preserves original name

            **Examples:**
            - Case: `CASE123_`
              - Before: `evidence.pdf`
              - After: `CASE123_evidence.pdf`

            #### 6. Add Date Prefix
            Adds current date as a prefix in YYYYMMDD format.

            **Logic:**
            - Adds today's date as prefix
            - Format: YYYYMMDD_
            - Preserves original name

            **Examples:**
            - Before: `report.pdf`
            - After: `20250221_report.pdf`

            ### Important Notes:
            - All patterns enforce SharePoint's 128-character filename limit
            - Invalid characters are automatically removed
            - Files over length limits are intelligently truncated
            - Extensions are always preserved
            - Spaces and special characters are handled safely
            """)

        # Add pattern selection
        pattern_type = st.sidebar.radio(
            "Choose Pattern Type",
            options=[
                "Extract Last Part",
                "No Pattern (Keep Original)",
                "Custom Pattern",
                "Add Prefix",
                "Add Case Number",
                "Add Date Prefix"
            ]
        )

        # Initialize rename pattern based on selection
        if pattern_type == "Extract Last Part":
            st.session_state.rename_pattern = "__extract_last_part__"
        elif pattern_type == "No Pattern (Keep Original)":
            st.session_state.rename_pattern = "{name}{ext}"
        elif pattern_type == "Custom Pattern":
            st.session_state.rename_pattern = st.sidebar.text_input(
                "Rename Pattern",
                value=st.session_state.rename_pattern,
                help="""
                Use these placeholders:
                - {name} = original filename without extension
                - {ext} = original extension (including dot)

                Examples:
                - prefix_{name}{ext}
                - CASE123_{name}{ext}
                - {name}_v1{ext}
                """
            )
        elif pattern_type == "Add Prefix":
            prefix = st.sidebar.text_input("Enter Prefix", value="DOC_")
            st.session_state.rename_pattern = f"{prefix}{{name}}{{ext}}"
        elif pattern_type == "Add Case Number":
            case_number = st.sidebar.text_input("Enter Case Number", value="CASE123_")
            st.session_state.rename_pattern = f"{case_number}{{name}}{{ext}}"
        elif pattern_type == "Add Date Prefix":
            import datetime
            today = datetime.datetime.now().strftime("%Y%m%d_")
            st.session_state.rename_pattern = f"{today}{{name}}{{ext}}"

        # Show pattern preview
        if pattern_type not in ["Extract Last Part"]:
            st.sidebar.info(f"Pattern Preview: {st.session_state.rename_pattern}")

        # Add select all button
        if st.sidebar.button("Select All Files"):
            st.session_state.selected_files = {
                file['Id']: file
                for files in files_by_path.values()
                for file in files
            }
            logger.info(f"Selected all files: {len(st.session_state.selected_files)} total")
            st.rerun()

        # Add clear selection button
        if st.sidebar.button("Clear Selection"):
            st.session_state.selected_files = {}
            logger.info("Cleared all file selections")
            st.rerun()

        # Bulk rename button and logic
        if st.sidebar.button("Rename Selected Files"):
            logger.info(f"Attempting bulk rename with pattern: {st.session_state.rename_pattern}")
            logger.info(f"Selected files count: {len(st.session_state.selected_files)}")

            if not st.session_state.selected_files:
                st.error("Please select files to rename first")
                return

            # Create rename operations for all selected files
            rename_operations = []

            for file_id, file in st.session_state.selected_files.items():
                try:
                    current_name = file['Name']
                    new_name = apply_rename_pattern(current_name, st.session_state.rename_pattern)
                    logger.info(f"Processing file: {current_name} -> {new_name}")

                    if validate_filename(new_name):
                        rename_operations.append({
                            'old_name': current_name,
                            'new_name': new_name,
                            'file_id': file_id
                        })
                        logger.info(f"Added rename operation: {current_name} -> {new_name}")
                    else:
                        logger.warning(f"Invalid new filename for {current_name}: {new_name}")
                except Exception as e:
                    logger.error(f"Error preparing rename for {file['Name']}: {str(e)}")

            logger.info(f"Created {len(rename_operations)} rename operations")

            if rename_operations:
                with st.spinner(f"Renaming {len(rename_operations)} files..."):
                    try:
                        logger.info(f"Executing {len(rename_operations)} rename operations")
                        results = st.session_state.client.bulk_rename_files(
                            library_name,
                            rename_operations
                        )

                        # Show results summary
                        success_count = sum(1 for r in results if r['success'])
                        if success_count > 0:
                            st.success(f"Successfully renamed {success_count} out of {len(results)} files")
                            logger.info(f"Successfully renamed {success_count} out of {len(results)} files")

                        # Show errors in expandable section if any
                        failed = [r for r in results if not r['success']]
                        if failed:
                            with st.expander("Show Failed Operations"):
                                for failure in failed:
                                    error_msg = f"Failed to rename {failure['old_name']}: {failure.get('error', 'Unknown error')}"
                                    st.error(error_msg)
                                    logger.error(error_msg)

                        # Only clear selection after successful operations
                        if success_count == len(rename_operations):
                            st.session_state.selected_files = {}
                            st.session_state.problematic_files = []
                            time.sleep(1)
                            st.rerun()
                        else:
                            # Keep failed files selected for retry
                            failed_ids = {r['file_id'] for r in failed}
                            st.session_state.selected_files = {
                                file_id: file
                                for file_id, file in st.session_state.selected_files.items()
                                if file_id in failed_ids
                            }
                            st.warning("Some files were not renamed. They remain selected for retry.")
                    except Exception as e:
                        st.error(f"Error during bulk rename: {str(e)}")
                        logger.error(f"Error during bulk rename: {str(e)}")
            else:
                st.error("No valid rename operations could be created. Check the filename pattern.")
                logger.error("No valid rename operations could be created")

        # Display total selected files count
        st.sidebar.write(f"Selected: {len(st.session_state.selected_files)} files")

        # Display files grouped by folders
        for folder_name, folder_files in files_by_path.items():
            with st.expander(f"📁 {folder_name} ({len(folder_files)} files)", expanded=folder_name == 'Root'):
                for file in folder_files:
                    col1, col2 = st.columns([0.5, 3.5])

                    with col1:
                        # Checkbox for selection
                        is_selected = st.checkbox("", key=f"select_{file['Id']}",
                                                    value=file['Id'] in st.session_state.selected_files)
                        if is_selected:
                            st.session_state.selected_files[file['Id']] = file
                        elif file['Id'] in st.session_state.selected_files:
                            del st.session_state.selected_files[file['Id']]

                    with col2:
                        full_path = (file.get('ParentPath', '') + '/' + file['Name']).strip('/')
                        st.write(f"📄 {file['Name']}")
                        name_length = len(file['Name'])
                        path_length = len(full_path)

                        if name_length > 128 or path_length > 256:
                            warning_msg = []
                            if name_length > 128:
                                warning_msg.append(f"Long filename ({name_length} chars)")
                            if path_length > 256:
                                warning_msg.append(f"Long path ({path_length} chars)")
                            st.warning(" & ".join(warning_msg))

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
    show_navigation()


    if not st.session_state.authenticated:
        st.warning("⚠️ Please connect to SharePoint to access libraries and file management features.")
        authenticate()
    else:
        if st.sidebar.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.rerun()

        show_library_selector()

def authenticate():
    """Handle SharePoint authentication"""
    st.write("### Connect to SharePoint")
    st.info("""
    To get started with the SharePoint File Name Manager:
    1. Enter your SharePoint site URL below
    2. Make sure you have configured your Azure AD credentials
    3. Click Connect to access your SharePoint libraries
    """)

    # Add navigation buttons with unique keys
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⚙️ Configure Azure AD Credentials", key="auth_manage_creds"):
            st.session_state['show_credentials'] = True
            st.session_state['show_setup'] = False
            st.rerun()

    with col2:
        if st.button("📚 View Setup Guide", key="auth_setup_guide"):
            st.session_state['show_setup'] = True
            st.session_state['show_credentials'] = False
            st.rerun()

    if st.session_state.get('show_credentials', False):
        show_credentials_manager()
        return

    if st.session_state.get('show_setup', False):
        show_setup_guide()
        if st.button("← Back to Login", key="setup_back_to_login"):
            st.session_state['show_setup'] = False
            st.rerun()
        return

    with st.form("authentication_form"):
        site_url = st.text_input("SharePoint Site URL",
                                help="Enter the full SharePoint site URL (e.g., https://your-tenant.sharepoint.com/sites/your-site)")

        st.info("Make sure you have configured your Azure AD credentials before connecting.")

        submit = st.form_submit_button("🔗 Connect to SharePoint", use_container_width=True)

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

        # Add Create Test Library button
        if st.button("🧪 Create Test Library"):
            with st.spinner("Creating test library with sample data..."):
                try:
                    st.session_state.client.create_test_library()
                    st.success("Test library created successfully! Please refresh the library list.")
                    time.sleep(2)
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to create test library: {str(e)}")
                    logger.error(f"Failed to create test library: {str(e)}")

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