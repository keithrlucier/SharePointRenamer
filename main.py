import streamlit as st
import logging
from sharepoint_client import SharePointClient
from utils import setup_logging, validate_filename, sanitize_filename
import time
import re
import os
from setup import show_setup_guide
from credentials import show_credentials_manager
from models import User, Tenant, ClientCredential, db
import bcrypt
import pyotp
from datetime import datetime, timedelta
from app import app
from mfa import show_mfa_setup

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

# App version
APP_VERSION = "1.0.0"

# Initialize database with admin user if not exists
def initialize_database():
    """Initialize database with admin user if not exists"""
    with app.app_context():
        try:
            # Check if any admin user exists
            admin_exists = User.query.filter_by(is_admin=True).first() is not None
            if not admin_exists:
                # Create initial admin user only if no admin exists
                admin = User(
                    email='admin@example.com',
                    is_admin=True,
                    is_active=True
                )
                admin.set_password('admin123')
                db.session.add(admin)
                db.session.commit()
                logger.info("Created initial admin user")
        except Exception as e:
            logger.error(f"Error initializing database: {str(e)}")

initialize_database()

# Initialize Flask app context
app.app_context().push()

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

    col1, col2, col3, col4, col5, col6 = st.columns([2, 1, 1, 1, 1, 1])

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
        if st.button("🏠 Libraries", key="nav_home", use_container_width=True):
            st.session_state['current_page'] = 'home'
            st.session_state['show_setup'] = False
            st.session_state['show_credentials'] = False
            if not st.session_state.get('authenticated', False):
                st.warning("Please connect to SharePoint first to access libraries.")

    with col3:
        if st.button("📚 Setup Guide", key="nav_setup", use_container_width=True):
            st.session_state['current_page'] = 'setup'
            st.session_state['show_setup'] = True
            st.session_state['show_credentials'] = False

    with col4:
        if st.button("⚙️ Credentials", key="nav_credentials", use_container_width=True):
            st.session_state['current_page'] = 'credentials'
            st.session_state['show_credentials'] = True
            st.session_state['show_setup'] = False

    with col5:
        if st.button("🔄 Rename Files", key="nav_rename", use_container_width=True):
            if st.session_state.get('authenticated', False):
                st.session_state['current_page'] = 'rename'
                st.session_state['show_setup'] = False
                st.session_state['show_credentials'] = False
            else:
                st.warning("Please connect to SharePoint first to access file renaming.")

    with col6:
        if 'user' in st.session_state:
            if st.button("🔐 2FA Setup", key="nav_2fa", use_container_width=True):
                st.session_state['current_page'] = 'mfa_setup'
                st.session_state['show_setup'] = False
                st.session_state['show_credentials'] = False

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
    if 'current_page' not in st.session_state:
        st.session_state['current_page'] = 'home'
    if 'user' not in st.session_state:
        st.session_state['user'] = None
    if 'is_admin' not in st.session_state:
        st.session_state['is_admin'] = False

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

def show_login():
    """Display login form"""
    st.write("### Login")

    with st.form("login_form"):
        email = st.text_input("Email")
        password = st.text_input("Password", type="password")

        # Only show MFA field if user exists and has MFA enabled
        if 'pending_mfa_user' in st.session_state:
            st.markdown("---")
            st.markdown("#### Two-Factor Authentication Required")
            st.info("Please enter the verification code from your authenticator app")
            mfa_code = st.text_input("Verification Code", key="mfa_input")

        submit = st.form_submit_button("Login")

        if submit:
            if email and password:
                with app.app_context():
                    user = User.query.filter_by(email=email).first()
                    if user and user.check_password(password):
                        if user.mfa_enabled:
                            if 'pending_mfa_user' not in st.session_state:
                                # First step: Show MFA input
                                st.session_state['pending_mfa_user'] = user.id
                                st.rerun()
                            else:
                                # Second step: Verify MFA
                                mfa_code = st.session_state.get("mfa_input")
                                if mfa_code and user.verify_mfa(mfa_code):
                                    st.session_state['user'] = user.id
                                    st.session_state['is_admin'] = user.is_admin
                                    st.session_state['authenticated'] = True
                                    if 'pending_mfa_user' in st.session_state:
                                        del st.session_state['pending_mfa_user']
                                    st.success("Login successful!")
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.error("Invalid verification code")
                        else:
                            # No MFA required
                            st.session_state['user'] = user.id
                            st.session_state['is_admin'] = user.is_admin
                            st.session_state['authenticated'] = True
                            st.success("Login successful!")
                            time.sleep(1)
                            st.rerun()
                    else:
                        st.error("Invalid email or password")



def show_admin_panel():
    """Display admin panel with contextual help"""
    st.write("### Admin Panel")

    # Use Flask app context for database operations
    with app.app_context():
        # Add tabs for different admin functions with help text
        st.markdown("""
        Welcome to the Admin Panel! Each section below provides different management capabilities.
        Use the tabs to navigate between different administrative functions.
        """)

        users_tab, tenants_tab, credentials_tab = st.tabs(["Users", "Tenants", "Credentials"])

        with users_tab:
            st.write("#### Manage Users")
            st.info("""
            **User Management Help**
            - Create new users with email and password
            - Set admin privileges for users
            - Enable/disable MFA for enhanced security
            - Assign users to specific tenants
            - Reset MFA if users lose access

            Note: At least one admin user must exist at all times.
            """)

            users = User.query.all()
            admin_count = User.query.filter_by(is_admin=True).count()

            # Create new user
            with st.expander("Create New User"):
                st.markdown("Add a new user to the system with specific roles and tenant assignment")
                with st.form("create_user"):
                    col1, col2 = st.columns(2)
                    with col1:
                        email = st.text_input(
                            "Email",
                            help="Enter the user's email address - this will be their login username"
                        )
                    with col2:
                        password = st.text_input(
                            "Password",
                            type="password",
                            help="Set an initial password for the user"
                        )

                    col3, col4 = st.columns(2)
                    with col3:
                        is_admin = st.checkbox(
                            "Is Admin",
                            help="Grant admin privileges - admins can manage users and tenants"
                        )
                    with col4:
                        tenant = st.selectbox(
                            "Select Tenant",
                            options=Tenant.query.all(),
                            format_func=lambda x: x.name,
                            help="Assign user to a specific tenant organization"
                        )

                    if st.form_submit_button("Create User"):
                        try:
                            user = User(
                                email=email,
                                is_admin=is_admin,
                                tenant_id=tenant.id
                            )
                            user.set_password(password)
                            db.session.add(user)
                            db.session.commit()
                            st.success("User created successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error creating user: {str(e)}")

            # List existing users
            st.write("#### Existing Users")
            st.markdown("View and manage all users in the system")

            for user in users:
                with st.expander(f"User: {user.email}"):
                    st.markdown("##### User Details")
                    st.write(f"Admin: {'Yes' if user.is_admin else 'No'}")
                    st.write(f"MFA Enabled: {'Yes' if user.mfa_enabled else 'No'}")
                    st.write(f"Tenant: {user.tenant.name if user.tenant else 'None'}")

                    # Check if this is the last admin
                    is_last_admin = user.is_admin and admin_count <= 1
                    can_delete = not is_last_admin

                    col1, col2 = st.columns([1, 3])
                    with col1:
                        st.button(
                            "Delete User",
                            key=f"delete_user_{user.id}",
                            disabled=not can_delete,
                            help="Remove this user from the system (cannot delete last admin)"
                        )

                    # Help text for admin status
                    if is_last_admin:
                        st.warning("Cannot delete the last admin user")
                    elif user.is_admin:
                        st.info("You can delete this admin user because other admin users exist")

                    st.button(
                        "Reset MFA",
                        key=f"reset_mfa_{user.id}",
                        help="Reset the user's MFA settings if they lost access to their authenticator"
                    )

        with tenants_tab:
            st.write("#### Manage Tenants")
            st.info("""
            **Tenant Management Help**
            - Create new tenant organizations
            - Set tenant subscription status
            - View users associated with each tenant
            - Manage tenant-specific settings

            Tenants are separate organizations that can have their own users and settings.
            """)

            tenants = Tenant.query.all()

            # Create new tenant
            with st.expander("Create New Tenant"):
                st.markdown("Add a new tenant organization to the system")
                with st.form("create_tenant"):
                    name = st.text_input(
                        "Tenant Name",
                        help="Enter a descriptive name for the new tenant"
                    )
                    subscription_status = st.selectbox(
                        "Subscription Status",
                        options=['trial', 'active', 'cancelled'],
                        help="Select the current subscription status for this tenant"
                    )
                    subscription_end = st.date_input(
                        "Subscription End Date",
                        value=datetime.now() + timedelta(days=30),
                        help="Specify the date when the tenant's subscription expires"
                    )

                    if st.form_submit_button("Create Tenant"):
                        try:
                            tenant = Tenant(
                                name=name,
                                subscription_status=subscription_status,
                                subscription_end=subscription_end
                            )
                            db.session.add(tenant)
                            db.session.commit()
                            st.success("Tenant created successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error creating tenant: {str(e)}")

            # List existing tenants
            st.write("#### Existing Tenants")
            st.markdown("View and manage all existing tenant organizations")

            for tenant in tenants:
                with st.expander(f"Tenant: {tenant.name}"):
                    st.markdown("##### Tenant Details")
                    st.write(f"Status: {tenant.subscription_status}")
                    st.write(f"Subscription End: {tenant.subscription_end}")
                    st.write(f"Users: {len(tenant.users)}")

                    st.button(
                        "Delete Tenant",
                        key=f"delete_tenant_{tenant.id}",
                        help="Remove this tenant organization from the system"
                    )

        with credentials_tab:
            st.write("#### Manage Client Credentials")
            st.info("""
            **Credentials Management Help**
            - Configure SharePoint connection settings
            - Manage API keys and secrets
            - Test connection status
            - View connected tenants

            These settings are optional and only needed if you want toenable SharePoint integration.
            """)

            credentials = ClientCredential.query.all()

            # Create new credentials
            with st.expander("Add New Credentials"):
                st.markdown("Add new client credentials for SharePoint integration")
                with st.form("create_credentials"):
                    tenant = st.selectbox(
                        "Select Tenant",
                        options=Tenant.query.all(),
                        format_func=lambda x: x.name,
                        key="cred_tenant",
                        help="Select the tenant this credential set belongs to"
                    )
                    client_id = st.text_input(
                        "Client ID",
                        help="Enter the Client ID from your SharePoint application registration"
                    )
                    client_secret = st.text_input(
                        "Client Secret",
                        type="password",
                        help="Enter the Client Secret from your SharePoint application registration"
                    )
                    tenant_id_azure = st.text_input(
                        "Azure Tenant ID",
                        help="Enter your Azure Tenant ID (Directory ID)"
                    )

                    if st.form_submit_button("Save Credentials"):
                        try:
                            credential = ClientCredential(
                                tenant_id=tenant.id,
                                client_id=client_id,
                                client_secret=client_secret,
                                tenant_id_azure=tenant_id_azure
                            )
                            db.session.add(credential)
                            db.session.commit()
                            st.success("Credentials saved successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error saving credentials: {str(e)}")

            # List existing credentials
            st.write("#### Existing Credentials")
            st.markdown("View and manage all existing SharePoint client credentials")

            for cred in credentials:
                with st.expander(f"Credentials for: {cred.tenant.name}"):
                    st.markdown("##### Credential Details")
                    st.write(f"Client ID: {cred.client_id}")
                    st.write(f"Azure Tenant ID: {cred.tenant_id_azure}")
                    st.write(f"Last Updated: {cred.last_updated}")

                    st.button(
                        "Delete Credentials",
                        key=f"delete_cred_{cred.id}",
                        help="Remove this credential set from the system"
                    )

def show_library_selector():
    """Display SharePoint library selector"""
    try:
        with st.spinner("Loading SharePoint libraries..."):
            libraries = st.session_state.client.get_libraries()
            if not libraries:
                st.warning("No document libraries found in this SharePoint site.")
                return

        # Add Create TestLibrary button
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


def main():
    initialize_session_state()

    # Check login state first
    if not st.session_state.get('user'):
        show_login()
        return

    # After login, show navigation
    show_navigation()

    # Handle different pages based on current_page state
    current_page = st.session_state.get('current_page', 'home')

    if current_page == 'home':
        if st.session_state.get('is_admin', False):
            # For admin users, show admin dashboard by default
            show_admin_panel()
        else:
            # For regular users, show SharePoint connection setup
            if not st.session_state.get('client'):
                st.warning("Please set up your SharePoint credentials first.")
                if st.button("Configure SharePoint Credentials"):
                    st.session_state['show_credentials'] = True
                    st.session_state['show_setup'] = False
                    st.rerun()
            elif st.session_state.get('authenticated', False):
                try:
                    # Show SharePoint libraries
                    libraries = st.session_state.client.get_libraries()
                    if libraries:
                        selected_library = st.selectbox("Select Library", libraries)
                        if selected_library:
                            show_file_manager(selected_library)
                    else:
                        st.info("No libraries found. Please check your SharePoint permissions.")
                except Exception as e:
                    st.error("Error connecting to SharePoint. Please check your credentials.")
                    logger.error(f"SharePoint connection error: {str(e)}")
                    st.session_state['authenticated'] = False
                    if st.button("Reconfigure SharePoint Connection"):
                        st.session_state['show_credentials'] = True
                        st.session_state['show_setup'] = False
                        st.rerun()
            else:
                st.warning("Please connect to SharePoint first.")
                if st.button("Configure SharePoint Connection"):
                    st.session_state['show_credentials'] = True
                    st.session_state['show_setup'] = False
                    st.rerun()

    elif current_page == 'mfa_setup':
        show_mfa_setup()

    elif current_page in ['admin', 'user_management', 'tenant_settings']:
        if st.session_state.get('is_admin', False):
            show_admin_panel()
        else:
            st.error("Access denied. Admin privileges required.")

    elif st.session_state.get('show_setup', False):
        show_setup_guide()

    elif st.session_state.get('show_credentials', False):
        show_credentials_manager()

if __name__ == "__main__":
    main()