import streamlit as st
from models import User
from app import app

def show_navigation():
    """Display common navigation elements based on user role"""
    with st.sidebar:
        # Logout button at the top
        if st.button("🚪 Logout"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        st.title("Navigation")

        # Basic navigation for all users
        if st.button("🏠 Home", key="nav_home"):
            st.session_state['current_page'] = 'home'
            st.session_state['show_setup'] = False
            st.session_state['show_credentials'] = False
            st.rerun()

        if st.button("🔐 MFA Setup", key="nav_mfa"):
            st.session_state['current_page'] = 'mfa_setup'
            st.session_state['show_setup'] = False
            st.session_state['show_credentials'] = False
            st.rerun()

        # Admin controls
        if 'user' in st.session_state:
            with app.app_context():
                user = User.query.get(st.session_state['user'])
                if user and user.is_admin:
                    st.markdown("---")
                    st.markdown("### Admin Controls")

                    if st.button("⚙️ Admin Dashboard", key="nav_admin"):
                        st.session_state['current_page'] = 'admin'
                        st.session_state['show_setup'] = False
                        st.session_state['show_credentials'] = False
                        st.rerun()

                    if st.button("👥 User Management", key="nav_users"):
                        st.session_state['current_page'] = 'user_management'
                        st.session_state['show_setup'] = False
                        st.session_state['show_credentials'] = False
                        st.rerun()

                    if st.button("🏢 Tenant Settings", key="nav_tenant"):
                        st.session_state['current_page'] = 'tenant_settings'
                        st.session_state['show_setup'] = False
                        st.session_state['show_credentials'] = False
                        st.rerun()