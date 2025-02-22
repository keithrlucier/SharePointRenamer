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
        if st.button("🏠 Home"):
            st.session_state['current_page'] = 'home'
            st.session_state['show_setup'] = False
            st.session_state['show_credentials'] = False
            st.rerun()

        if st.button("🔐 MFA Setup"):
            st.session_state['current_page'] = 'mfa_setup'
            st.session_state['show_setup'] = False
            st.session_state['show_credentials'] = False
            st.rerun()

        # Admin controls
        if st.session_state.get('is_admin', False):
            st.markdown("---")
            st.markdown("### Admin Controls")

            col1, col2 = st.columns(2)

            with col1:
                if st.button("👥 Users"):
                    st.session_state['current_page'] = 'user_management'
                    st.session_state['show_setup'] = False
                    st.session_state['show_credentials'] = False
                    st.rerun()

            with col2:
                if st.button("🏢 Tenants"):
                    st.session_state['current_page'] = 'tenant_settings'
                    st.session_state['show_setup'] = False
                    st.session_state['show_credentials'] = False
                    st.rerun()

            # SharePoint options for admins (optional)
            st.markdown("### SharePoint Options")
            if st.button("📚 Setup Guide"):
                st.session_state['current_page'] = 'setup'
                st.session_state['show_setup'] = True
                st.session_state['show_credentials'] = False
                st.rerun()

            if st.button("⚙️ Credentials"):
                st.session_state['current_page'] = 'credentials'
                st.session_state['show_credentials'] = True
                st.session_state['show_setup'] = False
                st.rerun()