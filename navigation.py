import streamlit as st
from models import User

def show_navigation():
    """Display common navigation elements based on user role"""
    if 'user' not in st.session_state:
        return

    with st.sidebar:
        st.title("Navigation")
        
        # Basic navigation for all users
        if st.button("🏠 Home"):
            st.session_state['page'] = 'home'
            st.rerun()
            
        if st.button("🔐 MFA Setup"):
            st.session_state['page'] = 'mfa'
            st.rerun()
            
        # Admin controls
        user = User.query.get(st.session_state['user'])
        if user and user.is_admin:
            st.markdown("---")
            st.markdown("### Admin Controls")
            if st.button("⚙️ Admin Dashboard"):
                st.session_state['page'] = 'admin'
                st.rerun()
            
            if st.button("👥 User Management"):
                st.session_state['page'] = 'user_management'
                st.rerun()
                
            if st.button("🏢 Tenant Settings"):
                st.session_state['page'] = 'tenant_settings'
                st.rerun()
