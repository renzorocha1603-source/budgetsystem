import streamlit as st
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
import io
import re
import os
from datetime import datetime
import json

# Import the Excel fixer module
from excel_fixer import extract_pdf_data, fix_excel, DEFAULT_CONFIG

st.set_page_config(
    page_title="Budget Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================
# USER MANAGEMENT SYSTEM
# ============================================

USERS_FILE = "users.json"

DEFAULT_USERS = {
    "admin@budget.com": {
        "password": "admin123",
        "name": "Administrator",
        "role": "admin",
        "created_at": datetime.now().isoformat()
    }
}

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r') as f:
            return json.load(f)
    else:
        with open(USERS_FILE, 'w') as f:
            json.dump(DEFAULT_USERS, f, indent=2)
        return DEFAULT_USERS

def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f, indent=2)

def authenticate_user(email, password):
    users = load_users()
    if email in users and users[email]["password"] == password:
        return users[email]
    return None

def create_user(email, password, name, role="user"):
    users = load_users()
    if email in users:
        return False, "User already exists"
    users[email] = {
        "password": password,
        "name": name,
        "role": role,
        "created_at": datetime.now().isoformat()
    }
    save_users(users)
    return True, "User created successfully"

def delete_user(email):
    users = load_users()
    if email == "admin@budget.com":
        return False, "Cannot delete admin user"
    if email in users:
        del users[email]
        save_users(users)
        return True, "User deleted successfully"
    return False, "User not found"

def update_user_password(email, new_password):
    users = load_users()
    if email in users:
        users[email]["password"] = new_password
        save_users(users)
        return True, "Password updated successfully"
    return False, "User not found"

# ============================================
# CUSTOM CSS
# ============================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #070E1A;
        color: #C8D6E8;
    }

    .login-container {
        max-width: 420px;
        margin: 0 auto;
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: #0D1B2E;
        border-radius: 20px;
        border: 1px solid #1A3050;
        padding: 2.5rem;
        box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5);
        width: 90%;
    }
    
    .login-header {
        text-align: center;
        margin-bottom: 2rem;
    }
    
    .login-header h1 {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.8rem;
        font-weight: 600;
        color: #00D4FF;
        letter-spacing: 0.08em;
        margin-bottom: 0.5rem;
    }
    
    .login-header .subtitle {
        font-size: 0.7rem;
        color: #4A6E8A;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    .db-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1rem 1.75rem;
        background: linear-gradient(90deg, #0D1B2E 0%, #0A2040 60%, #071428 100%);
        border-radius: 10px;
        border: 1px solid #1A3050;
        border-top: 2px solid #00D4FF;
        margin-bottom: 1.5rem;
    }

    .db-header-left h1 {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.2rem;
        font-weight: 600;
        color: #F0F4FA;
        margin: 0;
    }

    .card {
        background: #0D1B2E;
        border-radius: 8px;
        border: 1px solid #172A42;
        padding: 1rem 1.25rem;
        margin-bottom: 1rem;
    }

    .card-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.65rem;
        font-weight: 500;
        color: #00D4FF;
        letter-spacing: 0.15em;
        text-transform: uppercase;
        margin-bottom: 0.8rem;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #172A42;
    }

    .scrollable-container {
        max-height: 300px;
        overflow-y: auto;
        padding-right: 8px;
    }
    
    .chat-wrapper {
        max-height: 400px;
        overflow-y: auto;
        padding-right: 8px;
    }
    
    .user-message {
        background: linear-gradient(135deg, #0E2A45 0%, #112E50 100%);
        border: 1px solid #1A4470;
        border-right: 2px solid #00D4FF;
        padding: 0.6rem 0.9rem;
        border-radius: 8px 2px 8px 8px;
        margin-bottom: 0.7rem;
        max-width: 82%;
        margin-left: auto;
        font-size: 0.8rem;
    }

    .assistant-message {
        background: #0A1628;
        border: 1px solid #172A42;
        border-left: 2px solid #C9A84C;
        padding: 0.6rem 0.9rem;
        border-radius: 2px 8px 8px 8px;
        margin-bottom: 0.7rem;
        max-width: 88%;
        font-size: 0.8rem;
    }

    .metric-row {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 0.6rem;
        margin-top: 0.5rem;
    }

    .metric-block {
        background: #070E1A;
        border: 1px solid #172A42;
        border-top: 2px solid;
        border-radius: 6px;
        padding: 0.6rem 0.8rem;
    }

    .metric-block.cyan { border-top-color: #00D4FF; }
    .metric-block.gold { border-top-color: #C9A84C; }
    .metric-block.teal { border-top-color: #00C9A7; }

    .metric-value {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1rem;
        font-weight: 600;
    }
    .metric-value.cyan { color: #00D4FF; }
    .metric-value.gold { color: #C9A84C; }
    .metric-value.teal { color: #00C9A7; }

    .rule {
        height: 1px;
        background: linear-gradient(to right, transparent, #1A3050, transparent);
        margin: 0.8rem 0;
    }

    .db-footer {
        text-align: center;
        padding: 1rem 0 0.3rem;
        font-size: 0.55rem;
        color: #253A52;
        border-top: 1px solid #0D1B2E;
        margin-top: 0.8rem;
    }

    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ============================================
# LOGIN PAGE
# ============================================

def login_page():
    st.markdown("""
    <div class="login-container">
        <div class="login-header">
            <h1>▸ BUDGET SYSTEM</h1>
            <div class="subtitle">by Only Solutions Inc.</div>
        </div>
    """, unsafe_allow_html=True)
    
    with st.form("login_form", clear_on_submit=False):
        email = st.text_input("EMAIL", placeholder="admin@budget.com", label_visibility="collapsed")
        password = st.text_input("PASSWORD", type="password", placeholder="••••••••", label_visibility="collapsed")
        submitted = st.form_submit_button("▸  LOGIN", use_container_width=True)
        
        if submitted:
            user = authenticate_user(email, password)
            if user:
                st.session_state.authenticated = True
                st.session_state.user_email = email
                st.session_state.user_name = user["name"]
                st.session_state.user_role = user["role"]
                st.rerun()
            else:
                st.error("✗ Invalid email or password")
    
    st.markdown("""
        <div class="login-footer">🔐 Authorized Personnel Only</div>
    </div>
    """, unsafe_allow_html=True)

# ============================================
# USER MANAGEMENT (Admin only)
# ============================================

def user_management_panel():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">👥 USER MANAGEMENT</div>', unsafe_allow_html=True)
    st.markdown('<div class="scrollable-container">', unsafe_allow_html=True)
    
    with st.expander("➕ CREATE NEW USER", expanded=False):
        with st.form("create_user_form"):
            new_email = st.text_input("Email", placeholder="user@example.com")
            new_name = st.text_input("Full Name", placeholder="John Doe")
            new_password = st.text_input("Password", type="password")
            confirm = st.text_input("Confirm Password", type="password")
            
            if st.form_submit_button("Create User"):
                if not all([new_email, new_name, new_password]):
                    st.error("All fields required")
                elif new_password != confirm:
                    st.error("Passwords don't match")
                else:
                    success, msg = create_user(new_email, new_password, new_name)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
    
    st.markdown("### 📋 USER LIST")
    users = load_users()
    for email, data in users.items():
        col1, col2, col3, col4 = st.columns([2, 2, 1, 0.8])
        col1.write(f"**{data['name']}**")
        col2.write(f"<small>{email}</small>", unsafe_allow_html=True)
        col3.write("🔴 ADMIN" if data['role'] == 'admin' else "🟢 USER")
        if data['role'] != 'admin':
            if col4.button("🗑️", key=f"del_{email}"):
                delete_user(email)
                st.rerun()
        st.markdown("---")
    
    with st.expander("🔑 RESET PASSWORD", expanded=False):
        users_list = load_users()
        selected = st.selectbox("Select User", [f"{d['name']} ({e})" for e, d in users_list.items()])
        if selected:
            email = selected.split("(")[-1].replace(")", "")
            with st.form("reset_form"):
                new_pass = st.text_input("New Password", type="password")
                confirm_pass = st.text_input("Confirm", type="password")
                if st.form_submit_button("Reset"):
                    if new_pass and new_pass == confirm_pass:
                        update_user_password(email, new_pass)
                        st.success("Password updated")
                    else:
                        st.error("Passwords don't match")
    
    st.markdown('</div></div>', unsafe_allow_html=True)

# ============================================
# MAIN DASHBOARD
# ============================================

def main_dashboard():
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    role_badge = "👑 ADMIN" if st.session_state.user_role == 'admin' else "👤 USER"
    
    st.markdown(f"""
    <div class="db-header">
        <div class="db-header-left">
            <h1>▸ BUDGET DASHBOARD</h1>
            <p>{st.session_state.user_name} | {role_badge}</p>
        </div>
        <div class="db-header-right">
            <span>{now_str}</span>
            <span>🟢 LIVE</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("---")
        st.caption(f"📧 {st.session_state.user_email}")
        st.markdown("---")
        
        if st.button("🚪 LOGOUT", use_container_width=True):
            for key in ['authenticated', 'user_email', 'user_name', 'user_role', 'messages', 'uploaded_files', 'extracted_revenue', 'fixed_excel']:
                st.session_state[key] = None if key != 'uploaded_files' else False if key == 'uploaded_files' else [] if key == 'messages' else {} if key == 'extracted_revenue' else False
            st.rerun()
        
        if st.session_state.user_role == 'admin':
            user_management_panel()
    
    # Initialize session
    if 'messages' not in st.session_state:
        st.session_state.messages = []
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = False
    if 'extracted_revenue' not in st.session_state:
        st.session_state.extracted_revenue = {}
    if 'fixed_excel' not in st.session_state:
        st.session_state.fixed_excel = None
    
    # Mistral AI setup
    MISTRAL_API_KEY = "YOUR_NEW_MISTRAL_API_KEY_HERE"
    try:
        from mistralai import Mistral
        client = Mistral(api_key=MISTRAL_API_KEY)
        mistral_available = True
    except:
        client = None
        mistral_available = False
    
    # Main layout
    col_left, col_right = st.columns([1, 1.5])
    
    # LEFT: Chat
    with col_left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">🤖 AI ASSISTANT</div>', unsafe_allow_html=True)
        
        col_clear1, col_clear2 = st.columns([3, 1])
        with col_clear2:
            if st.button("🗑️ CLEAR"):
                st.session_state.messages = []
                st.rerun()
        
        st.markdown('<div class="rule"></div>', unsafe_allow_html=True)
        st.markdown('<div class="chat-wrapper">', unsafe_allow_html=True)
        
        for msg in st.session_state.messages:
            if msg['role'] == 'user':
                st.markdown(f'<div class="user-message">👤 {msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="assistant-message">🤖 {msg["content"]}</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        user_input = st.chat_input("Ask about budget...")
        if user_input and mistral_available:
            st.session_state.messages.append({'role': 'user', 'content': user_input})
            try:
                response = client.chat.complete(
                    model="mistral-small-latest",
                    messages=[{"role": "user", "content": user_input}]
                )
                answer = response.choices[0].message.content
                st.session_state.messages.append({'role': 'assistant', 'content': answer})
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # RIGHT: File Management
    with col_right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📁 FILE MANAGEMENT</div>', unsafe_allow_html=True)
        
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            excel_file = st.file_uploader("Excel Template", type=["xlsx"], key="excel")
        with col_u2:
            pdf_file = st.file_uploader("PDF Report", type=["pdf"], key="pdf")
        
        if excel_file and pdf_file and not st.session_state.uploaded_files:
            with st.spinner("Processing..."):
                revenue_data = extract_pdf_data(pdf_file)
                st.session_state.extracted_revenue = revenue_data
                st.session_state.uploaded_files = True
                st.session_state.excel_data = excel_file
                st.success("✅ Files processed!")
        
        if st.session_state.extracted_revenue:
            rev = st.session_state.extracted_revenue
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-block cyan">Transient<div class="metric-value cyan">${rev['transient']:,.0f}</div></div>
                <div class="metric-block gold">Monthly<div class="metric-value gold">${rev['monthly']:,.0f}</div></div>
                <div class="metric-block teal">Total<div class="metric-value teal">${rev['total']:,.0f}</div></div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.uploaded_files:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">⚙️ PROCESSING</div>', unsafe_allow_html=True)
            
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                parking_name = st.selectbox("Parking", ["CMO142 (LUNA)", "CMO143", "CMO144"])
            with col_c2:
                parking_type = st.selectbox("Type", ["SC (School)", "RG (Tourism)"])
            
            supervisor_hours = st.number_input("Supervisor Hours/Day", 0, 24, 1)
            
            if st.button("▸  RUN BUDGET PROCESS", use_container_width=True):
                with st.spinner("Fixing Excel..."):
                    config = {
                        "parking_name": parking_name,
                        "parking_type": parking_type,
                        "supervisor_hours": supervisor_hours
                    }
                    output, log = fix_excel(st.session_state.excel_data, st.session_state.extracted_revenue, config)
                    st.session_state.fixed_excel = output
                    
                    # Show what was updated
                    with st.expander("📋 Update Log"):
                        for item in log:
                            st.write(item)
                    st.success("✅ Budget processed!")
            
            if st.session_state.fixed_excel:
                parking_clean = parking_name.replace(" ", "_").replace("(", "").replace(")", "")
                st.download_button(
                    label="↓ DOWNLOAD BUDGET FILE",
                    data=st.session_state.fixed_excel,
                    file_name=f"{parking_clean}_budget_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    use_container_width=True
                )
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="db-footer">Budget Dashboard | Only Solutions Inc.</div>', unsafe_allow_html=True)

# ============================================
# RUN APP
# ============================================

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if st.session_state.authenticated:
    main_dashboard()
else:
    login_page()
