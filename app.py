import streamlit as st
import pdfplumber
from openpyxl import load_workbook
import io
import re
import os
import json
from datetime import datetime

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Budget System · Only Solutions",
    page_icon="▸",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
#  TRANSLATIONS
# ─────────────────────────────────────────────
TEXTS = {
    "en": {
        "brand":           "BUDGET SYSTEM",
        "brand_sub":       "Only Solutions Inc.",
        "email_lbl":       "Email address",
        "pass_lbl":        "Password",
        "login_btn":       "Sign in",
        "logout_btn":      "Sign out",
        "wrong_creds":     "Incorrect email or password.",
        "dashboard_title": "Dashboard",
        "ai_title":        "AI Assistant",
        "clear_chat":      "Clear",
        "chat_hint":       "Ask about budget, forecasts, or calculations…",
        "no_msgs":         "No messages yet.",
        "files_title":     "File Upload",
        "excel_lbl":       "Excel Template (.xlsx)",
        "pdf_lbl":         "PDF Report (.pdf)",
        "processing":      "Processing files…",
        "files_ok":        "Files extracted — review metrics below.",
        "metrics_title":   "Extracted Metrics",
        "transient":       "Transient",
        "monthly":         "Monthly",
        "total":           "Total",
        "config_title":    "Workflow Configuration",
        "parking_lbl":     "Parking location",
        "type_lbl":        "Parking type",
        "hours_lbl":       "Supervisor hours / day",
        "run_btn":         "▸  Run Workflow",
        "running":         "Running workflow…",
        "run_ok":          "Budget processed — file ready.",
        "dl_btn":          "↓  Download Budget File",
        "admin_title":     "Admin Panel",
        "new_user_title":  "Create user",
        "nm_lbl":          "Full name",
        "new_email_lbl":   "Email",
        "new_pass_lbl":    "Password",
        "role_lbl":        "Role",
        "create_btn":      "Create user",
        "user_exists":     "A user with that email already exists.",
        "user_created":    "User created successfully.",
        "users_title":     "User list",
        "delete_btn":      "Remove",
        "del_ok":          "User removed.",
        "no_del_admin":    "The admin account cannot be removed.",
        "reset_title":     "Reset password",
        "select_user":     "Select user",
        "new_pw_lbl":      "New password",
        "reset_btn":       "Reset password",
        "reset_ok":        "Password updated.",
        "footer":          "Budget System · Only Solutions Inc.",
        "lang_en":         "EN",
        "lang_fr":         "FR",
    },
    "fr": {
        "brand":           "SYSTÈME BUDGÉTAIRE",
        "brand_sub":       "Only Solutions Inc.",
        "email_lbl":       "Adresse courriel",
        "pass_lbl":        "Mot de passe",
        "login_btn":       "Se connecter",
        "logout_btn":      "Se déconnecter",
        "wrong_creds":     "Courriel ou mot de passe incorrect.",
        "dashboard_title": "Tableau de bord",
        "ai_title":        "Assistant IA",
        "clear_chat":      "Effacer",
        "chat_hint":       "Posez une question sur le budget, les prévisions…",
        "no_msgs":         "Aucun message.",
        "files_title":     "Téléverser les fichiers",
        "excel_lbl":       "Modèle Excel (.xlsx)",
        "pdf_lbl":         "Rapport PDF (.pdf)",
        "processing":      "Traitement des fichiers…",
        "files_ok":        "Fichiers extraits — vérifiez les métriques ci-dessous.",
        "metrics_title":   "Métriques extraites",
        "transient":       "Transitoire",
        "monthly":         "Mensuel",
        "total":           "Total",
        "config_title":    "Configuration du workflow",
        "parking_lbl":     "Stationnement",
        "type_lbl":        "Type de stationnement",
        "hours_lbl":       "Heures superviseur / jour",
        "run_btn":         "▸  Exécuter le workflow",
        "running":         "Exécution en cours…",
        "run_ok":          "Budget traité — fichier prêt.",
        "dl_btn":          "↓  Télécharger le fichier",
        "admin_title":     "Panneau admin",
        "new_user_title":  "Créer un utilisateur",
        "nm_lbl":          "Nom complet",
        "new_email_lbl":   "Courriel",
        "new_pass_lbl":    "Mot de passe",
        "role_lbl":        "Rôle",
        "create_btn":      "Créer l'utilisateur",
        "user_exists":     "Un utilisateur avec ce courriel existe déjà.",
        "user_created":    "Utilisateur créé avec succès.",
        "users_title":     "Liste des utilisateurs",
        "delete_btn":      "Supprimer",
        "del_ok":          "Utilisateur supprimé.",
        "no_del_admin":    "Le compte administrateur ne peut pas être supprimé.",
        "reset_title":     "Réinitialiser le mot de passe",
        "select_user":     "Sélectionner un utilisateur",
        "new_pw_lbl":      "Nouveau mot de passe",
        "reset_btn":       "Réinitialiser",
        "reset_ok":        "Mot de passe mis à jour.",
        "footer":          "Système budgétaire · Only Solutions Inc.",
        "lang_en":         "EN",
        "lang_fr":         "FR",
    },
}

# ─────────────────────────────────────────────
#  USER MANAGEMENT
# ─────────────────────────────────────────────
USERS_FILE = "users.json"

DEFAULT_USERS = {
    "admin@onlys.com": {
        "password": "12345",
        "name":     "Administrator",
        "role":     "admin",
        "created":  datetime.now().isoformat(),
    }
}

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r") as f:
            return json.load(f)
    save_users(DEFAULT_USERS)
    return DEFAULT_USERS

def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)

def authenticate(email, password):
    users = load_users()
    u = users.get(email)
    return u if u and u["password"] == password else None

def create_user(email, name, password, role="user"):
    users = load_users()
    if email in users:
        return False, "User already exists"
    users[email] = {"password": password, "name": name, "role": role, "created": datetime.now().isoformat()}
    save_users(users)
    return True, "User created successfully"

def delete_user(email):
    if email == "admin@onlys.com":
        return False, "Cannot delete admin"
    users = load_users()
    if email in users:
        del users[email]
        save_users(users)
        return True, "User deleted"
    return False, "User not found"

def reset_password(email, new_pw):
    users = load_users()
    if email in users:
        users[email]["password"] = new_pw
        save_users(users)
        return True, "Password updated"
    return False, "User not found"

# ─────────────────────────────────────────────
#  PDF / EXCEL PROCESSING
# ─────────────────────────────────────────────
def extract_pdf_data(pdf_file):
    result = {"transient": 0, "monthly": 0, "total": 0}
    with pdfplumber.open(pdf_file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() or ""
    nums = re.findall(r'\$?([\d,]+\.?\d*)', text)
    result["total"] = float(nums[0].replace(",", "")) if nums else 40000
    for line in text.split("\n"):
        ll = line.lower()
        ns = re.findall(r'\$?([\d,]+\.?\d*)', line)
        if "transient" in ll and ns:
            result["transient"] = float(ns[0].replace(",", ""))
        elif "monthly" in ll and ns:
            result["monthly"] = float(ns[0].replace(",", ""))
    if result["transient"] == 0:
        result["transient"] = result["total"] * 0.6
    if result["monthly"] == 0:
        result["monthly"] = result["total"] * 0.4
    return result

def run_excel_workflow(excel_bytes, revenue, parking_type, supervisor_hours):
    wb = load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    log = []

    def setcell(addr, val):
        try:
            ws[addr] = val
            ws[addr].number_format = "#,##0.00"
        except Exception:
            pass

    setcell("K17", revenue["transient"])
    setcell("K18", revenue["monthly"])
    setcell("K26", revenue["total"])
    log.append(f"✓ Revenue cells updated (K17=${revenue['transient']:,.0f}, K18=${revenue['monthly']:,.0f}, K26=${revenue['total']:,.0f})")

    if "School" in parking_type:
        mult = [1.2, 1.2, 1.2, 1.1, 0.8, 0.8, 0.8, 0.8, 1.1, 1.1, 1.2, 1.2]
    else:
        mult = [0.8, 0.8, 0.9, 0.9, 1.3, 1.3, 1.3, 1.2, 1.0, 1.0, 0.8, 0.8]

    base = revenue["total"] / 12
    for i, m in enumerate(mult):
        setcell(f"K{20 + i}", base * m)
    log.append("✓ Monthly projections applied (K20–K31)")

    sup_cost = 25 * supervisor_hours * 30
    setcell("K30", sup_cost)
    log.append(f"✓ Supervisor cost set to ${sup_cost:,.2f} / month")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, log

# ─────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background: #060D18 !important;
    color: #B8CCDE;
}
.main { padding: 0 !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }
#MainMenu, footer, header { visibility: hidden; }
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: #0A1525; }
::-webkit-scrollbar-thumb { background: #1A3050; border-radius: 4px; }

.login-container {
    max-width: 420px;
    margin: 0 auto;
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    background: #0D1B2E;
    border-radius: 16px;
    border: 1px solid #1A3050;
    padding: 2rem;
    width: 90%;
}
.login-header {
    text-align: center;
    margin-bottom: 1.5rem;
}
.login-header h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.5rem;
    color: #00D4FF;
    margin-bottom: 0.25rem;
}
.login-header .subtitle {
    font-size: 0.65rem;
    color: #4A6E8A;
}

.db-header {
    background: #0D1B2E;
    padding: 0.8rem 1.5rem;
    border-radius: 8px;
    border-bottom: 2px solid #00D4FF;
    margin-bottom: 1rem;
    display: flex;
    justify-content: space-between;
}
.scard {
    background: #0C1929;
    border: 1px solid #172A42;
    border-radius: 8px;
    padding: 1rem 1.25rem;
    margin-bottom: 1rem;
}
.scard-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.63rem;
    font-weight: 500;
    color: #00D4FF;
    text-transform: uppercase;
    margin-bottom: 0.8rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid #172A42;
}
.metric-row {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 0.6rem;
    margin-top: 0.75rem;
}
.metric-block {
    background: #070E1A;
    border: 1px solid #172A42;
    border-radius: 6px;
    padding: 0.6rem 0.8rem;
    text-align: center;
}
.metric-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.1rem;
    font-weight: 600;
    color: #00D4FF;
}
.chat-scroll {
    max-height: 350px;
    overflow-y: auto;
    padding: 0.25rem 0;
}
.bubble-user {
    background: #0E2A45;
    border-right: 2px solid #00D4FF;
    padding: 0.6rem 0.85rem;
    border-radius: 8px 2px 8px 8px;
    margin-bottom: 0.65rem;
    margin-left: auto;
    max-width: 80%;
}
.bubble-bot {
    background: #070E1A;
    border-left: 2px solid #C9A84C;
    padding: 0.6rem 0.85rem;
    border-radius: 2px 8px 8px 8px;
    margin-bottom: 0.65rem;
    max-width: 80%;
}
.log-line {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    color: #00C9A7;
    padding: 0.2rem 0;
}
.hr {
    height: 1px;
    background: #172A42;
    margin: 0.8rem 0;
}
.db-footer {
    text-align: center;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.55rem;
    color: #1A3050;
    padding: 1rem 0 0.5rem;
    border-top: 1px solid #0E1E30;
    margin-top: 0.8rem;
}
.stButton > button {
    background: #0A2240 !important;
    color: #00D4FF !important;
    border: 1px solid #1A4460 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.72rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    border-radius: 5px !important;
    padding: 0.4rem 1rem !important;
    width: 100% !important;
}
.stButton > button:hover {
    background: #0D2E58 !important;
    border-color: #00D4FF88 !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(90deg, #003D5C 0%, #005C80 100%) !important;
    border-color: #007AAA !important;
}
.stDownloadButton > button {
    background: #071A10 !important;
    color: #00C9A7 !important;
    border: 1px solid #0D4030 !important;
}
section[data-testid="stSidebar"] {
    background: #0A1525 !important;
    border-right: 1px solid #172A42 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────────
_defaults = {
    "authenticated": False,
    "user_email":    "",
    "user_name":     "",
    "user_role":     "",
    "lang":          "en",
    "messages":      [],
    "excel_bytes":   None,
    "extracted_rev": {},
    "files_ready":   False,
    "fixed_excel":   None,
    "workflow_log":  [],
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

def T(key):
    return TEXTS[st.session_state.lang].get(key, key)

def do_logout():
    for k, v in _defaults.items():
        st.session_state[k] = v
    st.rerun()

# ─────────────────────────────────────────────
#  MISTRAL SETUP
# ─────────────────────────────────────────────
MISTRAL_API_KEY = "em5oqjSdA1Nus9iUpa1MNAJtQA4YfCtK"

try:
    from mistralai import Mistral
    _mistral_client = Mistral(api_key=MISTRAL_API_KEY)
    MISTRAL_OK = True
except Exception:
    _mistral_client = None
    MISTRAL_OK = False

def ask_mistral(messages_history):
    if not MISTRAL_OK:
        return "⚠️ Mistral AI not configured. Please add your API key."
    try:
        sys_msg = {"role": "system", "content": "You are a professional budget assistant for parking operations. Be concise and actionable."}
        payload = [sys_msg] + [{"role": m["role"], "content": m["content"]} for m in messages_history]
        resp = _mistral_client.chat.complete(model="mistral-small-latest", messages=payload)
        return resp.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

# ─────────────────────────────────────────────
#  LOGIN PAGE
# ─────────────────────────────────────────────
def page_login():
    t = TEXTS[st.session_state.lang]

    lc1, lc2, lc3 = st.columns([8, 0.6, 0.6])
    with lc2:
        if st.button("EN"):
            st.session_state.lang = "en"
            st.rerun()
    with lc3:
        if st.button("FR"):
            st.session_state.lang = "fr"
            st.rerun()

    st.markdown(f"""
    <div class="login-container">
        <div class="login-header">
            <h1>▸ {t['brand']}</h1>
            <div class="subtitle">{t['brand_sub']}</div>
        </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        email = st.text_input(t["email_lbl"], placeholder="admin@onlys.com")
        password = st.text_input(t["pass_lbl"], type="password", placeholder="12345")
        submitted = st.form_submit_button(t["login_btn"], use_container_width=True)

        if submitted:
            user = authenticate(email, password)
            if user:
                st.session_state.authenticated = True
                st.session_state.user_email = email
                st.session_state.user_name = user["name"]
                st.session_state.user_role = user["role"]
                st.rerun()
            else:
                st.error(t["wrong_creds"])

    st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ADMIN SIDEBAR
# ─────────────────────────────────────────────
def render_admin_sidebar():
    t = TEXTS[st.session_state.lang]
    with st.sidebar:
        st.markdown("---")
        st.markdown(f"### {t['admin_title']}")
        
        with st.expander("➕ " + t["new_user_title"], expanded=False):
            with st.form("create_user_form"):
                new_name = st.text_input(t["nm_lbl"], key="cu_name")
                new_email = st.text_input(t["new_email_lbl"], key="cu_email")
                new_pass = st.text_input(t["new_pass_lbl"], type="password", key="cu_pw")
                role = st.selectbox(t["role_lbl"], ["user", "admin"], key="cu_role")
                if st.form_submit_button(t["create_btn"]):
                    if new_email and new_name and new_pass:
                        ok, msg = create_user(new_email, new_name, new_pass, role)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("All fields required")
        
        with st.expander("📋 " + t["users_title"], expanded=False):
            users = load_users()
            for ue, ud in users.items():
                col1, col2 = st.columns([3, 1])
                with col1:
                    badge = "👑 ADMIN" if ud["role"] == "admin" else "👤 USER"
                    st.markdown(f"**{ud['name']}**  \n`{ue}`  \n*{badge}*")
                with col2:
                    if ud["role"] != "admin":
                        if st.button("🗑️", key=f"del_{ue}"):
                            delete_user(ue)
                            st.rerun()
                st.markdown("---")
        
        with st.expander("🔑 " + t["reset_title"], expanded=False):
            users = load_users()
            sel = st.selectbox(t["select_user"], list(users.keys()))
            npw = st.text_input(t["new_pw_lbl"], type="password")
            if st.button(t["reset_btn"]):
                if sel and npw:
                    ok, msg = reset_password(sel, npw)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)

# ─────────────────────────────────────────────
#  MAIN DASHBOARD
# ─────────────────────────────────────────────
def page_dashboard():
    t = TEXTS[st.session_state.lang]

    if st.session_state.user_role == "admin":
        render_admin_sidebar()

    # Header
    st.markdown(f"""
    <div class="db-header">
        <div><strong>{t['brand']}</strong><br><small>{st.session_state.user_name} | {st.session_state.user_role.upper()}</small></div>
        <div><small>{datetime.now().strftime('%Y-%m-%d %H:%M')}</small></div>
    </div>
    """, unsafe_allow_html=True)

    # Language and logout buttons
    col_l1, col_l2, col_l3, col_l4 = st.columns([6, 0.8, 0.8, 1.2])
    with col_l2:
        if st.button("EN"):
            st.session_state.lang = "en"
            st.rerun()
    with col_l3:
        if st.button("FR"):
            st.session_state.lang = "fr"
            st.rerun()
    with col_l4:
        if st.button(t["logout_btn"]):
            do_logout()

    st.markdown("---")

    # Main layout - 3 columns
    col_chat, col_files, col_workflow = st.columns([1, 1.2, 1])

    # COLUMN 1: AI CHAT
    with col_chat:
        st.markdown(f'<div class="scard"><div class="scard-title">{t["ai_title"]}</div>', unsafe_allow_html=True)
        
        if st.button(t["clear_chat"], key="clear_chat"):
            st.session_state.messages = []
            st.rerun()
        
        st.markdown('<div class="chat-scroll">', unsafe_allow_html=True)
        if not st.session_state.messages:
            st.info(t["no_msgs"])
        for msg in st.session_state.messages:
            if msg["role"] == "user":
                st.markdown(f'<div class="bubble-user"><strong>You:</strong><br>{msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="bubble-bot"><strong>AI:</strong><br>{msg["content"]}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        user_input = st.chat_input(t["chat_hint"])
        if user_input:
            st.session_state.messages.append({"role": "user", "content": user_input})
            reply = ask_mistral(st.session_state.messages)
            st.session_state.messages.append({"role": "assistant", "content": reply})
            st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

    # COLUMN 2: FILE UPLOAD
    with col_files:
        st.markdown(f'<div class="scard"><div class="scard-title">{t["files_title"]}</div>', unsafe_allow_html=True)
        
        excel_file = st.file_uploader(t["excel_lbl"], type=["xlsx"], key="excel_up")
        pdf_file = st.file_uploader(t["pdf_lbl"], type=["pdf"], key="pdf_up")
        
        if excel_file and pdf_file and not st.session_state.files_ready:
            with st.spinner(t["processing"]):
                rev = extract_pdf_data(pdf_file)
                st.session_state.extracted_rev = rev
                st.session_state.excel_bytes = excel_file.read()
                st.session_state.files_ready = True
                st.session_state.workflow_log = []
            st.success(t["files_ok"])
            st.rerun()
        
        if st.session_state.extracted_rev:
            rev = st.session_state.extracted_rev
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-block"><div class="metric-value">${rev['transient']:,.0f}</div><div>{t['transient']}</div></div>
                <div class="metric-block"><div class="metric-value">${rev['monthly']:,.0f}</div><div>{t['monthly']}</div></div>
                <div class="metric-block"><div class="metric-value">${rev['total']:,.0f}</div><div>{t['total']}</div></div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # COLUMN 3: WORKFLOW
    with col_workflow:
        if st.session_state.files_ready:
            st.markdown(f'<div class="scard"><div class="scard-title">{t["config_title"]}</div>', unsafe_allow_html=True)
            
            parking = st.selectbox(t["parking_lbl"], ["CMO142 (LUNA)", "CMO143", "CMO144"])
            p_type = st.selectbox(t["type_lbl"], ["SC (School)", "RG (Tourism)"])
            hours = st.number_input(t["hours_lbl"], min_value=0, max_value=24, value=1, step=1)
            
            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
            
            # THE THIRD BUTTON - RUN WORKFLOW
            if st.button("🚀 " + t["run_btn"], use_container_width=True, type="primary"):
                with st.spinner(t["running"]):
                    try:
                        output, log = run_excel_workflow(
                            st.session_state.excel_bytes,
                            st.session_state.extracted_rev,
                            p_type,
                            hours
                        )
                        st.session_state.fixed_excel = output
                        st.session_state.workflow_log = log
                        for line in log:
                            st.markdown(f'<div class="log-line">{line}</div>', unsafe_allow_html=True)
                        st.success(t["run_ok"])
                    except Exception as e:
                        st.error(f"Error: {e}")
            
            if st.session_state.fixed_excel:
                st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
                pclean = parking.replace(" ", "_").replace("(", "").replace(")", "")
                st.download_button(
                    label="📥 " + t["dl_btn"],
                    data=st.session_state.fixed_excel,
                    file_name=f"{pclean}_budget_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="scard"><div class="scard-title">{t["config_title"]}</div><p style="color:#4A7090;">👈 Upload Excel and PDF files to begin</p></div>', unsafe_allow_html=True)

    # Footer
    st.markdown(f'<div class="db-footer">{t["footer"]}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ROUTER
# ─────────────────────────────────────────────
if st.session_state.authenticated:
    page_dashboard()
else:
    page_login()
