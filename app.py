import streamlit as st
import pdfplumber
from openpyxl import load_workbook
import io
import re
import os
import json
import requests
from datetime import datetime
import time
import base64
import tempfile

# ─────────────────────────────────────────────────────────────────
# DEEPGRAM VOICE SETUP - YOUR API KEY ADDED
# ─────────────────────────────────────────────────────────────────
# Your Deepgram API Key - $200 free credits included
DEEPGRAM_API_KEY = "3de1f753938a73b6e3f8d025c72ce235a3f41823"

try:
    from deepgram import DeepgramClient, SpeakOptions, PrerecordedOptions, FileSource
    import asyncio
    import simple_websocket
    DEEPGRAM_AVAILABLE = True
except ImportError:
    DEEPGRAM_AVAILABLE = False
    st.warning("Deepgram SDK not installed. Run: pip install deepgram-sdk")

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Budget System · Only Solutions",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────
# MISTRAL — raw HTTP (no SDK needed, errors surface correctly)
# ─────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = "em5oqjSdA1Nus9iUpa1MNAJtQA4YfCtK"
MISTRAL_URL = "https://api.mistral.ai/v1/chat/completions"
MISTRAL_MODEL = "mistral-small-latest"

def ask_mistral(history: list) -> str:
    system = {
        "role": "system",
        "content": (
            "You are a professional budget analyst for parking operations, you're also an expert in traffic data and parking data your expertise is concentrated in the province of Quebec with a special knowledge of the metropolitan area of Montreal, at Only Solutions Inc. "
            "Be precise and concise but not cold, you're warm and resourceful, you're a co-worker act like it. When revenue figures are available in the conversation, "
            "reference them directly in your analysis."
        ),
    }
    try:
        resp = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [system] + history,
            },
            timeout=30,
        )
        if resp.status_code == 200:
            return resp.json()["choices"][0]["message"]["content"]
        elif resp.status_code == 401:
            return "⚠️ **API key rejected (401).** Open [console.mistral.ai](https://console.mistral.ai) → API Keys → verify the key is active."
        elif resp.status_code == 403:
            return "⚠️ **IP not allowed (403).** Your Mistral key has an IP allowlist restriction. Go to **console.mistral.ai → API Keys → edit your key → remove the IP restriction** (or add the server IP). Then reload this page."
        elif resp.status_code == 429:
            return "⚠️ **Rate limit hit.** Wait a few seconds and try again."
        else:
            return f"⚠️ Mistral error {resp.status_code}: {resp.text[:300]}"
    except requests.exceptions.Timeout:
        return "⚠️ Request timed out. Check your internet connection."
    except Exception as e:
        return f"⚠️ Unexpected error: {e}"

# ─────────────────────────────────────────────────────────────────
# DEEPGRAM VOICE FUNCTIONS
# ─────────────────────────────────────────────────────────────────

def transcribe_audio(audio_bytes):
    """Convert speech to text using Deepgram's Nova-2 model"""
    if not DEEPGRAM_AVAILABLE:
        return "⚠️ Deepgram not installed. Please run: pip install deepgram-sdk"
    
    try:
        # Initialize Deepgram client
        deepgram = DeepgramClient(DEEPGRAM_API_KEY)
        
        # Configure for Nova-2 model (best accuracy)
        options = PrerecordedOptions(
            model="nova-2",  # Most accurate model
            smart_format=True,  # Auto-punctuation and casing
            language="en",  # English
        )
        
        # Transcribe the audio
        response = deepgram.listen.prerecorded.v("1").transcribe_file(
            {"buffer": audio_bytes}, options
        )
        
        # Extract transcript
        transcript = response["results"]["channels"][0]["alternatives"][0]["transcript"]
        return transcript if transcript else "I couldn't hear anything. Please try again."
        
    except Exception as e:
        return f"⚠️ Transcription error: {str(e)}"

def text_to_speech(text):
    """Convert text to speech using Deepgram's TTS"""
    if not DEEPGRAM_AVAILABLE:
        return None
    
    try:
        # Initialize Deepgram client
        deepgram = DeepgramClient(DEEPGRAM_API_KEY)
        
        # Configure TTS options - using Aura voice for natural speech
        options = SpeakOptions(
            model="aura-asteria-en",  # Natural female voice
            encoding="mp3",
        )
        
        # Generate speech
        response = deepgram.speak.v("1").stream({"text": text}, options)
        
        # Get the audio data
        audio_data = response.stream.getvalue()
        return audio_data
        
    except Exception as e:
        st.error(f"TTS error: {str(e)}")
        return None

# ─────────────────────────────────────────────────────────────────
# TRANSLATIONS
# ─────────────────────────────────────────────────────────────────
T_DATA = {
    "en": {
        "brand": "BUDGET SYSTEM",
        "brand_sub": "Only Solutions Inc.",
        "email_lbl": "Email address",
        "pass_lbl": "Password",
        "login_btn": "Sign in",
        "logout_btn": "Sign out",
        "wrong_creds": "Incorrect email or password.",
        "ai_title": "🤖 AI Assistant",
        "clear_chat": "Clear chat",
        "chat_hint": "Ask about budget, forecasts, or calculations…",
        "voice_hint": "🎤 Or click the microphone to speak",
        "recording": "🔴 Recording... Click stop when finished",
        "processing_voice": "Processing your voice...",
        "voice_button": "🎤 Speak",
        "play_response": "🔊 Play Response",
        "no_msgs": "No messages yet — start a conversation.",
        "files_title": "File Upload",
        "excel_lbl": "Excel Template (.xlsx)",
        "pdf_lbl": "PDF Report (.pdf)",
        "processing": "Processing files…",
        "files_ok": "Files ready — metrics extracted.",
        "transient": "Transient",
        "monthly": "Monthly",
        "total": "Total",
        "config_title": "Workflow",
        "parking_lbl": "Parking",
        "type_lbl": "Type",
        "hours_lbl": "Supervisor hrs/day",
        "run_btn": "Run Workflow",
        "running": "Running…",
        "run_ok": "Done — file ready to download.",
        "dl_btn": "Download Budget File",
        "upload_first": "Upload Excel + PDF to unlock workflow.",
        "admin_title": "Admin Panel",
        "new_user_title": "Create New User",
        "nm_lbl": "Full name",
        "new_email_lbl": "Email",
        "new_pass_lbl": "Password",
        "role_lbl": "Role",
        "create_btn": "Create User",
        "user_exists": "Email already in use.",
        "user_created": "User created successfully.",
        "users_title": "User List",
        "delete_btn": "Delete",
        "reset_title": "Reset password",
        "select_user": "Select user",
        "new_pw_lbl": "New password",
        "reset_btn": "Reset Password",
        "reset_ok": "Password updated successfully.",
        "theme_dark": "🌙 Dark Mode",
        "theme_light": "☀️ Light Mode",
        "footer": "Budget System · Only Solutions Inc.",
        "settings": "Settings",
        "profile": "Profile",
        "language": "Language",
        "appearance": "Appearance",
    },
    "fr": {
        "brand": "SYSTÈME BUDGÉTAIRE",
        "brand_sub": "Only Solutions Inc.",
        "email_lbl": "Adresse courriel",
        "pass_lbl": "Mot de passe",
        "login_btn": "Se connecter",
        "logout_btn": "Se déconnecter",
        "wrong_creds": "Courriel ou mot de passe incorrect.",
        "ai_title": "🤖 Assistant IA",
        "clear_chat": "Effacer",
        "chat_hint": "Budget, prévisions, calculs…",
        "voice_hint": "🎤 Ou cliquez sur le microphone pour parler",
        "recording": "🔴 Enregistrement... Cliquez sur stop quand vous avez fini",
        "processing_voice": "Traitement de votre voix...",
        "voice_button": "🎤 Parler",
        "play_response": "🔊 Écouter la réponse",
        "no_msgs": "Aucun message — commencez une conversation.",
        "files_title": "Fichiers",
        "excel_lbl": "Modèle Excel (.xlsx)",
        "pdf_lbl": "Rapport PDF (.pdf)",
        "processing": "Traitement…",
        "files_ok": "Fichiers prêts — métriques extraites.",
        "transient": "Transitoire",
        "monthly": "Mensuel",
        "total": "Total",
        "config_title": "Workflow",
        "parking_lbl": "Stationnement",
        "type_lbl": "Type",
        "hours_lbl": "Heures sup./jour",
        "run_btn": "Exécuter le workflow",
        "running": "Exécution…",
        "run_ok": "Terminé — fichier prêt.",
        "dl_btn": "Télécharger le fichier",
        "upload_first": "Téléversez Excel + PDF pour débloquer le workflow.",
        "admin_title": "Admin",
        "new_user_title": "Créer un utilisateur",
        "nm_lbl": "Nom complet",
        "new_email_lbl": "Courriel",
        "new_pass_lbl": "Mot de passe",
        "role_lbl": "Rôle",
        "create_btn": "Créer",
        "user_exists": "Ce courriel est déjà utilisé.",
        "user_created": "Utilisateur créé avec succès.",
        "users_title": "Liste des utilisateurs",
        "delete_btn": "Supprimer",
        "reset_title": "Réinitialiser",
        "select_user": "Sélectionner",
        "new_pw_lbl": "Nouveau mot de passe",
        "reset_btn": "Réinitialiser",
        "reset_ok": "Mot de passe mis à jour.",
        "theme_dark": "🌙 Mode Sombre",
        "theme_light": "☀️ Mode Clair",
        "footer": "Système budgétaire · Only Solutions Inc.",
        "settings": "Paramètres",
        "profile": "Profil",
        "language": "Langue",
        "appearance": "Apparence",
    },
}

# ─────────────────────────────────────────────────────────────────
# USER MANAGEMENT
# ─────────────────────────────────────────────────────────────────
USERS_FILE = "users.json"
ADMIN_EMAIL = "admin@onlys.com"

DEFAULT_USERS = {
    ADMIN_EMAIL: {
        "password": "12345",
        "name": "Administrator",
        "role": "admin",
        "created": datetime.now().isoformat(),
    }
}

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE) as f:
            return json.load(f)
    save_users(DEFAULT_USERS)
    return DEFAULT_USERS

def save_users(u):
    with open(USERS_FILE, "w") as f:
        json.dump(u, f, indent=2)

def authenticate(email, pw):
    u = load_users().get(email)
    return u if (u and u["password"] == pw) else None

def create_user(email, name, pw, role="user"):
    users = load_users()
    if email in users:
        return False
    users[email] = {"password": pw, "name": name, "role": role, "created": datetime.now().isoformat()}
    save_users(users)
    return True

def delete_user(email):
    if email == ADMIN_EMAIL:
        return False
    users = load_users()
    if email in users:
        del users[email]
        save_users(users)
        return True
    return False

def reset_password(email, pw):
    users = load_users()
    if email in users:
        users[email]["password"] = pw
        save_users(users)
        return True
    return False

# ─────────────────────────────────────────────────────────────────
# PDF / EXCEL
# ─────────────────────────────────────────────────────────────────
def extract_pdf_data(pdf_file):
    r = {"transient": 0, "monthly": 0, "total": 0}
    with pdfplumber.open(pdf_file) as pdf:
        text = "".join(p.extract_text() or "" for p in pdf.pages)
    nums = re.findall(r'\$?([\d,]+\.?\d*)', text)
    r["total"] = float(nums[0].replace(",", "")) if nums else 40000
    for line in text.split("\n"):
        ll, ns = line.lower(), re.findall(r'\$?([\d,]+\.?\d*)', line)
        if "transient" in ll and ns:
            r["transient"] = float(ns[0].replace(",", ""))
        elif "monthly" in ll and ns:
            r["monthly"] = float(ns[0].replace(",", ""))
    if not r["transient"]:
        r["transient"] = r["total"] * 0.6
    if not r["monthly"]:
        r["monthly"] = r["total"] * 0.4
    return r

def run_excel_workflow(excel_bytes, revenue, parking_type, supervisor_hours):
    wb = load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    log = []

    def sc(addr, val):
        try:
            ws[addr] = val
            ws[addr].number_format = "#,##0.00"
        except Exception:
            pass

    sc("K17", revenue["transient"])
    sc("K18", revenue["monthly"])
    sc("K26", revenue["total"])
    log.append(f"Revenue → K17 ${revenue['transient']:,.0f} K18 ${revenue['monthly']:,.0f} K26 ${revenue['total']:,.0f}")

    mult = ([1.2,1.2,1.2,1.1,0.8,0.8,0.8,0.8,1.1,1.1,1.2,1.2] if "School" in parking_type 
            else [0.8,0.8,0.9,0.9,1.3,1.3,1.3,1.2,1.0,1.0,0.8,0.8])
    base = revenue["total"] / 12
    for i, m in enumerate(mult):
        sc(f"K{20+i}", base * m)
    log.append("Monthly projections → K20–K31")

    sup = 25 * supervisor_hours * 30
    sc("K30", sup)
    log.append(f"Supervisor cost → K30 ${sup:,.2f}/month")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, log

# ─────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────
_D = dict(
    authenticated=False,
    user_email="",
    user_name="",
    user_role="",
    lang="en",
    theme="dark",
    messages=[],
    excel_bytes=None,
    extracted_rev={},
    files_ready=False,
    fixed_excel=None,
    workflow_log=[],
    show_settings=False,
    thinking=False,
    voice_input=None,
    audio_response=None,
)
for k, v in _D.items():
    if k not in st.session_state:
        st.session_state[k] = v

def T(key):
    return T_DATA[st.session_state.lang].get(key, key)

def do_logout():
    for k, v in _D.items():
        st.session_state[k] = v
    st.rerun()

# ─────────────────────────────────────────────────────────────────
# THEME TOKENS
# ─────────────────────────────────────────────────────────────────
DARK = dict(
    bg="#0D1117",
    surface="#161B22",
    border="#30363D",
    accent="#58A6FF",
    accent2="#D2A8FF",
    accent3="#3FB950",
    text="#E6EDF3",
    text_secondary="#8B949E",
    muted="#8B949E",
    danger="#F85149",
    bubble_user="#1F6FEB",
    bubble_bot="#21262D",
    input_bg="#0D1117",
    navbar="#161B22",
    btn_bg="#E67E22",
    btn_border="#F39C12",
    btn_text="#FFFFFF",
    run_bg="#E67E22",
    run_bg2="#F39C12",
    dl_bg="#E67E22",
    dl_color="#FFFFFF",
    dl_border="#F39C12",
    highlight="#F39C12",
)

LIGHT = dict(
    bg="#F5F7FA",
    surface="#FFFFFF",
    border="#D4DCE8",
    accent="#0066CC",
    accent2="#7C3AED",
    accent3="#0E7933",
    text="#1A2E45",
    text_secondary="#5C6F8C",
    muted="#5C6F8C",
    danger="#DC2626",
    bubble_user="#E6F0FF",
    bubble_bot="#F8FAFE",
    input_bg="#FFFFFF",
    navbar="#F0F4F9",
    btn_bg="#E67E22",
    btn_border="#F39C12",
    btn_text="#FFFFFF",
    run_bg="#E67E22",
    run_bg2="#F39C12",
    dl_bg="#E67E22",
    dl_color="#FFFFFF",
    dl_border="#F39C12",
    highlight="#E67E22",
)

def TK():
    return DARK if st.session_state.theme == "dark" else LIGHT

def inject_css():
    C = TK()
    is_dark = st.session_state.theme == "dark"
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"], .stApp, .main, .stApp > div, div[data-testid="stAppViewContainer"], div[data-testid="stHeader"] {{
        background-color: {C['bg']} !important;
        color: {C['text']} !important;
    }}
    
    /* Voice recording button styling */
    .voice-button {{
        background: {C['btn_bg']} !important;
        color: white !important;
        border-radius: 50% !important;
        width: 50px !important;
        height: 50px !important;
        padding: 0 !important;
        font-size: 24px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
    }}
    
    .recording-active {{
        animation: pulse 1s infinite;
        background: #ff4444 !important;
    }}
    
    @keyframes pulse {{
        0% {{ transform: scale(1); opacity: 1; }}
        50% {{ transform: scale(1.05); opacity: 0.8; }}
        100% {{ transform: scale(1); opacity: 1; }}
    }}
    
    /* SCROLLABLE CHAT CONTAINER */
    .chat-container {{
        height: 380px;
        overflow-y: auto;
        overflow-x: hidden;
        padding-right: 8px;
        margin-bottom: 12px;
        scrollbar-width: thin;
    }}
    
    .chat-container::-webkit-scrollbar {{
        width: 6px;
    }}
    
    .chat-container::-webkit-scrollbar-track {{
        background: {C['border']};
        border-radius: 3px;
    }}
    
    .chat-container::-webkit-scrollbar-thumb {{
        background: {C['highlight']};
        border-radius: 3px;
    }}
    
    /* THINKING DOTS */
    .thinking-container {{
        display: inline-block;
        padding: 0.6rem 0.9rem;
        background: {C['bubble_bot']};
        border: 1px solid {C['border']};
        border-left: 2px solid {C['accent2']};
        border-radius: 2px 8px 8px 8px;
        margin-bottom: 0.5rem;
    }}
    
    .dot {{
        display: inline-block;
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: {C['highlight']};
        margin: 0 3px;
        animation: dot-bounce 1.4s infinite ease-in-out both;
    }}
    
    .dot:nth-child(1) {{ animation-delay: -0.32s; }}
    .dot:nth-child(2) {{ animation-delay: -0.16s; }}
    .dot:nth-child(3) {{ animation-delay: 0s; }}
    
    @keyframes dot-bounce {{
        0%, 80%, 100% {{ transform: scale(0.6); opacity: 0.4; }}
        40% {{ transform: scale(1); opacity: 1; }}
    }}
    
    /* Chat bubbles */
    .bubble-user {{ background: {C['bubble_user']}; border: 1px solid {C['border']}; border-right: 2px solid {C['accent']}; padding: 0.5rem 0.75rem; border-radius: 8px 2px 8px 8px; margin-bottom: 0.5rem; margin-left: auto; margin-right: 0; max-width: 85%; width: fit-content; font-size: 0.82rem; line-height: 1.5; }}
    .bubble-bot {{ background: {C['bubble_bot']}; border: 1px solid {C['border']}; border-left: 2px solid {C['accent2']}; padding: 0.5rem 0.75rem; border-radius: 2px 8px 8px 8px; margin-bottom: 0.5rem; margin-right: auto; margin-left: 0; max-width: 85%; width: fit-content; font-size: 0.82rem; line-height: 1.5; }}
    .bubble-lbl {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.52rem; letter-spacing: 0.1em; text-transform: uppercase; color: {C['text_secondary']} !important; margin-bottom: 0.2rem; opacity: 0.7; }}
    .no-msgs {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.68rem; color: {C['text_secondary']} !important; text-align: center; padding: 2rem 0 1.5rem; letter-spacing: 0.06em; }}

    .scard {{ background: {C['surface']}; border: 1px solid {C['border']}; border-radius: 8px; padding: 1rem 1.1rem 1.1rem; margin-bottom: 1rem; height: 100%; }}
    .scard-title {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.6rem; font-weight: 500; color: {C['accent']} !important; text-transform: uppercase; letter-spacing: 0.16em; margin-bottom: 0.85rem; padding-bottom: 0.45rem; border-bottom: 1px solid {C['border']}; display: flex; align-items: center; gap: 0.4rem; }}
    .scard-title::before {{ content: ''; width: 2px; height: 9px; background: {C['accent']}; border-radius: 2px; flex-shrink: 0; }}

    .metric-row {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0.5rem; margin-top: 0.75rem; }}
    .mblock {{ background: {C['bg']}; border: 1px solid {C['border']}; border-radius: 6px; padding: 0.6rem 0.7rem; text-align: center; }}
    .mblock-label {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.55rem; color: {C['text_secondary']} !important; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 0.2rem; }}
    .mblock-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 1rem; font-weight: 600; line-height: 1; }}
    .mblock-val.c {{ color: {C['accent']} !important; }}
    .mblock-val.g {{ color: {C['accent2']} !important; }}
    .mblock-val.t {{ color: {C['accent3']} !important; }}

    .log-line {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.65rem; color: {C['accent3']} !important; padding: 0.15rem 0; letter-spacing: 0.03em; }}
    .log-line::before {{ content: '▸ '; opacity: 0.5; }}
    .hr {{ height: 1px; background: {C['border']}; margin: 0.7rem 0; }}

    .login-box {{ background: {C['surface']}; border: 1px solid {C['border']}; border-top: 2px solid {C['accent']}; border-radius: 10px; padding: 2rem 2rem 1.75rem; margin-top: 8vh; }}
    .login-brand {{ font-family: 'IBM Plex Mono', monospace; font-size: 1.05rem; font-weight: 600; color: {C['accent']} !important; letter-spacing: 0.07em; margin-bottom: 0.15rem; }}
    .login-sub {{ font-size: 0.6rem; color: {C['text_secondary']} !important; letter-spacing: 0.13em; text-transform: uppercase; margin-bottom: 1.6rem; }}
    .db-footer {{ text-align: center; font-family: 'IBM Plex Mono', monospace; font-size: 0.55rem; color: {C['text_secondary']} !important; letter-spacing: 0.1em; text-transform: uppercase; padding: 1.25rem 0 0.5rem; border-top: 1px solid {C['border']}; margin-top: 0.5rem; }}
    
    .stButton > button {{
        background: {C['btn_bg']} !important;
        color: {C['btn_text']} !important;
        border: 1px solid {C['btn_border']} !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.09em !important;
        text-transform: uppercase !important;
        border-radius: 6px !important;
        padding: 0.45rem 0.9rem !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
    }}
    
    .stButton > button:hover {{
        background: {C['run_bg2']} !important;
        border-color: #FFFFFF !important;
        transform: scale(1.01) !important;
    }}
    
    .stDownloadButton > button {{
        background: {C['dl_bg']} !important;
        color: #FFFFFF !important;
        border: 1px solid {C['dl_border']} !important;
        width: 100% !important;
    }}
    
    section[data-testid="stSidebar"] {{
        background: {C['navbar']} !important;
        border-right: 1px solid {C['border']} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# LOGIN PAGE
# ─────────────────────────────────────────────────────────────────
def page_login():
    inject_css()
    
    r1, r2, r3, r4 = st.columns([8, 0.75, 0.55, 0.55])
    with r2:
        theme_label = T("theme_light") if st.session_state.theme == "dark" else T("theme_dark")
        if st.button(theme_label, key="login_theme"):
            st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"
            st.rerun()
    with r3:
        if st.button("EN", key="login_en"):
            st.session_state.lang = "en"
            st.rerun()
    with r4:
        if st.button("FR", key="login_fr"):
            st.session_state.lang = "fr"
            st.rerun()

    _, center, _ = st.columns([1, 1.05, 1])
    with center:
        st.markdown(f"""
        <div class="login-box">
            <div class="login-brand">🤖 {T('brand')}</div>
            <div class="login-sub">{T('brand_sub')}</div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form", clear_on_submit=False):
            email = st.text_input(T("email_lbl"), placeholder="admin@onlys.com")
            password = st.text_input(T("pass_lbl"), type="password", placeholder="••••••••")
            submit = st.form_submit_button(T("login_btn"), use_container_width=True)
            
            if submit:
                user = authenticate(email, password)
                if user:
                    st.session_state.authenticated = True
                    st.session_state.user_email = email
                    st.session_state.user_name = user["name"]
                    st.session_state.user_role = user["role"]
                    st.rerun()
                else:
                    st.error(T("wrong_creds"))
        
        st.markdown(f'<div class="db-footer">{T("footer")}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# SETTINGS MENU
# ─────────────────────────────────────────────────────────────────
def render_settings_menu():
    with st.expander(f"⚙️ {T('settings')}", expanded=st.session_state.show_settings):
        st.markdown(f"**{T('profile')}**")
        st.info(f"**{st.session_state.user_name}**  \n`{st.session_state.user_email}`  \nRole: **{st.session_state.user_role.upper()}**")
        
        st.markdown("---")
        
        st.markdown(f"**{T('appearance')}**")
        col1, col2 = st.columns(2)
        with col1:
            if st.button(T("theme_dark"), use_container_width=True):
                st.session_state.theme = "dark"
                st.rerun()
        with col2:
            if st.button(T("theme_light"), use_container_width=True):
                st.session_state.theme = "light"
                st.rerun()
        
        st.markdown(f"**{T('language')}**")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("English", use_container_width=True):
                st.session_state.lang = "en"
                st.rerun()
        with col2:
            if st.button("Français", use_container_width=True):
                st.session_state.lang = "fr"
                st.rerun()
        
        st.markdown("---")
        
        # Voice Settings section
        st.markdown("**🎤 Voice Settings**")
        if DEEPGRAM_AVAILABLE:
            st.success("✅ Voice assistant is ready!")
            st.caption("Click the microphone button in the chat to speak. Your voice will be transcribed and sent to the AI.")
            st.caption(f"Using Deepgram API Key: {DEEPGRAM_API_KEY[:10]}...")
        else:
            st.warning("⚠️ Voice assistant not installed. Run: pip install deepgram-sdk")
            st.caption("Get your free $200 credit at [deepgram.com](https://deepgram.com)")
        
        st.markdown("---")
        
        if st.button(T("logout_btn"), use_container_width=True):
            do_logout()
        
        if st.session_state.user_role == "admin":
            st.markdown("---")
            st.markdown("### 👑 Admin Management")
            st.markdown(f"### 👤 {T('new_user_title')}")
            with st.form("create_user_form"):
                col1, col2 = st.columns(2)
                with col1:
                    new_name = st.text_input(T("nm_lbl"), placeholder="John Doe")
                with col2:
                    new_email = st.text_input(T("new_email_lbl"), placeholder="user@example.com")
                new_password = st.text_input(T("new_pass_lbl"), type="password", placeholder="••••••••")
                new_role = st.selectbox(T("role_lbl"), ["user", "admin"])
                
                if st.form_submit_button("➕ " + T("create_btn"), type="primary"):
                    if new_email and new_name and new_password:
                        if create_user(new_email, new_name, new_password, new_role):
                            st.success(f"✅ {T('user_created')}")
                            st.rerun()
                        else:
                            st.error(f"❌ {T('user_exists')}")
                    else:
                        st.warning("All fields required")
            
            with st.expander("📋 " + T("users_title")):
                users = load_users()
                for ue, ud in users.items():
                    ca, cb = st.columns([5, 1])
                    with ca:
                        tag = "👑 ADMIN" if ud["role"] == "admin" else "👤 USER"
                        st.markdown(f"**{ud['name']}** \n`{ue}`  \n*{tag}*", unsafe_allow_html=True)
                    with cb:
                        if ud["role"] != "admin":
                            if st.button("🗑️", key=f"del_{ue}"):
                                delete_user(ue)
                                st.rerun()
                    st.markdown("---")
            
            with st.expander("🔑 " + T("reset_title")):
                users = load_users()
                sel = st.selectbox(T("select_user"), list(users.keys()), key="reset_select")
                npw = st.text_input(T("new_pw_lbl"), type="password", key="reset_password_input")
                if st.button(T("reset_btn"), use_container_width=True):
                    if sel and npw:
                        reset_password(sel, npw)
                        st.success(T("reset_ok"))

# ─────────────────────────────────────────────────────────────────
# MAIN DASHBOARD
# ─────────────────────────────────────────────────────────────────
def page_dashboard():
    inject_css()
    
    # Navbar
    n1, n2, n3 = st.columns([6, 2, 0.65])
    with n1:
        st.markdown(f"""
        <div class="navbar">
            <span class="navbar-brand">🤖 {T('brand')}</span>
            <span class="navbar-user">{st.session_state.user_name} · {st.session_state.user_role.upper()}</span>
        </div>
        """, unsafe_allow_html=True)
    with n2:
        st.markdown(f"<div style='font-family:IBM Plex Mono,monospace;font-size:0.58rem;color:{TK()['text_secondary']};padding-top:0.6rem;text-align:right;'>{datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)
    with n3:
        if st.button("⚙️", key="settings_btn"):
            st.session_state.show_settings = not st.session_state.show_settings

    render_settings_menu()

    col_chat, col_files, col_wf = st.columns([1, 1.15, 0.95])

    with col_chat:
        st.markdown(f'<div class="scard"><div class="scard-title">{T("ai_title")}</div>', unsafe_allow_html=True)
        
        # Clear chat button
        _, btn_col = st.columns([3, 1])
        with btn_col:
            if st.button(T("clear_chat"), key="clr"):
                st.session_state.messages = []
                st.session_state.voice_input = None
                st.rerun()
        
        # SCROLLABLE CHAT CONTAINER
        st.markdown('<div class="chat-container">', unsafe_allow_html=True)
        
        if not st.session_state.messages:
            st.markdown(f'<div class="no-msgs">— {T("no_msgs")} —</div>', unsafe_allow_html=True)
        
        # Display all messages
        for idx, msg in enumerate(st.session_state.messages):
            if msg["role"] == "user":
                st.markdown(f'<div class="bubble-user"><div class="bubble-lbl">You</div>{msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="bubble-bot"><div class="bubble-lbl">Assistant</div>{msg["content"]}</div>', unsafe_allow_html=True)
                
                # Add play button for assistant responses if voice is enabled
                if DEEPGRAM_AVAILABLE:
                    if st.button(f"{T('play_response')}", key=f"play_{idx}"):
                        audio_data = text_to_speech(msg["content"])
                        if audio_data:
                            st.audio(audio_data, format="audio/mp3")
        
        # Show thinking animation while waiting for response
        if st.session_state.thinking:
            st.markdown("""
            <div class="thinking-container">
                <span class="dot"></span>
                <span class="dot"></span>
                <span class="dot"></span>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Voice input button (new feature)
        if DEEPGRAM_AVAILABLE:
            st.caption(T("voice_hint"))
            audio_value = st.audio_input("🎤 Click to record", key="voice_recorder")
            
            if audio_value and not st.session_state.thinking:
                with st.spinner(T("processing_voice")):
                    # Convert recorded audio to text
                    transcript = transcribe_audio(audio_value.getvalue())
                    if transcript and "⚠️" not in transcript and "couldn't hear" not in transcript.lower():
                        # Add the transcribed text as a user message
                        st.session_state.messages.append({"role": "user", "content": transcript})
                        st.session_state.thinking = True
                        st.rerun()
                    elif transcript:
                        st.error(transcript)
        
        # Text input (existing)
        user_input = st.chat_input(T("chat_hint"))
        if user_input and not st.session_state.thinking:
            ctx_suffix = ""
            if st.session_state.extracted_rev and not st.session_state.messages:
                rev = st.session_state.extracted_rev
                ctx_suffix = (f" [Budget context: Transient ${rev['transient']:,.0f}, "
                            f"Monthly ${rev['monthly']:,.0f}, Total ${rev['total']:,.0f}]")
            
            st.session_state.messages.append({"role": "user", "content": user_input})
            st.session_state.thinking = True
            st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

    with col_files:
        st.markdown(f'<div class="scard"><div class="scard-title">{T("files_title")}</div>', unsafe_allow_html=True)
        
        excel_file = st.file_uploader(T("excel_lbl"), type=["xlsx"], key="xl")
        pdf_file = st.file_uploader(T("pdf_lbl"), type=["pdf"], key="pd")
        
        if excel_file and pdf_file and not st.session_state.files_ready:
            with st.spinner(T("processing")):
                rev = extract_pdf_data(pdf_file)
                st.session_state.extracted_rev = rev
                st.session_state.excel_bytes = excel_file.read()
                st.session_state.files_ready = True
                st.session_state.fixed_excel = None
            st.success(T("files_ok"))
            st.rerun()
        
        if st.session_state.extracted_rev:
            rev = st.session_state.extracted_rev
            st.markdown(f"""
            <div class="metric-row">
                <div class="mblock">
                    <div class="mblock-label">{T('transient')}</div>
                    <div class="mblock-val c">${rev['transient']:,.0f}</div>
                </div>
                <div class="mblock">
                    <div class="mblock-label">{T('monthly')}</div>
                    <div class="mblock-val g">${rev['monthly']:,.0f}</div>
                </div>
                <div class="mblock">
                    <div class="mblock-label">{T('total')}</div>
                    <div class="mblock-val t">${rev['total']:,.0f}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    with col_wf:
        st.markdown(f'<div class="scard"><div class="scard-title">{T("config_title")}</div>', unsafe_allow_html=True)
        
        if not st.session_state.files_ready:
            st.markdown(f'<div style="font-size:0.78rem;color:{TK()["text_secondary"]};padding:1rem 0;">{T("upload_first")}</div>', unsafe_allow_html=True)
        else:
            parking = st.selectbox(T("parking_lbl"), ["CMO142 (LUNA)", "CMO143", "CMO144"])
            p_type = st.selectbox(T("type_lbl"), ["SC (School)", "RG (Tourism)"])
            hours = st.number_input(T("hours_lbl"), min_value=0, max_value=24, value=1, step=1)
            
            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
            
            if st.button("🚀 " + T("run_btn"), use_container_width=True, type="primary"):
                with st.spinner(T("running")):
                    try:
                        output, log = run_excel_workflow(
                            st.session_state.excel_bytes,
                            st.session_state.extracted_rev,
                            p_type,
                            hours,
                        )
                        st.session_state.fixed_excel = output
                        st.session_state.workflow_log = log
                        for line in log:
                            st.markdown(f'<div class="log-line">{line}</div>', unsafe_allow_html=True)
                        st.success(T("run_ok"))
                    except Exception as e:
                        st.error(f"Workflow error: {e}")
            
            if st.session_state.fixed_excel:
                st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
                pclean = parking.replace(" ", "_").replace("(", "").replace(")", "")
                st.download_button(
                    label="📥 " + T("dl_btn"),
                    data=st.session_state.fixed_excel,
                    file_name=f"{pclean}_budget_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown(f'<div class="db-footer">{T("footer")}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# Process AI response after rerun
# ─────────────────────────────────────────────────────────────────
if st.session_state.thinking:
    # Get the last user message
    user_messages = [m for m in st.session_state.messages if m["role"] == "user"]
    if user_messages:
        last_user_msg = user_messages[-1]["content"]
        
        ctx_suffix = ""
        if st.session_state.extracted_rev:
            rev = st.session_state.extracted_rev
            ctx_suffix = (f" [Budget context: Transient ${rev['transient']:,.0f}, "
                        f"Monthly ${rev['monthly']:,.0f}, Total ${rev['total']:,.0f}]")
        
        hist = st.session_state.messages[:-1] + [{"role": "user", "content": last_user_msg + ctx_suffix}]
        reply = ask_mistral(hist)
        st.session_state.messages.append({"role": "assistant", "content": reply})
        st.session_state.thinking = False
        st.rerun()

# ─────────────────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────────────────
if st.session_state.authenticated:
    page_dashboard()
else:
    page_login()
