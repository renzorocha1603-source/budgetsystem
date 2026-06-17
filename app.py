import streamlit as st
import pdfplumber
from openpyxl import load_workbook
import io
import re
import os
import json
import requests
from datetime import datetime
import csv
import zipfile
from xml.etree import ElementTree
from audio_recorder_streamlit import audio_recorder
from deepgram import DeepgramClient
import base64
from excel_fixer import fix_excel, get_parking_codes_from_pnl

# ============================================================================
# PAGE CONFIG
# ============================================================================
st.set_page_config(
    page_title="Budget System · Only Solutions",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================================
# DEEPGRAM CONFIGURATION
# ============================================================================
DEEPGRAM_API_KEY = "3de1f753938a73b6e3f8d025c72ce235a3f41823"

# ============================================================================
# MISTRAL CONFIGURATION
# ============================================================================
MISTRAL_API_KEY = "em5oqjSdA1Nus9iUpa1MNAJtQA4YfCtK"
MISTRAL_URL = "https://api.mistral.ai/v1/chat/completions"
MISTRAL_MODEL = "mistral-small-latest"

stop_requested = False


def ask_mistral(history: list) -> str:
    """Send chat history to Mistral and return the AI response."""
    global stop_requested
    stop_requested = False

    system = {
        "role": "system",
        "content": (
            "You are Allison, a senior budget analyst and operations specialist at Only Solutions Inc.\n\n"
            "Your expertise covers parking operations, budget forecasting, traffic data analysis, and "
            "inflation modeling with deep, specific knowledge of the province of Quebec and the "
            "greater Montreal metropolitan area (boroughs, traffic corridors, seasonal patterns, "
            "municipal context, and local operators).\n\n"
            "You also have broad knowledge of the private parking industry in Canada, including how "
            "major private operators are typically structured, how they approach budgeting, staffing, "
            "and facility management, and how the competitive landscape works in markets like Montreal. "
            "You speak about this in general, industry-level terms rather than referencing any specific "
            "company's internal data.\n\n"
            "Your personality:\n"
            "- You are warm, direct, and collegial, a trusted co-worker, not a formal consultant\n"
            "- You speak like a sharp colleague who genuinely wants to help, not like a report generator\n"
            "- You keep answers precise and actionable, but never cold or robotic\n"
            "- You use natural conversational language — short paragraphs, no unnecessary filler\n\n"
            "Your rules — non-negotiable:\n"
            "- If you don't know something, say so clearly and honestly: \"I don't have that data\" "
            "or \"I'm not sure about that one\" — never guess, never fill gaps with assumptions\n"
            "- If revenue or operational figures are present in the conversation, reference them "
            "directly and specifically in your analysis — never speak in generalities when "
            "real numbers are available\n"
            "- Never fabricate statistics, benchmarks, or regulatory details — Quebec parking "
            "regulations, SAAQ rules, municipal bylaws, or industry practices must only be cited "
            "if you are certain they are accurate\n"
            "- Never claim to have insider or proprietary knowledge of any specific company's "
            "internal operations — speak only in general industry terms\n"
            "- If asked something outside your domain, say so and redirect helpfully\n\n"
            "You are Allison. You know your stuff, you're here to make the work easier, "
            "and you treat every question like it deserves a real answer."
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


# ============================================================================
# VOICE: Speech-to-Text (transcription)
# ============================================================================
def transcribe_with_deepgram(audio_bytes, lang="en"):
    """Transcribe audio using Deepgram SDK v7"""
    if not DEEPGRAM_API_KEY:
        return None

    try:
        deepgram = DeepgramClient(DEEPGRAM_API_KEY)

        payload = {
            "buffer": audio_bytes,
        }

        if lang == "fr":
            speech_language = "fr"
        else:
            speech_language = "en"

        options = {
            "model": "nova-2",
            "smart_format": True,
            "language": speech_language,
        }

        response = deepgram.listen.prerecorded.v("1").transcribe_file(
            payload,
            options
        )

        transcript = response["results"]["channels"][0]["alternatives"][0]["transcript"]
        return transcript
    except Exception as e:
        return None


# ============================================================================
# VOICE: Text-to-Speech (Allison speaks back)
# ============================================================================
def clean_text_for_speech(text):
    """Remove markdown and symbols for clean speech"""
    clean = text
    clean = re.sub(r'\*\*(.*?)\*\*', r'\1', clean)
    clean = re.sub(r'\*(.*?)\*', r'\1', clean)
    clean = re.sub(r'`(.*?)`', r'\1', clean)
    clean = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', clean)
    clean = re.sub(r'#+\s*', '', clean)
    clean = re.sub(r'[-*]\s', '', clean)
    clean = re.sub(r'[\$\€\£\%\^\(\)\[\]\{\}]', '', clean)
    clean = re.sub(r'\s+', ' ', clean)
    return clean.strip()


def text_to_speech(text, lang="en"):
    """Convert Allison's text response to speech using Deepgram TTS (Aura-2)"""
    try:
        clean_text = clean_text_for_speech(text)

        deepgram = DeepgramClient(DEEPGRAM_API_KEY)

        if lang == "fr":
            voice_model = "aura-2-agathe-fr"
        else:
            voice_model = "aura-2-asteria-en"

        options = {
            "model": voice_model,
        }

        response = deepgram.speak.v("1").save(
            "allison_audio.mp3",
            {"text": clean_text},
            options
        )

        with open("allison_audio.mp3", "rb") as f:
            audio_bytes = f.read()

        audio_b64 = base64.b64encode(audio_bytes).decode()
        return audio_b64
    except Exception as e:
        return None


def play_audio_html(audio_b64):
    """Create HTML audio player that autoplays"""
    if audio_b64:
        audio_html = f"""
        <audio autoplay>
            <source src="data:audio/mp3;base64,{audio_b64}" type="audio/mp3">
        </audio>
        """
        st.markdown(audio_html, unsafe_allow_html=True)


# ============================================================================
# FILE EXTRACTION FUNCTIONS (for AI chat file analysis)
# ============================================================================
def extract_text_from_excel(file_bytes):
    """Extract text from Excel file for AI context"""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    excel_data = []
    for row in ws.iter_rows(min_row=1, max_row=min(50, ws.max_row), values_only=True):
        excel_data.append([str(cell) if cell is not None else "" for cell in row])
    return excel_data


def extract_text_from_csv(file_bytes):
    """Extract text from CSV file for AI context"""
    text = file_bytes.decode('utf-8', errors='ignore')
    reader = csv.reader(io.StringIO(text))
    csv_data = []
    for i, row in enumerate(reader):
        if i >= 50:
            break
        csv_data.append(row)
    return csv_data


def extract_text_from_pdf(file_bytes):
    """Extract text from PDF file using pdfplumber"""
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            text = "".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        pass
    return text if text.strip() else "Could not extract text from PDF"


def extract_text_from_docx(file_bytes):
    """Extract text from Word document"""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            if 'word/document.xml' in z.namelist():
                xml_content = z.read('word/document.xml')
                tree = ElementTree.fromstring(xml_content)
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                paragraphs = tree.findall('.//w:p', ns)
                text_parts = []
                for p in paragraphs:
                    texts = p.findall('.//w:t', ns)
                    para_text = ''.join(t.text for t in texts if t.text)
                    if para_text:
                        text_parts.append(para_text)
                return '\n'.join(text_parts)
    except Exception:
        pass
    return "Could not extract text from this document"


def extract_text_from_txt(file_bytes):
    """Extract text from text file"""
    return file_bytes.decode('utf-8', errors='ignore')


def process_any_file(uploaded_file):
    """Process any uploaded file and return extracted content for AI chat"""
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.read()

    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        return extract_text_from_excel(file_bytes), "excel", file_bytes
    elif file_name.endswith('.csv'):
        return extract_text_from_csv(file_bytes), "csv", file_bytes
    elif file_name.endswith('.pdf'):
        return extract_text_from_pdf(file_bytes), "pdf", file_bytes
    elif file_name.endswith('.docx'):
        return extract_text_from_docx(file_bytes), "docx", file_bytes
    elif file_name.endswith('.txt') or file_name.endswith('.md') or file_name.endswith('.py') or file_name.endswith('.json') or file_name.endswith('.xml') or file_name.endswith('.html') or file_name.endswith('.css') or file_name.endswith('.js'):
        return extract_text_from_txt(file_bytes), "text", file_bytes
    else:
        try:
            return file_bytes.decode('utf-8', errors='ignore'), "text", file_bytes
        except Exception:
            return f"Binary file: {file_name} (content cannot be displayed as text)", "binary", file_bytes


# ============================================================================
# TRANSLATIONS
# ============================================================================
T_DATA = {
    "en": {
        "brand": "BUDGET SYSTEM",
        "brand_sub": "Only Solutions Inc.",
        "email_lbl": "Email address",
        "pass_lbl": "Password",
        "login_btn": "Sign in",
        "logout_btn": "Sign out",
        "wrong_creds": "Incorrect email or password.",
        "ai_title": "🤖 Allison · AI Assistant",
        "clear_chat": "Clear chat",
        "chat_hint": "Ask Allison about budget, forecasts, or calculations…",
        "no_msgs": "No messages yet — start a conversation with Allison.",
        "files_title": "File Upload",
        "excel_lbl": "📊 Upload Excel Template (.xlsx)",
        "monthly_current_lbl": "📁 Current Year Monthly Reports (4 files: Jan-Apr 2026)",
        "monthly_previous_lbl": "📁 Previous Year Monthly Reports (8 files: May-Dec 2025)",
        "budget_initial_lbl": "📁 Budget Initial Source (Dec 2025 Page 3 / P&L)",
        "fiche_stationnement_lbl": "📁 Fiche Stationnement Source (Dec 2024 Page 3 / P&L)",
        "processing": "Processing files…",
        "files_ready": "✅ All files ready — you can now run the workflow.",
        "files_ready_partial": "✅ Template + Current Year data ready. Other files recommended.",
        "config_title": "Workflow",
        "run_btn": "🚀 Run Workflow",
        "running": "Processing your budget...",
        "run_ok": "✅ Done — file ready to download.",
        "dl_btn": "📥 Download Budget File",
        "upload_first": "⚠️ Upload Excel Template + Data files to unlock workflow.",
        "theme_dark": "🌙 Dark Mode",
        "theme_light": "☀️ Light Mode",
        "footer": "Budget System · Only Solutions Inc.",
        "settings": "Settings",
        "profile": "Profile",
        "language": "Language",
        "appearance": "Appearance",
        "send": "Send",
        "stop": "⏹ STOP",
        "ai_file_upload": "📎 Upload any file to analyze",
        "ai_file_loaded": "ready for questions",
        "clear_workflow": "🔄 Clear Workflow",
        "speak_now": "🎤 SPEAK NOW",
        "allison_online": "🟢 Allison is online and ready",
        "parking_code_lbl": "🏢 Select Parking Code",
        "parking_code_help": "Choose the parking location",
        "no_parking_codes": "No parking codes detected in uploaded files",
        "file_note_monthly": "Upload PDF or Excel files. Each file should be named with the month (e.g., January 2026.xlsx).",
        "file_note_page3": "Upload the December Page 3 Financial Summary (PDF or Excel). Extracts YTD Actual / Cumulatif courant values.",
        "file_note_accept": "Accepts: Excel, PDF, CSV",
    },
    "fr": {
        "brand": "SYSTÈME BUDGÉTAIRE",
        "brand_sub": "Only Solutions Inc.",
        "email_lbl": "Adresse courriel",
        "pass_lbl": "Mot de passe",
        "login_btn": "Se connecter",
        "logout_btn": "Se déconnecter",
        "wrong_creds": "Courriel ou mot de passe incorrect.",
        "ai_title": "🤖 Allison · Assistant IA",
        "clear_chat": "Effacer",
        "chat_hint": "Demandez à Allison budget, prévisions, calculs…",
        "no_msgs": "Aucun message — commencez une conversation avec Allison.",
        "files_title": "Fichiers",
        "excel_lbl": "📊 Téléverser le modèle Excel (.xlsx)",
        "monthly_current_lbl": "📁 Rapports mensuels année courante (4 fichiers: Jan-Avr 2026)",
        "monthly_previous_lbl": "📁 Rapports mensuels année précédente (8 fichiers: Mai-Déc 2025)",
        "budget_initial_lbl": "📁 Source Budget Initial (Déc 2025 Page 3 / P&L)",
        "fiche_stationnement_lbl": "📁 Source Fiche Stationnement (Déc 2024 Page 3 / P&L)",
        "processing": "Traitement…",
        "files_ready": "✅ Tous les fichiers prêts — exécutez le workflow.",
        "files_ready_partial": "✅ Modèle + données année courante prêts. Autres fichiers recommandés.",
        "config_title": "Workflow",
        "run_btn": "🚀 Exécuter",
        "running": "Traitement de votre budget...",
        "run_ok": "✅ Terminé — fichier prêt.",
        "dl_btn": "📥 Télécharger le fichier",
        "upload_first": "⚠️ Téléversez le modèle Excel + fichiers de données.",
        "theme_dark": "🌙 Mode Sombre",
        "theme_light": "☀️ Mode Clair",
        "footer": "Système budgétaire · Only Solutions Inc.",
        "settings": "Paramètres",
        "profile": "Profil",
        "language": "Langue",
        "appearance": "Apparence",
        "send": "Envoyer",
        "stop": "⏹ STOP",
        "ai_file_upload": "📎 Téléverser tout fichier à analyser",
        "ai_file_loaded": "prêt pour les questions",
        "clear_workflow": "🔄 Effacer Workflow",
        "speak_now": "🎤 PARLEZ",
        "allison_online": "🟢 Allison est en ligne et prête",
        "parking_code_lbl": "🏢 Sélectionner le code stationnement",
        "parking_code_help": "Choisissez le stationnement",
        "no_parking_codes": "Aucun code stationnement détecté dans les fichiers",
        "file_note_monthly": "Téléversez des fichiers PDF ou Excel. Chaque fichier doit être nommé avec le mois (ex: Janvier 2026.xlsx).",
        "file_note_page3": "Téléversez le Page 3 Résumé Financier de décembre (PDF ou Excel). Extrait les valeurs YTD Actual / Cumulatif courant.",
        "file_note_accept": "Accepte: Excel, PDF, CSV",
    },
}


# ============================================================================
# PERSISTENT CHAT MEMORY
# ============================================================================
CHAT_HISTORY_FILE = "chat_history.json"


def save_chat_history(messages):
    """Save chat messages to a file"""
    try:
        with open(CHAT_HISTORY_FILE, "w") as f:
            json.dump(messages, f, indent=2)
    except Exception:
        pass


def load_chat_history():
    """Load chat messages from file"""
    try:
        if os.path.exists(CHAT_HISTORY_FILE):
            with open(CHAT_HISTORY_FILE) as f:
                return json.load(f)
    except Exception:
        pass
    return []


def clear_chat_history():
    """Delete the chat history file"""
    try:
        if os.path.exists(CHAT_HISTORY_FILE):
            os.remove(CHAT_HISTORY_FILE)
    except Exception:
        pass


# ============================================================================
# USER MANAGEMENT
# ============================================================================
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
    """Load users from JSON file"""
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE) as f:
            return json.load(f)
    save_users(DEFAULT_USERS)
    return DEFAULT_USERS


def save_users(u):
    """Save users to JSON file"""
    with open(USERS_FILE, "w") as f:
        json.dump(u, f, indent=2)


def authenticate(email, pw):
    """Authenticate a user by email and password"""
    u = load_users().get(email)
    return u if (u and u["password"] == pw) else None


def create_user(email, name, pw, role="user"):
    """Create a new user"""
    users = load_users()
    if email in users:
        return False
    users[email] = {
        "password": pw,
        "name": name,
        "role": role,
        "created": datetime.now().isoformat()
    }
    save_users(users)
    return True


def delete_user(email):
    """Delete a user (cannot delete admin)"""
    if email == ADMIN_EMAIL:
        return False
    users = load_users()
    if email in users:
        del users[email]
        save_users(users)
        return True
    return False


def reset_password(email, pw):
    """Reset a user's password"""
    users = load_users()
    if email in users:
        users[email]["password"] = pw
        save_users(users)
        return True
    return False


# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================
_D = dict(
    authenticated=False,
    user_email="",
    user_name="",
    user_role="",
    lang="en",
    theme="dark",
    messages=[],
    excel_bytes=None,
    monthly_current_files=[],
    monthly_previous_files=[],
    budget_initial_file=None,
    fiche_stationnement_file=None,
    files_ready=False,
    fixed_excel=None,
    workflow_log=[],
    show_settings=False,
    thinking=False,
    thinking_from_voice=False,
    ai_file_data=None,
    ai_file_name="",
    ai_file_type="",
    voice_text="",
    last_audio=None,
    last_processed_text="",
    parking_codes=[],
    selected_parking_code=None,
)

for k, v in _D.items():
    if k not in st.session_state:
        st.session_state[k] = v


def T(key):
    """Get translated string for current language"""
    return T_DATA[st.session_state.lang].get(key, key)


def do_logout():
    """Log out the current user and clear session"""
    save_chat_history(st.session_state.messages)
    for k, v in _D.items():
        st.session_state[k] = v
    st.rerun()


# ============================================================================
# THEME TOKENS
# ============================================================================
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
    """Return current theme token dictionary"""
    return DARK if st.session_state.theme == "dark" else LIGHT


def inject_css():
    """Inject custom CSS based on current theme"""
    C = TK()
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"], .stApp, .main, .stApp > div, div[data-testid="stAppViewContainer"], div[data-testid="stHeader"] {{
        background-color: {C['bg']} !important;
        color: {C['text']} !important;
    }}
    
    p, span, div, label, .stMarkdown, .stText, .stCaption, .stException, .stCodeBlock, .stAlert, 
    .stSuccess, .stInfo, .stWarning, .stError, .stDataFrame, .stTable, .stJson,
    .element-container, .stTextInput label, .stSelectbox label, .stNumberInput label,
    .stFileUploader label, .stMultiSelect label, .stTextArea label {{
        color: {C['text']} !important;
    }}
    
    .chat-messages {{
        max-height: 400px;
        overflow-y: auto;
        overflow-x: hidden;
        padding-right: 8px;
        margin-bottom: 10px;
        scrollbar-width: thin;
    }}
    
    .chat-messages::-webkit-scrollbar {{ width: 6px; }}
    .chat-messages::-webkit-scrollbar-track {{ background: {C['border']}; border-radius: 3px; }}
    .chat-messages::-webkit-scrollbar-thumb {{ background: {C['highlight']}; border-radius: 3px; }}
    .chat-messages::-webkit-scrollbar-thumb:hover {{ background: {C['run_bg2']}; }}
    
    .thinking-container {{
        display: inline-flex !important;
        align-items: center;
        gap: 6px;
        padding: 0.5rem 0.8rem;
        background: {C['bubble_bot']};
        border: 1px solid {C['border']};
        border-left: 2px solid {C['accent2']};
        border-radius: 2px 8px 8px 8px;
        margin-bottom: 0.5rem;
    }}
    
    .robot-icon {{ font-size: 1.1rem; line-height: 1; }}
    
    .dot {{
        display: inline-block;
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: {C['highlight']};
        animation: dot-bounce 1.4s infinite ease-in-out both;
    }}
    
    .dot:nth-child(2) {{ animation-delay: -0.32s; }}
    .dot:nth-child(3) {{ animation-delay: -0.16s; }}
    .dot:nth-child(4) {{ animation-delay: 0s; }}
    
    @keyframes dot-bounce {{
        0%, 80%, 100% {{ transform: scale(0.6); opacity: 0.4; }}
        40% {{ transform: scale(1); opacity: 1; }}
    }}
    
    @keyframes pulse-green {{
        0%, 100% {{ box-shadow: 0 0 0 0 rgba(63, 185, 80, 0.7); transform: scale(1); }}
        50% {{ box-shadow: 0 0 0 6px rgba(63, 185, 80, 0); transform: scale(1.15); }}
    }}
    
    .alive-dot {{
        display: inline-block;
        width: 8px;
        height: 8px;
        background-color: #3FB950;
        border-radius: 50%;
        animation: pulse-green 2s infinite ease-in-out;
        vertical-align: middle;
    }}

    .stop-btn > button {{
        background: #DC2626 !important;
        color: #FFFFFF !important;
        border: 1px solid #DC2626 !important;
        animation: pulse-stop 1s infinite;
    }}
    
    @keyframes pulse-stop {{
        0%, 100% {{ box-shadow: 0 0 0 0 rgba(220, 38, 38, 0.5); }}
        50% {{ box-shadow: 0 0 0 6px rgba(220, 38, 38, 0); }}
    }}

    div[data-baseweb="select"] input[type="text"] {{
        color: #FFFFFF !important;
        background-color: #1A1A2E !important;
    }}
    
    div[data-baseweb="popover"] {{
        max-height: 300px !important;
        overflow-y: auto !important;
    }}
    
    ul[role="listbox"] {{
        max-height: 250px !important;
        overflow-y: auto !important;
    }}
    
    ul[role="listbox"] li {{
        color: #FFFFFF !important;
        background-color: #161B22 !important;
        padding: 8px 12px !important;
        cursor: pointer !important;
    }}
    
    ul[role="listbox"] li:hover {{
        background-color: #F39C12 !important;
        color: #000000 !important;
    }}
    
    div[data-baseweb="select"] input {{
        color: #FFFFFF !important;
        caret-color: #FFFFFF !important;
    }}
    
    div[data-testid="stSelectbox"] > div {{
        background-color: #161B22 !important;
        border: 1px solid #30363D !important;
        border-radius: 6px !important;
    }}
    
    div[data-testid="stSelectbox"] label {{
        color: #E6EDF3 !important;
        font-weight: 600 !important;
    }}

    div[data-baseweb="select"] div, 
    div[data-baseweb="select"] span,
    div[data-testid="stSelectbox"] div,
    .stSelectbox div,
    .stSelectbox span,
    div[role="listbox"] div,
    div[role="option"] {{
        color: {C['text']} !important;
        background-color: {C['surface']} !important;
    }}
    
    div[data-baseweb="popover"] div,
    div[data-baseweb="popover"] span,
    ul[role="listbox"] li div {{
        color: {C['text']} !important;
        background-color: {C['surface']} !important;
    }}
    
    div[data-baseweb="select"] [aria-selected="true"] {{
        background-color: {C['highlight']} !important;
        color: #000000 !important;
    }}
    
    div[data-testid="stFileUploader"] span, 
    div[data-testid="stFileUploader"] p, 
    div[data-testid="stFileUploader"] div,
    .stFileUploader div,
    .uploadedFileName,
    .stFileUploader span,
    div[data-testid="stFileUploader"] small,
    .stFileUploader small {{
        color: {C['highlight']} !important;
        opacity: 1 !important;
        font-weight: 500 !important;
    }}
    
    div[data-testid="stFileUploaderDropzone"] p {{
        color: {C['highlight']} !important;
        font-weight: 500 !important;
    }}
    
    .stFileUploader .e1y5xznm0 {{ color: {C['highlight']} !important; }}
    .stMarkdown small, .stCaption, .text-muted, .secondary-text {{ color: {C['highlight']} !important; }}
    
    .stButton > button {{
        background: #E67E22 !important;
        color: #FFFFFF !important;
        border: 1px solid #F39C12 !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.09em !important;
        text-transform: uppercase !important;
        border-radius: 6px !important;
        padding: 0.45rem 0.9rem !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
        cursor: pointer !important;
    }}
    
    .stButton > button:hover {{
        background: #F39C12 !important;
        border-color: #FFFFFF !important;
        transform: scale(1.01) !important;
        box-shadow: 0 2px 8px rgba(230,126,34,0.3) !important;
    }}
    
    .stButton > button[kind="primary"] {{
        background: #E67E22 !important;
        color: #FFFFFF !important;
        border-color: #F39C12 !important;
    }}
    
    .stButton > button[kind="primary"]:hover {{
        background: #F39C12 !important;
        transform: scale(1.01) !important;
    }}
    
    div[data-testid="stFormSubmitButton"] > button {{
        background: #E67E22 !important;
        color: #FFFFFF !important;
        border: 1px solid #F39C12 !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.09em !important;
        text-transform: uppercase !important;
        border-radius: 6px !important;
        padding: 0.45rem 0.9rem !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
        cursor: pointer !important;
    }}
    
    div[data-testid="stFormSubmitButton"] > button:hover {{
        background: #F39C12 !important;
        border-color: #FFFFFF !important;
        transform: scale(1.01) !important;
        box-shadow: 0 2px 8px rgba(230,126,34,0.3) !important;
    }}
    
    .stDownloadButton > button {{
        background: #E67E22 !important;
        color: #FFFFFF !important;
        border: 1px solid #F39C12 !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.09em !important;
        text-transform: uppercase !important;
        border-radius: 6px !important;
        padding: 0.45rem 0.9rem !important;
        width: 100% !important;
    }}
    
    .stDownloadButton > button:hover {{
        background: #F39C12 !important;
        border-color: #FFFFFF !important;
    }}
    
    button[kind="secondary"][data-testid="baseButton-secondary"] {{
        background: transparent !important;
        border: none !important;
        color: #E67E22 !important;
        font-size: 1.3rem !important;
        padding: 0.1rem 0.2rem !important;
        width: auto !important;
        min-width: auto !important;
        line-height: 1 !important;
    }}
    
    button[kind="secondary"][data-testid="baseButton-secondary"]:hover {{
        background: transparent !important;
        color: #F39C12 !important;
        transform: scale(1.15) !important;
        box-shadow: none !important;
    }}
    
    details {{
        background: {C['surface']} !important;
        border: 1px solid {C['border']} !important;
        border-radius: 8px !important;
        margin-bottom: 1rem !important;
    }}
    
    details summary {{
        color: {C['text']} !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.7rem !important;
        padding: 0.5rem !important;
    }}
    
    .main {{ padding: 0 !important; }}
    .block-container {{ padding: 1rem 1.5rem 2rem !important; max-width: 100% !important; }}
    #MainMenu, footer, header {{ visibility: hidden; }}
    
    ::-webkit-scrollbar {{ width: 4px; }}
    ::-webkit-scrollbar-track {{ background: {C['bg']}; }}
    ::-webkit-scrollbar-thumb {{ background: {C['border']}; border-radius: 4px; }}

    .navbar {{ display: flex; align-items: center; padding: 0.55rem 0.25rem; gap: 0.75rem; border-bottom: 1px solid {C['border']}; margin-bottom: 1.25rem; }}
    .navbar-brand {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem; font-weight: 600; color: {C['accent']} !important; letter-spacing: 0.07em; flex: 1; }}
    .navbar-user {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.6rem; color: {C['text_secondary']} !important; border-left: 1px solid {C['border']}; padding-left: 0.75rem; margin-left: 0.25rem; }}

    .scard {{ background: {C['surface']}; border: 1px solid {C['border']}; border-radius: 8px; padding: 1rem 1.1rem 1.1rem; margin-bottom: 1rem; }}
    .scard-title {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.6rem; font-weight: 500; color: {C['accent']} !important; text-transform: uppercase; letter-spacing: 0.16em; margin-bottom: 0.85rem; padding-bottom: 0.45rem; border-bottom: 1px solid {C['border']}; display: flex; align-items: center; gap: 0.4rem; }}
    .scard-title::before {{ content: ''; width: 2px; height: 9px; background: {C['accent']}; border-radius: 2px; flex-shrink: 0; }}

    .bubble-user {{ background: {C['bubble_user']}; border: 1px solid {C['border']}; border-right: 2px solid {C['accent']}; padding: 0.5rem 0.75rem; border-radius: 8px 2px 8px 8px; margin-bottom: 0.5rem; margin-left: auto; margin-right: 0; max-width: 85%; width: fit-content; font-size: 0.82rem; line-height: 1.5; color: {C['text']} !important; }}
    .bubble-bot {{ background: {C['bubble_bot']}; border: 1px solid {C['border']}; border-left: 2px solid {C['accent2']}; padding: 0.5rem 0.75rem; border-radius: 2px 8px 8px 8px; margin-bottom: 0.5rem; margin-right: auto; margin-left: 0; max-width: 85%; width: fit-content; font-size: 0.82rem; line-height: 1.5; color: {C['text']} !important; }}
    .bubble-lbl {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.52rem; letter-spacing: 0.1em; text-transform: uppercase; color: {C['text_secondary']} !important; margin-bottom: 0.2rem; opacity: 0.7; }}
    .no-msgs {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.68rem; color: {C['text_secondary']} !important; text-align: center; padding: 2rem 0 1.5rem; letter-spacing: 0.06em; }}

    .log-line {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.65rem; color: {C['accent3']} !important; padding: 0.15rem 0; letter-spacing: 0.03em; }}
    .log-line::before {{ content: '▸ '; opacity: 0.5; }}

    .hr {{ height: 1px; background: {C['border']}; margin: 0.7rem 0; }}

    .login-box {{ background: {C['surface']}; border: 1px solid {C['border']}; border-top: 2px solid {C['accent']}; border-radius: 10px; padding: 2rem 2rem 1.75rem; margin-top: 8vh; }}
    .login-brand {{ font-family: 'IBM Plex Mono', monospace; font-size: 1.05rem; font-weight: 600; color: {C['accent']} !important; letter-spacing: 0.07em; margin-bottom: 0.15rem; }}
    .login-sub {{ font-size: 0.6rem; color: {C['text_secondary']} !important; letter-spacing: 0.13em; text-transform: uppercase; margin-bottom: 1.6rem; }}

    .db-footer {{ text-align: center; font-family: 'IBM Plex Mono', monospace; font-size: 0.55rem; color: {C['text_secondary']} !important; letter-spacing: 0.1em; text-transform: uppercase; padding: 1.25rem 0 0.5rem; border-top: 1px solid {C['border']}; margin-top: 0.5rem; }}
    
    .streamlit-expanderContent {{ background: {C['surface']} !important; }}
    
    div[data-testid="stFileUploader"] button {{
        background: #E67E22 !important;
        color: #FFFFFF !important;
        border: 1px solid #F39C12 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }}
    
    div[data-testid="stFileUploader"] button:hover {{
        background: #F39C12 !important;
        border-color: #FFFFFF !important;
    }}
    
    .stCaption, caption, .help-text, .hint-text, .stFileUploader .e1y5xznm0 {{ color: {C['highlight']} !important; }}
    .metric-label {{ color: {C['highlight']} !important; }}
    .stSelectbox div[data-baseweb="select"] div {{ color: {C['text']} !important; }}
    
    .stInfo {{
        background-color: {C['surface']} !important;
        border: 1px solid {C['highlight']} !important;
        color: {C['text']} !important;
    }}
    
    .stInfo .stMarkdown {{ color: {C['text']} !important; }}
    
    .file-upload-note {{
        font-size: 0.55rem !important;
        color: {C['text_secondary']} !important;
        margin-top: -0.5rem !important;
        margin-bottom: 0.3rem !important;
        font-style: italic;
    }}
    
    .parking-selector {{
        background: {C['surface']} !important;
        border: 1px solid {C['highlight']} !important;
        border-radius: 6px !important;
        padding: 0.6rem 0.8rem !important;
        margin-bottom: 0.8rem !important;
    }}
    </style>
    """, unsafe_allow_html=True)


# ============================================================================
# LOGIN PAGE
# ============================================================================
def page_login():
    """Render the login page"""
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
                    st.session_state.messages = load_chat_history()
                    st.rerun()
                else:
                    st.error(T("wrong_creds"))

        st.markdown(f'<div class="db-footer">{T("footer")}</div>', unsafe_allow_html=True)


# ============================================================================
# SETTINGS MENU
# ============================================================================
def render_settings_menu():
    """Render the settings expander"""
    if st.session_state.show_settings:
        with st.expander(f"⚙️ {T('settings')}", expanded=True):
            st.markdown(f"**{T('profile')}**")
            st.info(f"**{st.session_state.user_name}**  \n`{st.session_state.user_email}`  \nRole: **{st.session_state.user_role.upper()}**")
            st.markdown("---")
            st.markdown(f"**{T('appearance')}**")
            col1, col2 = st.columns(2)
            with col1:
                if st.button(T("theme_dark"), use_container_width=True, key="theme_dark_settings"):
                    st.session_state.theme = "dark"
                    st.rerun()
            with col2:
                if st.button(T("theme_light"), use_container_width=True, key="theme_light_settings"):
                    st.session_state.theme = "light"
                    st.rerun()
            st.markdown(f"**{T('language')}**")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("English", use_container_width=True, key="en_settings"):
                    st.session_state.lang = "en"
                    st.rerun()
            with col2:
                if st.button("Français", use_container_width=True, key="fr_settings"):
                    st.session_state.lang = "fr"
                    st.rerun()
            st.markdown("---")
            if st.button(T("logout_btn"), use_container_width=True, key="logout_settings"):
                do_logout()
            if st.session_state.user_role == "admin":
                st.markdown("---")
                st.markdown("### 👑 Admin Management")
                st.markdown("### 👤 Create New User")
                with st.form("create_user_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        new_name = st.text_input("Full name", placeholder="John Doe")
                    with col2:
                        new_email = st.text_input("Email", placeholder="user@example.com")
                    new_password = st.text_input("Password", type="password", placeholder="••••••••")
                    new_role = st.selectbox("Role", ["user", "admin"])
                    if st.form_submit_button("➕ Create User", type="primary"):
                        if new_email and new_name and new_password:
                            if create_user(new_email, new_name, new_password, new_role):
                                st.success("✅ User created successfully.")
                                st.rerun()
                            else:
                                st.error("❌ Email already in use.")
                        else:
                            st.warning("All fields required")
                with st.expander("📋 User List"):
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
                with st.expander("🔑 Reset password"):
                    users = load_users()
                    sel = st.selectbox("Select user", list(users.keys()), key="reset_select")
                    npw = st.text_input("New password", type="password", key="reset_password_input")
                    if st.button("Reset Password", use_container_width=True):
                        if sel and npw:
                            reset_password(sel, npw)
                            st.success("Password updated successfully.")


# ============================================================================
# MAIN DASHBOARD
# ============================================================================
def page_dashboard():
    """Render the main dashboard page after login"""
    inject_css()

    # ── Navbar ──
    n1, n2 = st.columns([6, 0.65])
    with n1:
        st.markdown(f"""
        <div class="navbar">
            <span class="navbar-brand">🤖 {T('brand')}</span>
            <span class="navbar-user">{st.session_state.user_name} · {st.session_state.user_role.upper()}</span>
        </div>
        """, unsafe_allow_html=True)
    with n2:
        col_time, col_settings, col_alive = st.columns([2.5, 0.5, 0.5])
        with col_time:
            st.markdown(
                f"<div style='font-family:IBM Plex Mono,monospace;font-size:0.58rem;"
                f"color:{TK()['text_secondary']};padding-top:0.6rem;text-align:right;'>"
                f"{datetime.now().strftime('%Y-%m-%d %H:%M')}</div>",
                unsafe_allow_html=True
            )
        with col_settings:
            if st.button("⚙️", key="settings_btn"):
                st.session_state.show_settings = not st.session_state.show_settings
                st.rerun()
        with col_alive:
            st.markdown(
                '<div style="padding-top:0.55rem;text-align:center;">'
                '<span class="alive-dot"></span></div>',
                unsafe_allow_html=True
            )

    if st.session_state.show_settings:
        render_settings_menu()

    # ── AI CHAT SECTION ──
    st.markdown(f'<div class="scard"><div class="scard-title">{T("ai_title")}</div>', unsafe_allow_html=True)

    col_upload_area, col_clear_area = st.columns([1.5, 0.8])
    with col_upload_area:
        uploaded_file_for_ai = st.file_uploader(
            T("ai_file_upload"),
            type=None,
            key="ai_file_upload",
            label_visibility="collapsed"
        )
    with col_clear_area:
        if st.button(T("clear_chat"), key="clr", use_container_width=True):
            st.session_state.messages = []
            st.session_state.ai_file_data = None
            st.session_state.ai_file_name = ""
            st.session_state.ai_file_type = ""
            st.session_state.last_processed_text = ""
            clear_chat_history()
            st.rerun()

    if uploaded_file_for_ai and not st.session_state.ai_file_data:
        try:
            file_content, file_type, _ = process_any_file(uploaded_file_for_ai)
            st.session_state.ai_file_data = file_content
            st.session_state.ai_file_name = uploaded_file_for_ai.name
            st.session_state.ai_file_type = file_type
            st.success(f"✅ {uploaded_file_for_ai.name} loaded — ask me about it!")
            st.rerun()
        except Exception as e:
            st.error(f"Could not read file: {e}")

    if st.session_state.ai_file_data:
        file_name = st.session_state.ai_file_name
        file_type = st.session_state.ai_file_type
        st.markdown(
            f'<div style="font-size:0.6rem;color:{TK()["highlight"]};margin-bottom:0.5rem;">'
            f'📎 {file_name} ({file_type}) {T("ai_file_loaded")}</div>',
            unsafe_allow_html=True
        )

    # ── Chat Messages ──
    st.markdown('<div class="chat-messages">', unsafe_allow_html=True)
    if not st.session_state.messages:
        st.markdown(f'<div class="no-msgs">— {T("no_msgs")} —</div>', unsafe_allow_html=True)
    for msg in st.session_state.messages:
        if msg["role"] == "user":
            st.markdown(f'<div class="bubble-user"><div class="bubble-lbl">You</div>{msg["content"]}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="bubble-bot"><div class="bubble-lbl">Allison</div>{msg["content"]}</div>', unsafe_allow_html=True)
    if st.session_state.thinking:
        st.markdown(f"""<div class="thinking-container"><span class="robot-icon">🤖</span><span class="dot"></span><span class="dot"></span><span class="dot"></span></div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Chat Input ──
    with st.form(key="chat_form", clear_on_submit=True):
        col_input, col_send = st.columns([5, 1])
        with col_input:
            user_input = st.text_input(T("chat_hint"), placeholder=T("chat_hint"), label_visibility="collapsed", key="chat_input")
        with col_send:
            if st.session_state.thinking:
                st.markdown('<div class="stop-btn">', unsafe_allow_html=True)
                stop_clicked = st.form_submit_button("⏹ " + T("stop"), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
                submitted = False
            else:
                submitted = st.form_submit_button("➤ " + T("send"), use_container_width=True)
                stop_clicked = False
    st.markdown('</div>', unsafe_allow_html=True)

    if stop_clicked:
        st.session_state.thinking = False
        st.rerun()

    if submitted and user_input and user_input.strip() and user_input != st.session_state.last_processed_text:
        st.session_state.last_processed_text = user_input
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.session_state.thinking = True
        st.session_state.thinking_from_voice = False
        st.rerun()

    # ── Logo + Status + Microphone ──
    col_status_gap, col_logo_gap, col_mic_gap = st.columns([1.5, 1, 1.5])
    with col_status_gap:
        st.markdown(f"""<div style="text-align:right; padding-top: 12px; font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem; color: {TK()['highlight']};">{T("allison_online")}</div>""", unsafe_allow_html=True)
    with col_logo_gap:
        st.markdown(f"""<div style="text-align:center; padding: 0; margin: 0;"><img src="https://i.ibb.co/0yfv7KCS/image-1.jpg" width="130" style="border-radius: 8px; opacity: 0.9;"></div>""", unsafe_allow_html=True)
    with col_mic_gap:
        audio_bytes = audio_recorder(text=T("speak_now"), recording_color="#DC2626", neutral_color="#E67E22", icon_name="microphone", icon_size="1x", key="mic_recorder", pause_threshold=120.0, sample_rate=16000, energy_threshold=0.001)

    if audio_bytes:
        transcript = transcribe_with_deepgram(audio_bytes, st.session_state.lang)
        if transcript and transcript.strip() and transcript != st.session_state.last_processed_text:
            st.session_state.last_processed_text = transcript
            st.session_state.messages.append({"role": "user", "content": transcript})
            st.session_state.thinking = True
            st.session_state.thinking_from_voice = True
            st.rerun()

    # ── FILE UPLOAD + WORKFLOW SECTION ──
    col_files, col_wf = st.columns([1, 1])

    with col_files:
        st.markdown(f'<div class="scard"><div class="scard-title">{T("files_title")}</div>', unsafe_allow_html=True)

        # 1. Excel Template (REQUIRED)
        excel_file = st.file_uploader(T("excel_lbl"), type=["xlsx"], key="xl")

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # 2. Current Year Monthly Reports
        monthly_current = st.file_uploader(
            T("monthly_current_lbl"),
            type=["xlsx", "xls", "xlsm", "pdf", "csv"],
            accept_multiple_files=True,
            key="monthly_current"
        )
        st.markdown(f'<div class="file-upload-note">{T("file_note_monthly")}</div>', unsafe_allow_html=True)

        # 3. Previous Year Monthly Reports
        monthly_previous = st.file_uploader(
            T("monthly_previous_lbl"),
            type=["xlsx", "xls", "xlsm", "pdf", "csv"],
            accept_multiple_files=True,
            key="monthly_previous"
        )
        st.markdown(f'<div class="file-upload-note">{T("file_note_monthly")}</div>', unsafe_allow_html=True)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # 4. Budget Initial Source (Dec 2025 Page 3 or P&L)
        budget_initial_file = st.file_uploader(
            T("budget_initial_lbl"),
            type=["xlsx", "xls", "xlsm", "pdf", "csv", "tsv", "txt", "docx"],
            key="budget_initial"
        )
        st.markdown(f'<div class="file-upload-note">{T("file_note_page3")} {T("file_note_accept")}</div>', unsafe_allow_html=True)

        # 5. Fiche Stationnement Source (Dec 2024 Page 3 or P&L)
        fiche_stationnement_file = st.file_uploader(
            T("fiche_stationnement_lbl"),
            type=["xlsx", "xls", "xlsm", "pdf", "csv", "tsv", "txt", "docx"],
            key="fiche_stationnement"
        )
        st.markdown(f'<div class="file-upload-note">{T("file_note_page3")} {T("file_note_accept")}</div>', unsafe_allow_html=True)

        # Check if minimum required files are ready
        if excel_file:
            st.session_state.excel_bytes = excel_file
            st.session_state.budget_initial_file = budget_initial_file if budget_initial_file else None
            st.session_state.fiche_stationnement_file = fiche_stationnement_file if fiche_stationnement_file else None

            if monthly_current:
                st.session_state.monthly_current_files = list(monthly_current)
                st.session_state.monthly_previous_files = list(monthly_previous) if monthly_previous else []
                st.session_state.files_ready = True

                # Try to get parking codes from first monthly file
                try:
                    if len(monthly_current) > 0:
                        monthly_current[0].seek(0)
                        codes = get_parking_codes_from_pnl(monthly_current[0])
                        monthly_current[0].seek(0)
                        st.session_state.parking_codes = codes
                except Exception:
                    st.session_state.parking_codes = []

                if monthly_previous and budget_initial_file and fiche_stationnement_file:
                    st.success(T("files_ready"))
                else:
                    st.info(T("files_ready_partial"))
            else:
                st.warning("⚠️ Current Year monthly files are required.")
                st.session_state.files_ready = False

        st.markdown('</div>', unsafe_allow_html=True)

    with col_wf:
        st.markdown(f'<div class="scard"><div class="scard-title">{T("config_title")}</div>', unsafe_allow_html=True)

        if not st.session_state.files_ready:
            st.markdown(f'<div style="font-size:0.78rem;color:{TK()["text_secondary"]};padding:1rem 0;">{T("upload_first")}</div>', unsafe_allow_html=True)
        else:
            template_name = st.session_state.excel_bytes.name if hasattr(st.session_state.excel_bytes, 'name') else "Template"
            st.markdown(f"**Template:** {template_name}")
            st.markdown(f"**Current Year Files:** {len(st.session_state.monthly_current_files)} files")
            st.markdown(f"**Previous Year Files:** {len(st.session_state.monthly_previous_files)} files")

            bi_name = st.session_state.budget_initial_file.name if st.session_state.budget_initial_file and hasattr(st.session_state.budget_initial_file, 'name') else "Not provided"
            fs_name = st.session_state.fiche_stationnement_file.name if st.session_state.fiche_stationnement_file and hasattr(st.session_state.fiche_stationnement_file, 'name') else "Not provided"
            st.markdown(f"**Budget Initial Source:** {bi_name}")
            st.markdown(f"**Fiche Stationnement Source:** {fs_name}")

            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

            # ── PARKING CODE SELECTOR ──
            parking_codes = st.session_state.get("parking_codes", [])
            if parking_codes:
                selected_code = st.selectbox(T("parking_code_lbl"), options=parking_codes, help=T("parking_code_help"), key="parking_code_select")
                st.session_state.selected_parking_code = selected_code
            else:
                fallback_code = None
                if hasattr(st.session_state.excel_bytes, 'name'):
                    match = re.search(r'(CMO\d+)', st.session_state.excel_bytes.name, re.IGNORECASE)
                    if match:
                        fallback_code = match.group(1).upper()
                if fallback_code:
                    st.info(f"📌 Parking code from filename: **{fallback_code}**")
                    st.session_state.selected_parking_code = fallback_code
                else:
                    st.warning(T("no_parking_codes"))
                    st.session_state.selected_parking_code = None

            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

            col_run, col_clear_wf = st.columns([1, 1])
            with col_run:
                if st.button(T("run_btn"), use_container_width=True, type="primary"):
                    if not st.session_state.selected_parking_code:
                        st.error("❌ Please select a parking code before running.")
                    else:
                        with st.spinner(T("running")):
                            try:
                                fixed_excel, updates = fix_excel(
                                    excel_file=st.session_state.excel_bytes,
                                    monthly_files_current=st.session_state.monthly_current_files,
                                    monthly_files_previous=st.session_state.monthly_previous_files,
                                    budget_initial_file=st.session_state.budget_initial_file,
                                    fiche_stationnement_file=st.session_state.fiche_stationnement_file,
                                    parking_code=st.session_state.selected_parking_code,
                                    word_data=None
                                )

                                st.session_state.fixed_excel = fixed_excel
                                st.session_state.workflow_log = updates

                                if updates:
                                    for update in updates:
                                        st.markdown(f'<div class="log-line">{update}</div>', unsafe_allow_html=True)
                                    success_count = sum(1 for u in updates if u.startswith("✅"))
                                    if fixed_excel:
                                        if success_count > 0:
                                            st.success(T("run_ok"))
                                        else:
                                            st.warning("Workflow completed with warnings.")
                                    else:
                                        st.error("Workflow failed - no output file generated.")
                                else:
                                    st.warning("No updates were made.")
                            except Exception as e:
                                st.error(f"Workflow error: {str(e)}")

            with col_clear_wf:
                if st.button(T("clear_workflow"), use_container_width=True):
                    st.session_state.excel_bytes = None
                    st.session_state.monthly_current_files = []
                    st.session_state.monthly_previous_files = []
                    st.session_state.budget_initial_file = None
                    st.session_state.fiche_stationnement_file = None
                    st.session_state.files_ready = False
                    st.session_state.fixed_excel = None
                    st.session_state.workflow_log = []
                    st.session_state.parking_codes = []
                    st.session_state.selected_parking_code = None
                    st.rerun()

            if st.session_state.fixed_excel:
                st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
                template_name_clean = template_name.rsplit('.', 1)[0].replace(" ", "_")
                st.download_button(
                    label=T("dl_btn"),
                    data=st.session_state.fixed_excel,
                    file_name=f"{template_name_clean}_updated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f'<div class="db-footer">{T("footer")}</div>', unsafe_allow_html=True)


# ============================================================================
# PROCESS AI RESPONSE WITH MEMORY
# ============================================================================
if st.session_state.thinking:
    user_messages = [m for m in st.session_state.messages if m["role"] == "user"]
    if user_messages:
        last_user_msg = user_messages[-1]["content"]
        ctx_suffix = ""
        if st.session_state.ai_file_data:
            file_data = st.session_state.ai_file_data
            file_name = st.session_state.ai_file_name
            file_type = st.session_state.ai_file_type
            ctx_suffix += f"\n\n[Uploaded {file_type} file: {file_name}]\n"
            if file_type in ["excel", "csv"]:
                ctx_suffix += "Spreadsheet:\n"
                for row in file_data[:30]:
                    ctx_suffix += " | ".join(row) + "\n"
            else:
                ctx_suffix += f"Content:\n{str(file_data)[:3000]}\n"
            ctx_suffix += f"\n[End of {file_name}]\n"
        if st.session_state.extracted_rev:
            rev = st.session_state.extracted_rev
            ctx_suffix += f" [Budget: Transient ${rev['transient']:,.0f}, Monthly ${rev['monthly']:,.0f}, Total ${rev['total']:,.0f}]"
        recent_history = st.session_state.messages[-12:] if len(st.session_state.messages) > 12 else st.session_state.messages
        history_for_mistral = []
        for msg in recent_history[:-1]:
            history_for_mistral.append({"role": msg["role"], "content": msg["content"]})
        history_for_mistral.append({"role": "user", "content": last_user_msg + ctx_suffix})
        reply = ask_mistral(history_for_mistral)
        st.session_state.messages.append({"role": "assistant", "content": reply})
        st.session_state.thinking = False
        save_chat_history(st.session_state.messages)
        if st.session_state.thinking_from_voice:
            audio_b64 = text_to_speech(reply, st.session_state.lang)
            st.session_state.last_audio = audio_b64
        else:
            st.session_state.last_audio = None
        st.rerun()

if st.session_state.get("last_audio"):
    play_audio_html(st.session_state.last_audio)
    st.session_state.last_audio = None


# ============================================================================
# ROUTER
# ============================================================================
if st.session_state.authenticated:
    page_dashboard()
else:
    page_login()
