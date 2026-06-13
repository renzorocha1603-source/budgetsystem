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
import csv
import zipfile
from xml.etree import ElementTree
from audio_recorder_streamlit import audio_recorder
from deepgram import DeepgramClient
import base64

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
# DEEPGRAM CONFIGURATION
# ─────────────────────────────────────────────────────────────────
DEEPGRAM_API_KEY = "3de1f753938a73b6e3f8d025c72ce235a3f41823"

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
            "You are Allison, a senior budget analyst and operations specialist at Only Solutions Inc.\n\n"
            "Your expertise covers parking operations, budget forecasting, traffic data analysis, and "
            "inflation modeling — with deep, specific knowledge of the province of Quebec and the "
            "greater Montreal metropolitan area (boroughs, traffic corridors, seasonal patterns, "
            "municipal context, and local operators).\n\n"
            "Your personality:\n"
            "- You are warm, direct, and collegial — a trusted co-worker, not a formal consultant\n"
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
            "regulations, SAAQ rules, municipal bylaws, and ARTM data must only be cited "
            "if you are certain they are accurate\n"
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

# ─────────────────────────────────────────────────────────────────
# VOICE: Speech-to-Text (transcription)
# ─────────────────────────────────────────────────────────────────
def transcribe_with_deepgram(audio_bytes):
    """Transcribe audio using Deepgram SDK v7"""
    if not DEEPGRAM_API_KEY:
        return None
    
    try:
        deepgram = DeepgramClient(DEEPGRAM_API_KEY)
        
        payload = {
            "buffer": audio_bytes,
        }
        
        options = {
            "model": "nova-2",
            "smart_format": True,
            "language": "en",
        }
        
        response = deepgram.listen.prerecorded.v("1").transcribe_file(
            payload,
            options
        )
        
        transcript = response["results"]["channels"][0]["alternatives"][0]["transcript"]
        return transcript
    except Exception as e:
        return None

# ─────────────────────────────────────────────────────────────────
# VOICE: Text-to-Speech (Allison speaks back)
# ─────────────────────────────────────────────────────────────────
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

def text_to_speech(text):
    """Convert Allison's text response to speech using Deepgram TTS"""
    try:
        clean_text = clean_text_for_speech(text)
        
        deepgram = DeepgramClient(DEEPGRAM_API_KEY)
        
        options = {
            "model": "aura-asteria-en",
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

# ─────────────────────────────────────────────────────────────────
# FILE EXTRACTION FUNCTIONS (using only installed libraries)
# ─────────────────────────────────────────────────────────────────
def extract_text_from_excel(file_bytes):
    """Extract text from Excel file"""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    excel_data = []
    for row in ws.iter_rows(min_row=1, max_row=min(50, ws.max_row), values_only=True):
        excel_data.append([str(cell) if cell is not None else "" for cell in row])
    return excel_data

def extract_text_from_csv(file_bytes):
    """Extract text from CSV file"""
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
    except:
        pass
    return text if text.strip() else "Could not extract text from PDF"

def extract_text_from_docx(file_bytes):
    """Extract text from Word document using built-in zipfile and ElementTree"""
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
    except:
        pass
    return "Could not extract text from this document"

def extract_text_from_txt(file_bytes):
    """Extract text from text file"""
    return file_bytes.decode('utf-8', errors='ignore')

def process_any_file(uploaded_file):
    """Process any uploaded file and return extracted content"""
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
        except:
            return f"Binary file: {file_name} (content cannot be displayed as text)", "binary", file_bytes

def extract_pdf_revenue(pdf_file_bytes):
    """Extract revenue data from PDF for budget workflow"""
    r = {"transient": 0, "monthly": 0, "total": 0}
    try:
        with pdfplumber.open(io.BytesIO(pdf_file_bytes)) as pdf:
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
    except:
        r["total"] = 40000
        r["transient"] = 24000
        r["monthly"] = 16000
    return r

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
        "ai_title": "🤖 Allison · AI Assistant",
        "clear_chat": "Clear chat",
        "chat_hint": "Ask Allison about budget, forecasts, or calculations…",
        "no_msgs": "No messages yet — start a conversation with Allison.",
        "files_title": "File Upload",
        "excel_lbl": "Excel Template (any format)",
        "pdf_lbl": "PDF Report (any format)",
        "processing": "Processing files…",
        "files_ok": "✅ Files ready — you can now run the workflow.",
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
        "upload_first": "Upload files to unlock workflow.",
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
        "send": "Send",
        "ai_file_upload": "📎 Upload any file to analyze",
        "ai_file_loaded": "ready for questions",
        "clear_workflow": "Clear Workflow",
        "speak_now": "🎤 SPEAK NOW",
        "thinking_msg": "🤖 Allison is thinking...",
        "allison_online": "🟢 Allison is online and ready",
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
        "excel_lbl": "Modèle Excel (tout format)",
        "pdf_lbl": "Rapport PDF (tout format)",
        "processing": "Traitement…",
        "files_ok": "✅ Fichiers prêts — vous pouvez exécuter le workflow.",
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
        "upload_first": "Téléversez fichiers pour débloquer le workflow.",
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
        "send": "Envoyer",
        "ai_file_upload": "📎 Téléverser tout fichier à analyser",
        "ai_file_loaded": "prêt pour les questions",
        "clear_workflow": "Effacer Workflow",
        "speak_now": "🎤 PARLEZ",
        "thinking_msg": "🤖 Allison réfléchit...",
        "allison_online": "🟢 Allison est en ligne et prête",
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
    ai_file_data=None,
    ai_file_name="",
    ai_file_type="",
    voice_text="",
    last_audio=None,
    last_processed_text="",
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
    
    p, span, div, label, .stMarkdown, .stText, .stCaption, .stException, .stCodeBlock, .stAlert, 
    .stSuccess, .stInfo, .stWarning, .stError, .stDataFrame, .stTable, .stJson,
    .element-container, .stTextInput label, .stSelectbox label, .stNumberInput label,
    .stFileUploader label, .stMultiSelect label, .stTextArea label {{
        color: {C['text']} !important;
    }}
    
    /* SCROLLABLE CHAT CONTAINER */
    .chat-messages {{
        height: 400px;
        max-height: 400px;
        overflow-y: auto !important;
        overflow-x: hidden;
        padding-right: 8px;
        margin-bottom: 10px;
        scrollbar-width: thin;
    }}
    
    .chat-messages::-webkit-scrollbar {{
        width: 6px;
    }}
    
    .chat-messages::-webkit-scrollbar-track {{
        background: {C['border']};
        border-radius: 3px;
    }}
    
    .chat-messages::-webkit-scrollbar-thumb {{
        background: {C['highlight']};
        border-radius: 3px;
    }}
    
    .chat-messages::-webkit-scrollbar-thumb:hover {{
        background: {C['run_bg2']};
    }}
    
    /* THINKING DOTS WITH ROBOT */
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
    
    .robot-icon {{
        font-size: 1.1rem;
        line-height: 1;
    }}
    
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
    
    /* GREEN PULSING DOT - ALIVE INDICATOR */
    @keyframes pulse-green {{
        0%, 100% {{ 
            box-shadow: 0 0 0 0 rgba(63, 185, 80, 0.7);
            transform: scale(1);
        }}
        50% {{ 
            box-shadow: 0 0 0 6px rgba(63, 185, 80, 0);
            transform: scale(1.15);
        }}
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

    /* Make ALL selectbox options visible */
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
    ul[role="listbox"] li,
    ul[role="listbox"] li div {{
        color: {C['text']} !important;
        background-color: {C['surface']} !important;
    }}
    
    ul[role="listbox"] li:hover {{
        background-color: {C['highlight']} !important;
        color: #000000 !important;
    }}
    
    div[data-baseweb="select"] [aria-selected="true"] {{
        background-color: {C['highlight']} !important;
        color: #000000 !important;
    }}
    
    /* File upload hint text */
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
    
    .stFileUploader .e1y5xznm0 {{
        color: {C['highlight']} !important;
    }}
    
    .stMarkdown small, .stCaption, .text-muted, .secondary-text {{
        color: {C['highlight']} !important;
    }}
    
    /* ALL REGULAR BUTTONS - ORANGE */
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
    
    /* FORM SUBMIT BUTTONS - ALSO ORANGE */
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
    
    /* DOWNLOAD BUTTON */
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
    
    /* SETTINGS BUTTON - JUST ORANGE ICON, NO BOX */
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
    
    /* Settings expander */
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

    .metric-row {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0.5rem; margin-top: 0.75rem; }}
    .mblock {{ background: {C['bg']}; border: 1px solid {C['border']}; border-radius: 6px; padding: 0.6rem 0.7rem; }}
    .mblock-label {{ font-family: 'IBM Plex Mono', monospace; font-size: 0.55rem; color: {C['text_secondary']} !important; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 0.2rem; }}
    .mblock-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 1rem; font-weight: 600; line-height: 1; }}
    .mblock-val.c {{ color: {C['accent']} !important; }}
    .mblock-val.g {{ color: {C['accent2']} !important; }}
    .mblock-val.t {{ color: {C['accent3']} !important; }}

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
    
    .streamlit-expanderContent {{
        background: {C['surface']} !important;
    }}
    
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
    
    .stCaption, caption, .help-text, .hint-text, .stFileUploader .e1y5xznm0 {{
        color: {C['highlight']} !important;
    }}
    
    .metric-label {{
        color: {C['highlight']} !important;
    }}
    
    .stSelectbox div[data-baseweb="select"] div {{
        color: {C['text']} !important;
    }}
    
    .stInfo {{
        background-color: {C['surface']} !important;
        border: 1px solid {C['highlight']} !important;
        color: {C['text']} !important;
    }}
    
    .stInfo .stMarkdown {{
        color: {C['text']} !important;
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
# SETTINGS MENU (only shows when ⚙️ clicked)
# ─────────────────────────────────────────────────────────────────
def render_settings_menu():
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
    
    # Navbar - only ONE settings icon
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
            st.markdown(f"<div style='font-family:IBM Plex Mono,monospace;font-size:0.58rem;color:{TK()['text_secondary']};padding-top:0.6rem;text-align:right;'>{datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)
        with col_settings:
            if st.button("⚙️", key="settings_btn"):
                st.session_state.show_settings = not st.session_state.show_settings
                st.rerun()
        with col_alive:
            st.markdown('<div style="padding-top:0.55rem;text-align:center;"><span class="alive-dot"></span></div>', unsafe_allow_html=True)

    # Settings expander (only shows when toggled)
    if st.session_state.show_settings:
        render_settings_menu()

    # ============ AI CHAT - FULL WIDTH TOP ============
    st.markdown(f'<div class="scard"><div class="scard-title">{T("ai_title")}</div>', unsafe_allow_html=True)
    
    # Top row: File upload + Clear chat
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
            st.rerun()
    
    # Process uploaded file for AI context
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
    
    # Show indicator if file is loaded
    if st.session_state.ai_file_data:
        file_name = st.session_state.ai_file_name
        file_type = st.session_state.ai_file_type
        st.markdown(f'<div style="font-size:0.6rem;color:{TK()["highlight"]};margin-bottom:0.5rem;">📎 {file_name} ({file_type}) {T("ai_file_loaded")}</div>', unsafe_allow_html=True)
    
    # SCROLLABLE MESSAGES AREA
    st.markdown('<div class="chat-messages">', unsafe_allow_html=True)
    
    if not st.session_state.messages:
        st.markdown(f'<div class="no-msgs">— {T("no_msgs")} —</div>', unsafe_allow_html=True)
    
    for msg in st.session_state.messages:
        if msg["role"] == "user":
            st.markdown(f'<div class="bubble-user"><div class="bubble-lbl">You</div>{msg["content"]}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="bubble-bot"><div class="bubble-lbl">Allison</div>{msg["content"]}</div>', unsafe_allow_html=True)
    
    # THINKING INDICATOR - Native Streamlit component
    if st.session_state.thinking:
        st.info(T("thinking_msg"))
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ============ GAP FILLER: Logo + Status + Mic in one row ============
    col_logo, col_status, col_mic_gap = st.columns([1, 2, 1])
    
    with col_logo:
        st.markdown(f"""
        <div style="text-align:center; padding: 0; margin: 0;">
            <img src="https://i.ibb.co/0yfv7KCS/image-1.jpg" width="120" style="border-radius: 8px; opacity: 0.9;">
        </div>
        """, unsafe_allow_html=True)
    
    with col_status:
        st.markdown(f"""
        <div style="text-align:center; padding-top: 15px; font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: {TK()['highlight']};">
            {T("allison_online")}
        </div>
        """, unsafe_allow_html=True)
    
    with col_mic_gap:
        audio_bytes = audio_recorder(
            text=T("speak_now"),
            recording_color="#DC2626",
            neutral_color="#E67E22",
            icon_name="microphone",
            icon_size="1x",
            key="mic_recorder"
        )
    
    # Handle voice input - with duplicate prevention
    if audio_bytes:
        transcript = transcribe_with_deepgram(audio_bytes)
        if transcript and transcript.strip() and transcript != st.session_state.last_processed_text:
            st.session_state.last_processed_text = transcript
            st.session_state.messages.append({"role": "user", "content": transcript})
            st.session_state.thinking = True
            st.rerun()
    
    # ============ CHAT INPUT - text_input + button ============
    col_input, col_send = st.columns([5, 1])
    with col_input:
        user_input = st.text_input(
            T("chat_hint"),
            placeholder=T("chat_hint"),
            label_visibility="collapsed",
            key="chat_input"
        )
    with col_send:
        send_clicked = st.button("➤ " + T("send"), key="send_btn", use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if send_clicked and user_input and user_input.strip() and user_input != st.session_state.last_processed_text:
        st.session_state.last_processed_text = user_input
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.session_state.thinking = True
        st.rerun()

    # ============ FILE UPLOAD + WORKFLOW - SIDE BY SIDE ============
    col_files, col_wf = st.columns([1, 1])

    with col_files:
        st.markdown(f'<div class="scard"><div class="scard-title">{T("files_title")}</div>', unsafe_allow_html=True)
        
        excel_file = st.file_uploader(T("excel_lbl"), type=None, key="xl")
        pdf_file = st.file_uploader(T("pdf_lbl"), type=None, key="pd")
        
        if excel_file and pdf_file and not st.session_state.files_ready:
            with st.spinner(T("processing")):
                excel_bytes_read = excel_file.read()
                pdf_bytes_read = pdf_file.read()
                
                if pdf_file.name.lower().endswith('.pdf'):
                    rev = extract_pdf_revenue(pdf_bytes_read)
                else:
                    try:
                        _, _, _ = process_any_file(excel_file)
                        rev = {"transient": 24000, "monthly": 16000, "total": 40000}
                    except:
                        rev = {"transient": 24000, "monthly": 16000, "total": 40000}
                
                st.session_state.extracted_rev = rev
                st.session_state.excel_bytes = excel_bytes_read
                st.session_state.files_ready = True
                st.session_state.fixed_excel = None
            st.success(T("files_ok"))
            st.rerun()
        
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
            
            col_run, col_clear_wf = st.columns([1, 1])
            with col_run:
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
                            
                            for line in log:
                                st.markdown(f'<div class="log-line">{line}</div>', unsafe_allow_html=True)
                            st.success(T("run_ok"))
                        except Exception as e:
                            st.error(f"Workflow error: {e}")
            
            with col_clear_wf:
                if st.button("🔄 " + T("clear_workflow"), use_container_width=True):
                    st.session_state.extracted_rev = {}
                    st.session_state.excel_bytes = None
                    st.session_state.files_ready = False
                    st.session_state.fixed_excel = None
                    st.session_state.workflow_log = []
                    st.rerun()
            
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
# Process AI response with MEMORY (last 12 messages)
# ─────────────────────────────────────────────────────────────────
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
        
        audio_b64 = text_to_speech(reply)
        st.session_state.last_audio = audio_b64
        st.rerun()

# Play Allison's audio if available
if st.session_state.get("last_audio"):
    play_audio_html(st.session_state.last_audio)
    st.session_state.last_audio = None

# ─────────────────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────────────────
if st.session_state.authenticated:
    page_dashboard()
else:
    page_login()
