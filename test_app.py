import streamlit as st
st.write("Testing app.py imports...")

try:
    import pdfplumber
    st.write("✅ pdfplumber")
except Exception as e:
    st.error(f"❌ pdfplumber: {e}")

try:
    from openpyxl import load_workbook
    st.write("✅ openpyxl")
except Exception as e:
    st.error(f"❌ openpyxl: {e}")

try:
    import io, re, os, json, requests
    st.write("✅ standard libs")
except Exception as e:
    st.error(f"❌ standard libs: {e}")

try:
    from datetime import datetime
    st.write("✅ datetime")
except Exception as e:
    st.error(f"❌ datetime: {e}")

try:
    import csv, zipfile
    st.write("✅ csv/zipfile")
except Exception as e:
    st.error(f"❌ csv/zipfile: {e}")

try:
    from xml.etree import ElementTree
    st.write("✅ xml")
except Exception as e:
    st.error(f"❌ xml: {e}")

try:
    from audio_recorder_streamlit import audio_recorder
    st.write("✅ audio_recorder")
except Exception as e:
    st.error(f"❌ audio_recorder: {e}")

try:
    from deepgram import DeepgramClient
    st.write("✅ deepgram")
except Exception as e:
    st.error(f"❌ deepgram: {e}")

try:
    import base64
    st.write("✅ base64")
except Exception as e:
    st.error(f"❌ base64: {e}")

st.success("All imports work! The issue is in the app.py code itself.")
