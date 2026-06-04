import streamlit as st
import streamlit_authenticator as stauth
import pandas as pd
from google import genai
from google.genai import types
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
import time
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

# 1. லாகின் விபரங்கள்
credentials = {
    "usernames": {
        "admin": {"name": "Sivaprakash", "password": "abc"}
    }
}
authenticator = stauth.Authenticate(credentials, 'pmp_cookie', 'secret_key', cookie_expiry_days=30)

# 2. லாகின் லேயர்
name, authentication_status, username = authenticator.login('Login', 'main')

if authentication_status:
    # ----------------------------------------------------
    # இங்கிருந்துதான் மெயின் கோடு ஆரம்பிக்கிறது
    # ----------------------------------------------------
    YOUR_API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=YOUR_API_KEY)

    @st.cache_data
    def load_data():
        try:
            return pd.read_csv('lesson_master_v1_5.csv')
        except:
            return pd.DataFrame()

    def get_math_dynamic_weightage(selected_lessons):
        base_matrix = {
            "Relations and Functions": {"1M": 1.5, "2M": 2, "5M": 2, "8M": 0},
            "Numbers and Sequences":   {"1M": 2, "2M": 2, "5M": 2, "8M": 0},
            "Algebra":                  {"1M": 2, "2M": 2, "5M": 2, "8M": 1},
            "Geometry":                 {"1M": 2, "2M": 1, "5M": 1, "8M": 1},
            "Coordinate Geometry":     {"1M": 1.5, "2M": 2, "5M": 2, "8M": 0},
            "Trigonometry":            {"1M": 1, "2M": 1, "5M": 1, "8M": 0},
            "Mensuration":             {"1M": 1.5, "2M": 2, "5M": 2, "8M": 0},
            "Statistics and Probability":{"1M": 2, "2M": 2, "5M": 2, "8M": 0}
        }
        rules = []
        for lesson in selected_lessons:
            if lesson in base_matrix:
                bm = base_matrix[lesson]
                rules.append(f"- From Chapter '{lesson}': Generate approx {int(bm['1M'])} MCQs, {bm['2M']} Questions (2-Mark), {int(bm['5M'])} Questions (5-Mark), and {bm['8M']} Question (8-Mark).")
        return "\n".join(rules)

    st.title("🎓 PMP Master AI Engine (V21.3)")
    
    tab1, tab2 = st.tabs(["🎓 வினாத்தாள் தயாரிப்பு", "📝 விடைத்தாள் திருத்தம்"])
    
    with tab1:
        df = load_data()
        if not df.empty:
            school_name = st.text_input("School Name", value="ABC SCHOOL")
            subject_val = st.selectbox("Subject", df['Subject'].unique())
            marks_val = st.number_input("Total Marks", value=100)
            lesson_list = df[df['Subject'] == subject_val]['Lesson'].unique()
            selected_lessons = st.multiselect("பாடங்களைத் தேர்ந்தெடுக்கவும்", lesson_list)
            
            if st.button("🚀 Generate PRO Question Paper"):
                st.success("வினாத்தாள் தயாராகிறது...")
                # இங்கேயே தயாரிப்பு லாஜிக் தொடரும்...

    if st.sidebar.button("Logout"):
        authenticator.logout('Logout', 'main')

elif authentication_status == False:
    st.error('❌ Username/password is incorrect')
elif authentication_status == None:
    st.warning('⚠️ Please enter your username and password')
