import streamlit as st
import pandas as pd
from google import genai
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

# ==========================================
# 1. API Configuration & Secrets Lock
# ==========================================
YOUR_API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=YOUR_API_KEY)

# ==========================================
# 2. Dynamic Database Loading
# ==========================================
@st.cache_data
def load_data():
    try:
        return pd.read_csv('lesson_master_v1_5.csv')
    except:
        return pd.DataFrame()

# 💡 கணித வெயிட்டேஜ் கால்குலேட்டர் என்ஜின்
def calculate_math_weightage_from_csv(selected_lessons):
    try:
        math_df = pd.read_csv('2023-2026 கணிதம்.csv')
        rules = []
        for lesson in selected_lessons:
            lesson_data = math_df[math_df['Lesson'].str.lower() == lesson.lower()]
            if not lesson_data.empty:
                p1_count = round(len(lesson_data[lesson_data['Part'] == 'I']) / 4, 1)
                p2_count = round(len(lesson_data[lesson_data['Part'] == 'II']) / 4, 1)
                p3_count = round(len(lesson_data[lesson_data['Part'] == 'III']) / 4, 1)
                p4_count = round(len(lesson_data[lesson_data['Part'] == 'IV']) / 4, 1)
                sample_topics = ", ".join(lesson_data['Topic'].unique()[:3])
                rules.append(f"- From '{lesson}' ({sample_topics}): Average {p1_count} Qs in Part-I, {p2_count} Qs in Part-II, {p3_count} Qs in Part-III, and {p4_count} Qs in Part-IV.")
        return "\n".join(rules)
    except:
        return "- Follow standard distribution evenly across selected units."

# 💡 புதிய ஆங்கில போர்டு பேட்டர்ன் இன்டெலிஜென்ஸ் அல்காரிதம்
def get_english_blueprint_rules():
    return """
    [STRICT TN BOARD ENGLISH PATTERN INTEL LOCK]
    1. PART I (Q1-14): 1 Mark Questions ONLY. 
       - Q1-3: Synonyms from selected Prose text.
       - Q4-6: Antonyms from selected Prose text[cite: 4].
       - Q7-14: Structural grammar grids (Plural Form, Prefix/Suffix, Abbreviations, Phrasal Verbs, Compound Words, Prepositions, Tense, Linkers)[cite: 4].
    2. PART II (2 Marks):
       - Section 1: Poetry Appreciation questions[cite: 4].
       - Section 2: Poetic Devices (Rhyme Scheme, Rhyming Words, Alliteration, Figure of Speech)[cite: 4].
       - Section 3: Grammar Transformation (Strictly 4 Blocks: Voice Change, Reported Speech, Punctuation, Simple/Compound/Complex conversion)[cite: 4].
       - Section 4: Road Map Directions[cite: 4].
    3. PART III (5 Marks):
       - Fixed sections for Advertisement, Letter to Editor / Inquiry Letter, Notice Writing, and Picture Comprehension[cite: 4].
    """

def get_blueprint_defaults(total_marks, is_social=False, is_english=False):
    if is_english or is_social:
        return {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 10, "p3a": 7, "p4v": 8, "p4g": 2, "p4a": 2}
    return {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}

# ==========================================
# 3. Adaptive Subject Prompt Engine
# ==========================================
def generate_prompt_v18(subject, lessons_list, exam_type, exam_time, total_marks, exam_mode, blueprint_desc):
    lessons_str = ", ".join(lessons_list)
    sub_lower = subject.lower()
    
    is_english = "english" in sub_lower or "ஆங்கிலம்" in sub_lower
    is_math = "math" in sub_lower or "கணிதம்" in sub_lower
    
    if is_english:
        lang_instruction = "Language: Pure ENGLISH language content only. No translation."
        header_format = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: a) , b) , c) , d)"
        subject_blueprint_rules = get_english_blueprint_rules()
    elif is_math:
        lang_instruction = "Language: Pure TAMIL language only for text prose."
        header_format = "பகுதி [ROMAN_NUM] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = f"""
        [MANDATORY CRITICAL MATHEMATICS RULES]
        - Output tags strictly on its own line: [DRAW_TRIANGLE: AB=5cm, BC=6cm]. No Tamil inside square brackets.
        {calculate_math_weightage_from_csv(lessons_list)}
        """
    else:
        lang_instruction = "Language: Pure TAMIL language only."
        header_format = "பகுதி [ROMAN_NUM]"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = ""

    return f"""
    You are an Expert Question Paper Setter for TN Board Class 10. 
    Generate a professional question paper based strictly on requested constraints.

    Subject: {subject}
    Lessons to Cover: {lessons_str}
    Exam Type: {exam_type}
    Total Marks: {total_marks}
    Time Allowed: {exam_time}
    Exam Mode: {exam_mode}

    [STRICT BLUEPRINT PATTERN]
    {blueprint_desc}

    [STRICT HEADER FORMAT]
    {header_format}
    
    [OPTIONS FORMAT]
    {option_format}

    [{lang_instruction}]

    {subject_blueprint_rules}

    === ANSWER KEY ===
    """

# ==========================================
# 4. Word Document Export Engine
# ==========================================
def set_cell_margins(cell, **kwargs):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m in ["top", "start", "bottom", "end"]:
        if m in kwargs:
            node = OxmlElement(f"w:{m}")
            node.set(qn('w:w'), str(kwargs[m]))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
    tcPr.append(tcMar)

def add_solid_line(doc):
    p = doc.add_paragraph()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    p.paragraph_format.element.get_or_add_pPr().append(pBdr)

def write_markdown_to_word(doc, text):
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        
        clean_line = line.replace('*', '').replace('$', '').strip()
        
        if "பகுதி" in clean_line or "PART" in clean_line.upper():
            marks_match = re.search(r'\(?\d+\s*[xX*]\s*\d+\s*=\s*\d+\)?', clean_line)
            if marks_match:
                calc_str = marks_match.group(0)
                title_str = clean_line.replace(calc_str, "").strip(":- ")
                table = doc.add_table(rows=1, cols=2)
                c1, c2 = table.rows[0].cells
                c1.paragraphs[0].add_run(title_str).bold = True
                c2.paragraphs[0].add_run(calc_str).bold = True
                c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                continue

        p = doc.add_paragraph()
        if re.match(r'^\d+\.', clean_line):
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)
        p.add_run(clean_line)

def create_professional_docx(ai_response, school_name, class_val, subject_val, exam_type, time_val, marks_val):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.27), Inches(11.69)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.5)
    
    style = doc.styles['Normal']
    style.font.name = 'Nirmala UI'
    
    h_school = doc.add_paragraph()
    h_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_school.add_run(school_name.upper()).bold = True
    h_school.runs[0].font.size = Pt(14)
    
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).paragraphs[0].add_run(f"Class: {class_val} | Subject: {subject_val}")
    table.cell(0, 1).paragraphs[0].add_run(f"Time: {time_val}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table.cell(1, 0).paragraphs[0].add_run(f"Exam: {exam_type}")
    table.cell(1, 1).paragraphs[0].add_run(f"Total Marks: {marks_val}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    add_solid_line(doc)
    parts = ai_response.split("=== ANSWER KEY ===")
    write_markdown_to_word(doc, parts[0].strip())
    
    if len(parts) > 1:
        doc.add_page_break()
        doc.add_paragraph().add_run("ANSWER KEY / விடைகள்").bold = True
        add_solid_line(doc)
        write_markdown_to_word(doc, parts[1].strip())
        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==========================================
# 5. Streamlit Presentation Interface Layer
# ==========================================
st.set_page_config(page_title="PMP Multi-Subject Suite", page_icon="🎓", layout="wide")

st.title("🎓 PMP Multi-Subject Intelligence AI (V18.1 LIVE)")
df = load_data()

if not df.empty:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        school_name = st.text_input("School Name", value="SARAVANA ACADEMY")
        class_val = st.selectbox("Class", ["10"])
    with col2:
        subject_list = df['Subject'].unique()
        subject_val = st.selectbox("Subject", subject_list)
        exam_type = st.selectbox("Exam Type", ["Unit Test", "Quarterly Exam", "Half-Yearly Exam", "Public Revision"])
    with col3:
        time_val = st.selectbox("Time Allowed", ["1.00 Hour", "2.00 Hours", "3.00 Hours"], index=2)
    with col4:
        marks_val = st.number_input("Total Marks", value=100, step=1)
        exam_mode = st.selectbox("Exam Mode", ["🏛️ Public Exam Mode", "🏫 School Elite Mode"])
        
    lesson_list = df[df['Subject'] == subject_val]['Lesson'].unique()
    selected_lessons = st.multiselect("Select Chapters / அத்தியாயங்கள்", lesson_list)
    
    st.markdown("---")
    
    is_eng = "english" in subject_val.lower() or "ஆங்கிலம்" in subject_val.lower()
    is_soc = "social" in subject_val.lower() or "சமூக" in subject_val.lower()
    bp = get_blueprint_defaults(marks_val, is_social=is_soc, is_english=is_eng)
    
    st.markdown("### 📋 வினா வடிவமைப்பு தானியங்கிப் பிரிவு (Dynamic Blueprint Lock)")
    b_col1, b_col2, b_col3, b_col4 = st.columns(4)
    with b_col1:
        p1_ask = st.number_input("Part I (1-Mark Qs)", value=int(bp["p1"]), step=1)
    with b_col2:
        p2_get = st.number_input("Part II Given", value=int(bp["p2g"]), step=1)
        p2_ask = st.number_input("Part II Answer", value=int(bp["p2a"]), step=1)
    with b_col3:
        p3_get = st.number_input("Part III Given", value=int(bp["p3g"]), step=1)
        p3_ask = st.number_input("Part III Answer", value=int(bp["p3a"]), step=1)
    with b_col4:
        p4_val = st.selectbox("Long Question Marks", [5, 8], index=1 if is_eng or is_soc or marks_val==100 else 0)
        p4_get = st.number_input("Part IV Given", value=int(bp["p4g"]), step=1)
        p4_ask = st.number_input("Part IV Answer", value=int(bp["p4a"]), step=1)
        
    if st.button("🚀 Generate Data-Driven Question Paper", use_container_width=True):
        if not selected_lessons:
            st.warning("⚠️ தயவுசெய்து பாடங்களைத் தேர்ந்தெடுக்கவும்.")
        else:
            with st.spinner("⏳ 2023-2026 அசல் பொதுத்தேர்வு விதிகளின்படி வினாத்தாள் தயாராகிறது..."):
                blueprint_desc = f"Part I: {p1_ask} Qs. Part II: Answer {p2_ask}/{p2_get}. Part III: Answer {p3_ask}/{p3_get}. Part IV: Answer {p4_ask}/{p4_get} ({p4_val} Marks)."
                prompt = generate_prompt_v18(subject_val, selected_lessons, exam_type, time_val, marks_val, exam_mode, blueprint_desc)
                try:
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                    docx_file = create_professional_docx(response.text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                    st.session_state['docx_bytes_v18'] = docx_file.getvalue()
                    st.success("✅ கணிதம் + ஆங்கிலம் ஒருங்கிணைந்த மாஸ்டர் வினாத்தாள் தயாராக உள்ளது!")
                except Exception as e:
                    st.error(f"பிழை: {e}")

    if 'docx_bytes_v18' in st.session_state:
        st.download_button(label="📥 Download Master Word Document (.docx)", data=st.session_state['docx_bytes_v18'], file_name=f"PMP_Master_{subject_val}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
