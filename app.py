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

# 💡 சர்வர் மெமரி மற்றும் கிராஷ் பாதுகாப்பிற்கான பிரத்யேக மேட்லாட்லிப் லாக்
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
# 2. Database Loading
# ==========================================
@st.cache_data
def load_data():
    try:
        return pd.read_csv('lesson_master_v1_5.csv')
    except:
        return pd.DataFrame()

# 💡 விகிதாச்சாரத் தானியங்கிப் பாதுகாப்பு இன்ஜின் (Proportion Auto-Scaler)
def get_blueprint_defaults(total_marks, is_social=False):
    if is_social:
        return {"p1": 20, "p2g": 12, "p2a": 10, "p3g": 10, "p3a": 8, "p4v": 8, "p4g": 4, "p4a": 2}
    
    defaults = {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    if total_marks == 106:
        defaults = {"p1": 20, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    elif total_marks == 50:
        defaults = {"p1": 10, "p2g": 8, "p2a": 6, "p3g": 6, "p3a": 4, "p4v": 8, "p4g": 2, "p4a": 1}
    elif total_marks == 25:
        defaults = {"p1": 5, "p2g": 6, "p2a": 5, "p3g": 3, "p3a": 2, "p4v": 8, "p4g": 0, "p4a": 0}
    return defaults

# 💡 பைத்தானில் தானாகவே முக்கோணம்/சதுரம் வரைந்து தரும் மேஜிக் இன்ஜின்
def generate_geometry_image(shape_type, label_text=""):
    fig, ax = plt.subplots(figsize=(2.5, 2.5))
    shape_upper = shape_type.upper()
    
    if "TRIANGLE" in shape_upper:
        # அக்யூரேட் முக்கோணம் புள்ளிகள்
        points = np.array([[0, 0], [4, 0], [2, 3], [0, 0]])
        ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
        ax.text(-0.2, -0.2, 'A', fontsize=11, fontweight='bold')
        ax.text(4.1, -0.2, 'B', fontsize=11, fontweight='bold')
        ax.text(2, 3.2, 'C', fontsize=11, fontweight='bold')
        if label_text:
            ax.text(2, -0.5, label_text, fontsize=10, ha='center', fontweight='bold', color='blue')
            
    elif "SQUARE" in shape_upper:
        # அக்யூரேட் சதுரம் புள்ளிகள்
        points = np.array([[0, 0], [3, 0], [3, 3], [0, 3], [0, 0]])
        ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
        ax.text(-0.2, -0.2, 'A', fontsize=10)
        ax.text(3.2, -0.2, 'B', fontsize=10)
        ax.text(3.2, 3.2, 'C', fontsize=10)
        ax.text(-0.2, 3.2, 'D', fontsize=10)
        if label_text:
            ax.text(1.5, -0.5, label_text, fontsize=10, ha='center', fontweight='bold', color='blue')
            
    ax.set_aspect('equal')
    ax.axis('off')
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=100)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

# ==========================================
# 3. Adaptive Language & Subject Prompt Engine
# ==========================================
def generate_prompt_v16(subject, lessons_list, exam_type, exam_time, total_marks, exam_mode, blueprint_desc, custom_q, book_back_perf, interior_perf):
    lessons_str = ", ".join(lessons_list)
    sub_lower = subject.lower()
    
    is_english = "english" in sub_lower or "ஆங்கிலம்" in sub_lower
    is_tamil = "tamil" in sub_lower or "தமிழ்" in sub_lower
    is_social = "social" in sub_lower or "சமூக" in sub_lower
    is_math = "math" in sub_lower or "கணிதம்" in sub_lower
    
    if is_english:
        lang_instruction = "5. Language: The ENTIRE question paper text, headers, instructions, questions, and options MUST be in pure ENGLISH language only."
        header_format = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: a) , b) , c) , d)"
        subject_blueprint_rules = """
        [STRICT TN BOARD ENGLISH BLUEPRINT - MANDATORY]
        - PART I (20 MCQs): Q1-8 Vocabulary, Q9-12 Word Formation, Q13-18 Grammar, Q19-20 Textbook Facts.
        - PART II (12 Qs, Answer 10): 4 Recall, 4 Understanding, 4 Grammar Application.
        """
    elif is_tamil:
        lang_instruction = "5. Language: The ENTIRE question paper text MUST be in pure TAMIL only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (வினாக்கள் எண்ணிக்கை x மதிப்பெண் = மொத்த மதிப்பெண்கள்)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = """
        [அசல் தமிழ் பாடத்திட்ட ப்ளூபிரின்ட்]
        - பகுதி I (20 பலவுள் தெரிக): Q1-7 சொல்வளம், Q8-14 இலக்கணம், Q15-20 இலக்கியம்.
        - பகுதி II (12 குறுவினாக்கள், எழுதுக 10): 6 பாடப்பகுதி வினாக்கள், 2 இலக்கிய வினாக்கள், 4 இலக்கணப் பயிற்சிகள்.
        """
    elif is_social:
        lang_instruction = "5. Language: The ENTIRE question paper text MUST be in pure TAMIL only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (வினாக்கள் எண்ணிக்கை x மதிப்பெண் = மொத்த மதிப்பெண்கள்)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = """
        [MANDATORY CRITICAL SOCIAL SCIENCE BOARD EXPERT BLUEPRINT]
        - PART I: 15 Recall MCQs + 5 Application MCQs (Assertion-Reason, Match, Chronology).
        - PART III: 2 Source-Based Passage Questions, 2 Map Marking Lists, 2 HOTS Questions.
        - PART IV: 2 Traditional essays + 2 Competency-Based Case studies.
        """
    elif is_math:
        lang_instruction = "5. Language: The ENTIRE question paper text MUST be in pure TAMIL only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        
        # 💡 கணிதப் பாடத்திற்கான பிரத்தியேக டயக்ராம் இஞ்சின் லாக் விதிமுறை
        subject_blueprint_rules = """
        [MANDATORY CRITICAL MATHEMATICS EMBEDDED DIAGRAM RULE]
        1. ABSOLUTE BAN ON AI DISCLAIMERS: Do NOT ever output phrases like '[வரைபடத்தை இங்கு காட்சிப்படுத்த முடியாது]' or 'As an AI, I cannot draw'.
        2. DYNAMIC GEOMETRY TAGS: If a geometry construction question (e.g., Tangents, Triangles) or its answer key requires a visual diagram, you MUST output a specific text tag strictly on its own line:
           - For Triangles: [DRAW_TRIANGLE: AB = 6 cm, BC = 8 cm]
           - For Squares: [DRAW_SQUARE: பக்கம் = 5 செ.மீ]
           Our internal python pipeline will automatically replace this text tag with a perfectly drawn high-resolution image inside the Word document.
        3. GRAPH PAPER COORDINATES: For graph equations, output a clean markdown coordinates table showing x and y values.
        """
    else:
        lang_instruction = "5. Language: Pure TAMIL language only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்]"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = "[CORE BLUEPRINT] Split evenly across standard textbook parts."

    return f"""
    You are an Expert Question Paper Setter for TN Board Class 10. 
    Generate a professional question paper based strictly on the requested constraints and the specific subject's blueprint style.

    Subject: {subject}
    Lessons to Cover: {lessons_str}
    Exam Type: {exam_type}
    Total Marks: {total_marks}
    Time Allowed: {exam_time}
    Exam Mode: {exam_mode}

    [STRICT BLUEPRINT CHOICE PATTERN - MANDATORY RULE]
    {blueprint_desc}

    [STRICT HEADER FORMAT FOR WORD TABLES]
    {header_format}
    
    [CRITICAL OPTIONS FORMAT]
    {option_format}

    [STRICT NO-LATEX RULE FOR WORD TEMPLATE]
    Do NOT use '$' or '$$' or '\\frac' or any LaTeX symbols inside this Tab 1 Question Paper Generator. Write equations in plain, normal text format.

    [LANGUAGE DIRECTIVE]
    {lang_instruction}

    {subject_blueprint_rules}

    === ANSWER KEY ===
    """

# ==========================================
# 4. Word Document Export Core Engine (With Image Insertion Logic)
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
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)

def write_markdown_to_word(doc, text):
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        
        # 💡 அக்யூரேட் டயக்ராம் குறியீடுகளைத் தேடிப் பிடித்து இமேஜாக மாற்றும் பகுதி
        draw_match = re.search(r'\[DRAW_(TRIANGLE|SQUARE)[:\s]*(.*?)\]', line, re.IGNORECASE)
        if draw_match:
            shape_type = draw_match.group(1)
            label_text = draw_match.group(2)
            try:
                img_buf = generate_geometry_image(shape_type, label_text)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(6)
                p_img.paragraph_format.space_after = Pt(6)
                p_img.add_run().add_picture(img_buf, width=Inches(2.5))
            except Exception as e:
                doc.add_paragraph(f"[வரைபடம் இணைப்பதில் பிழை: {e}]")
            continue

        clean_line = line.replace('*', '').replace('$', '').strip()
        
        if "பகுதி" in clean_line or "PART" in clean_line.upper():
            marks_match = re.search(r'\(?\d+\s*[xX*]\s*\d+\s*=\s*\d+\)?', clean_line)
            if marks_match:
                calc_str = marks_match.group(0)
                title_str = clean_line.replace(calc_str, "").strip(":- ")
                table = doc.add_table(rows=1, cols=2)
                c1, c2 = table.rows[0].cells
                c1.paragraphs[0].add_run(title_str).bold = True
                c2.paragraphs[0].add_run(f"({calc_str})" if not calc_str.startswith("(") else calc_str).bold = True
                c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_cell_margins(c1, top=0, bottom=0, start=0, end=0)
                set_cell_margins(c2, top=0, bottom=0, start=0, end=0)
                continue

        option_markers = ["அ)", "ஆ)", "இ)", "ஈ)", "a)", "b)", "c)", "d)"]
        if any(marker in clean_line for marker in option_markers):
            raw_parts = re.split(r'(அ\)|ஆ\)|இ\)|ஈ\)|a\)|b\)|c\)|d\))', clean_line)
            parts = []
            current = ""
            for chunk in raw_parts:
                if chunk in option_markers:
                    if current.strip(): parts.append(current.strip())
                    current = chunk + " "
                else: current += chunk
            if current.strip(): parts.append(current.strip())
            
            if parts:
                table = doc.add_table(rows=1, cols=len(parts))
                for idx, opt in enumerate(parts):
                    cell = table.cell(0, idx)
                    cell.paragraphs[0].add_run(opt.replace("*", ""))
                    set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
                continue

        p = doc.add_paragraph()
        if re.match(r'^\d+\.', clean_line):
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)
        parts = re.split(r'\*\*(.*?)\*\*', line)
        for i, part in enumerate(parts):
            run = p.add_run(part.replace('$', ''))
            if i % 2 == 1: run.bold = True

def create_professional_docx(ai_response, school_name, class_val, subject_val, exam_type, time_val, marks_val):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.27), Inches(11.69)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.5)
    
    style = doc.styles['Normal']
    style.font.name = 'Nirmala UI'
    
    h_school = doc.add_paragraph(style='Normal')
    h_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_school.add_run(school_name.upper()).bold = True
    h_school.runs[0].font.size = Pt(15)
    
    h_exam = doc.add_paragraph()
    h_exam.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_exam.add_run(exam_type).bold = True
    
    table = doc.add_table(rows=2, cols=2)
    def format_cell(cell, text, align_right=False):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align_right else WD_ALIGN_PARAGRAPH.LEFT
        p.add_run(text).bold = True
        set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
        
    format_cell(table.cell(0, 0), f"வகுப்பு / Class : {class_val}")
    format_cell(table.cell(0, 1), f"நேரம் / Time : {time_val}", align_right=True)
    format_cell(table.cell(1, 0), f"பாடம் / Subject : {subject_val}")
    format_cell(table.cell(1, 1), f"மதிப்பெண்கள் / Marks : {marks_val}", align_right=True)
    
    add_solid_line(doc)
    parts = ai_response.split("=== ANSWER KEY ===")
    write_markdown_to_word(doc, parts[0].strip())
    
    if len(parts) > 1:
        doc.add_page_break()
        p_ak = doc.add_paragraph()
        p_ak.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ak.add_run("விடைகள் (ANSWER KEY)").bold = True
        add_solid_line(doc)
        write_markdown_to_word(doc, parts[1].strip())
        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==========================================
# 5. Streamlit Presentation UI Layer
# ==========================================
st.set_page_config(page_title="PMP AI Suite PRO", page_icon="🎓", layout="wide")
st.markdown("""<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}</style>""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🎓 வினாத்தாள் தயாரிப்பு (QP Generator)", "📝 விடைத்தாள் திருத்தம் (AI Math Evaluator)"])

# ------------------------------------------
# TAB 1: Question Paper Generator
# ------------------------------------------
with tab1:
    st.title("🎓 PMP Question Paper AI (V17.7 DIAGRAM ENGINE)")
    df = load_data()
    if df.empty:
        st.error("Database (lesson_master_v1_5.csv) கிடைக்கவில்லை!")
    else:
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            school_name = st.text_input("School Name", value="ABC SCHOOL")
            class_val = st.selectbox("Class", ["10"])
        with col2:
            subject_list = df['Subject'].unique()
            subject_val = st.selectbox("Subject", subject_list)
            exam_type = st.selectbox("Exam Type", ["Unit Test", "Revision Test", "Quarterly Exam", "Half-Yearly Exam", "Annual Exam"])
        with col3:
            time_val = st.selectbox("Time (நேரம்)", ["1.00 Hour", "1.30 Hours", "2.00 Hours", "2.30 Hours", "3.00 Hours"], index=3)
        with col4:
            marks_val = st.number_input("Total Marks (மொத்த மதிப்பெண்கள்)", value=106, step=1)
            exam_mode = st.selectbox("Exam Mode", ["🏛️ Public Exam Mode", "🏫 School Elite Mode"])
            
        lesson_list = df[df['Subject'] == subject_val]['Lesson'].unique()
        selected_lessons = st.multiselect("பாடம் / பாடங்களைத் தேர்ந்தெடுக்கவும்", lesson_list)
        
        st.markdown("---")
        
        is_soc = "social" in subject_val.lower() or "சமூக" in subject_val.lower()
        bp = get_blueprint_defaults(marks_val, is_social=is_soc)
        
        st.markdown("### 📋 வினா வடிவமைப்பு தானியங்கிப் பிரிவு (Auto-Adjusted Blueprint Options)")
        
        b_col1, b_col2, b_col3, b_col4 = st.columns(4)
        with b_col1:
            st.subheader("பகுதி I (1-Mark)")
            p1_ask = st.number_input("1-மார்க் வினாக்கள் எண்ணிக்கை", value=int(bp["p1"]), step=1)
        with b_col2:
            st.subheader("பகுதி II (2-Mark)")
            p2_get = st.number_input("2-மார்க் கொடுக்க வேண்டியவை (Given)", value=int(bp["p2g"]), step=1)
            p2_ask = st.number_input("2-மார்க் எழுத வேண்டியவை (Answer)", value=int(bp["p2a"]), step=1)
        with b_col3:
            st.subheader("பகுதி III (5-Mark)")
            p3_get = st.number_input("5-மார்க் கொடுக்க வேண்டியவை (Given)", value=int(bp["p3g"]), step=1)
            p3_ask = st.number_input("5-மார்க் எழுத வேண்டியவை (Answer)", value=int(bp["p3a"]), step=1)
        with b_col4:
            st.subheader("பகுதி IV (Long Qs)")
            p4_val = st.selectbox("நெடுவினா மதிப்பெண் (Per Question)", [8, 10], index=0 if bp["p4v"] == 8 else 1)
            p4_get = st.number_input("நெடுவினா கொடுக்க வேண்டியவை (Given)", value=int(bp["p4g"]), step=1)
            p4_ask = st.number_input("நெடுவினா எழுத வேண்டியவை (Answer)", value=int(bp["p4a"]), step=1)
            
        total_calculated = (p1_ask * 1) + (p2_ask * 2) + (p3_ask * 5) + (p4_ask * p4_val)
        
        if total_calculated == marks_val:
            st.success(f"✅ வெற்றிகரமாகப் பொருந்தியது! விகிதாச்சாரக் கணக்கீடு: {total_calculated} மார்க் = மொத்த மதிப்பெண்: {marks_val} மார்க்.")
            can_generate = True
        else:
            st.warning(f"⚠️ கணக்கீடு: {total_calculated} மார்க் | மொத்த மதிப்பெண்: {marks_val} மார்க். (சமமாக மாற்றவும்).")
            can_generate = False
            
        st.markdown("---")
        
        if st.button("🚀 Generate PRO Question Paper", use_container_width=True):
            if not can_generate:
                st.error("⚠️ மதிப்பெண்கள் சரியாகப் பொருந்தவில்லை!")
            elif not selected_lessons:
                st.warning("⚠️ தயவுசெய்து பாடங்களைத் தேர்ந்தெடுக்கவும்.")
            else:
                spinner_text = "⏳ அசல் அரசுப் பொதுத்தேர்வு தரத்தில் வினாத்தாள் தயாராகிறது..." if "Public" in exam_mode else "⏳ எலைட் போர்டு தரத்தில் வினாத்தாள் தயாராகிறது..."
                with st.spinner(spinner_text):
                    dynamic_blueprint_desc = f"- Part I: 1 Mark Questions Total: {p1_ask}. - Part II: Given {p2_get}, Answer {p2_ask}. - Part III: Given {p3_get}, Answer {p3_ask}. - Part IV: {p4_val} Mark Given {p4_get}, Answer {p4_ask}."
                    prompt = generate_prompt_v16(subject_val, selected_lessons, exam_type, time_val, marks_val, exam_mode, dynamic_blueprint_desc, "", 80, 20)
                    try:
                        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                        docx_file = create_professional_docx(response.text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                        st.session_state['ai_text'] = response.text
                        st.session_state['docx_bytes'] = docx_file.getvalue()
                        st.session_state['file_name'] = f"PMP_{subject_val}.docx"
                        st.success("✅ வினாத்தாள் வெற்றிகரமாகத் தயாராகிவிட்டது!")
                    except Exception as e:
                        st.error(f"பிழை: {e}")

        if 'ai_text' in st.session_state:
            st.download_button(label="📥 Download as Word File (.docx)", data=st.session_state['docx_bytes'], file_name=st.session_state['file_name'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# ------------------------------------------
# TAB 2: AI Handwriting Paper Evaluator
# ------------------------------------------
with tab2:
    st.title("📝 AI Math Paper Evaluator (LaTeX Edition)")
    uploaded_image = st.file_uploader("உங்கள் கையெழுத்து கணிதப் பக்கத்தை அப்லோட் செய்யவும்", type=["png", "jpg", "jpeg"])
    if uploaded_image is not None:
        image = Image.open(uploaded_image)
        st.image(image, caption="Uploaded Image", width=450)
        if st.button("🚀 Start AI Evaluation", use_container_width=True):
            with st.spinner("⏳ ஜெமினி AI விடைத்தாளைத் திருத்தி வருகிறது..."):
                eval_prompt = "You are an official TN Board Math Evaluator. Read handwriting and correct step-by-step. Write fractions as $\\frac{a}{b}$ and roots as $\\sqrt{x}$ inside single dollar signs. Respond in Tamil."
                try:
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=[image, eval_prompt])
                    st.markdown(response.text)
                except Exception as e:
                    st.error(f"❌ சர்வர் பிழை: {e}")
