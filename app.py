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
def get_blueprint_defaults(total_marks):
    # Fallback Standard (100 Marks)
    defaults = {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    
    if total_marks == 106:
        defaults = {"p1": 20, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    elif total_marks == 100:
        defaults = {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    elif total_marks == 90:
        defaults = {"p1": 14, "p2g": 11, "p2a": 9, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 2, "p4a": 1}
    elif total_marks == 80:
        defaults = {"p1": 12, "p2g": 11, "p2a": 9, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 0, "p4a": 0}
    elif total_marks == 75:
        defaults = {"p1": 11, "p2g": 9, "p2a": 7, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 0, "p4a": 0}
    elif total_marks == 50:
        defaults = {"p1": 10, "p2g": 8, "p2a": 6, "p3g": 6, "p3a": 4, "p4v": 8, "p4g": 2, "p4a": 1}
    elif total_marks == 25:
        defaults = {"p1": 5, "p2g": 6, "p2a": 5, "p3g": 3, "p3a": 2, "p4v": 8, "p4g": 0, "p4a": 0}
    else:
        # எந்தவொரு புதிய எண்ணுக்கும் துல்லியமான விகிதப் பகிர்வு அல்காரிதம்
        rem = total_marks
        p4a = max(0, int((total_marks * 0.15) / 8))
        rem -= (p4a * 8)
        p3a = max(0, int((rem * 0.60) / 5))
        rem -= (p3a * 5)
        p2a = max(0, int(rem / 3))
        rem -= (p2a * 2)
        p1 = rem
        defaults = {"p1": p1, "p2g": p2a + 2, "p2a": p2a, "p3g": p3a + 4, "p3a": p3a, "p4v": 8, "p4g": p4a + 2, "p4a": p4a}
    return defaults

# ==========================================
# 3. Adaptive Language Prompt Engine 
# ==========================================
def generate_prompt_v16(subject, lessons_list, exam_type, exam_time, total_marks, exam_mode, blueprint_desc, custom_q, book_back_perf, interior_perf):
    lessons_str = ", ".join(lessons_list)
    sub_lower = subject.lower()
    
    is_english = "english" in sub_lower or "ஆங்கிலம்" in sub_lower
    is_tamil = "tamil" in sub_lower or "தமிழ்" in sub_lower
    
    if is_english:
        lang_instruction = "5. Language: The ENTIRE question paper text, headers, instructions, questions, and options MUST be in pure ENGLISH language only."
        header_format = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)"
        option_format = "For every MCQ question, write the 4 options on a single new line using standard markers: a) , b) , c) , d)"
        
        subject_blueprint_rules = """
        [STRICT TN BOARD ENGLISH BLUEPRINT - MANDATORY]
        1. PART I (1-MARK MCQs): Dynamically distribute the generated questions following this strict ratio:
           - Vocabulary: Synonyms, Antonyms, Idiom, Phrasal Verb.
           - Word Formation: Prefix, Suffix, Compound Word, British/American English.
           - Grammar: Tense, Modal, Preposition, Question Tag, Sentence Pattern, Figure of Speech.
           - Textbook: Core Lesson Facts.
        2. PART II (2-MARK QUESTIONS): Strictly split into 3 core professional tiers:
           - Tier 1: Lesson Recall (Comprehension checks).
           - Tier 2: Lesson Understanding (Reasoning/Contextual depth).
           - Tier 3: Grammar Application (Voice, Reported Speech, Punctuation, Combining sentences).
           - Cognitive Target: 40% Recall, 40% Understanding, 20% Life Application.
        """
        
    elif is_tamil:
        lang_instruction = "5. Language: The ENTIRE question paper text, headers, instructions, questions, and options MUST be in pure TAMIL language only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (வினாக்கள் எண்ணிக்கை x மதிப்பெண் = மொத்த மதிப்பெண்கள்)"
        option_format = "ஒவ்வொரு கொள்குறி வினாவிற்கும் 4 விடைகளையும் ஒரே வரியில் இந்த குறியீடுகளைப் பயன்படுத்தி எழுதவும்: அ) , ஆ) , இ) , ஈ)"
        
        subject_blueprint_rules = """
        [அசல் கொள்குறி வகை மற்றும் குறுவினா தமிழ் ப்ளூபிரின்ட் - கட்டாயம்]
        1. பகுதி I (பலவுள் தெரிக): வினாக்களை இந்த விகிதாச்சாரப்படி துல்லியமாகப் பிரித்து வழங்கவும்:
           - சொல்வளம்: சொல் பொருள், எதிர்ச்சொல், இணைச்சொற்கள்.
           - இலக்கணம்: பெயரெச்சம், வினையெச்சம், வினையாலணையும் பெயர், இடைச்சொல், உரிச்சொல், இலக்கணக் குறிப்புகள்.
           - இலக்கியம் & பாடப்பகுதி: நூல்-ஆசிரியர், வரலாற்றுத் தகவல்கள், உவமை/அணி இலக்கணம், பழமொழி நிரப்புதல்.
        2. பகுதி II (2-மதிப்பெண் குறுவினாக்கள்):
           - வினாக்களை 3 பிரிவுகளாகப் பிரிக்கவும்: பாடப்பகுதி வினாக்கள் (உரைநடை/செய்யுள்/துணைப்பாடம் - 40% Recall, 40% Understanding, 20% Application), இலக்கிய வினாக்கள் (நூல்/ஆசிரியர்), இலக்கணப் பயன்பாடு (மொழிப்பயிற்சி/பிழை திருத்தம்).
        """
        
    else:
        lang_instruction = "5. Language: Pure TAMIL language only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (No_of_Qs x Marks = Total_Marks)"
        option_format = "For every MCQ question, write the 4 options on a single new line using standard markers: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = "[CORE SUBJECTS BLUEPRINT] Split into Part I (MCQs), Part II (2-Marks), Part III (5-Marks), Part IV (Long Qs)."

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
# 4. Word Document Export Core Engine 
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
    style.font.size = Pt(11)
    
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
    st.title("🎓 PMP Question Paper AI (V17.5 AUTO-PROPORTION)")
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
            # 💡 ஆசிரியர் இங்கே எண்களை மாற்றினால் கீழே விகிதாச்சாரம் தானாக மாறும்!
            marks_val = st.number_input("Total Marks (மொத்த மதிப்பெண்கள்)", value=106, step=1)
            exam_mode = st.selectbox("Exam Mode", ["🏛️ Public Exam Mode", "🏫 School Elite Mode"])
            
        lesson_list = df[df['Subject'] == subject_val]['Lesson'].unique()
        selected_lessons = st.multiselect("பாடம் / பாடங்களைத் தேர்ந்தெடுக்கவும்", lesson_list)
        
        st.markdown("---")
        
        # 💡 விகிதாச்சாரப் பாதுகாப்பு இன்ஜினை UI உடன் இணைத்தல்
        bp = get_blueprint_defaults(marks_val)
        
        st.markdown("### 📋 வினா வடிவமைப்பு தானியங்கிப் பிரிவு (Auto-Adjusted Blueprint Options)")
        st.caption("💡 குறிப்பு: நீங்கள் மேலே உள்ள 'Total Marks'-ஐ மாற்றும்போது, கீழே உள்ள வினாக்களின் எண்ணிக்கை அசல் விகிதப்படி தானாகவே மாறிக்கொள்ளும்.")
        
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
            st.warning(f"⚠️ கைகளால் எண்களை மாற்றியுள்ளீர்கள்! கணக்கீடு: {total_calculated} மார்க் | மொத்த மதிப்பெண்: {marks_val} மார்க். (சமமாக மாற்றவும்).")
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
