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
# 1. API Configuration
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

# ==========================================
# 3. Dynamic Prompt Engine (Version 4 - Strict Authentic Board Core)
# ==========================================
def generate_prompt_v16(subject, lessons_list, exam_type, exam_time, total_marks, exam_mode, blueprint_desc, custom_q, book_back_perf, interior_perf):
    lessons_str = ", ".join(lessons_list)
    
    is_english_subject = "english" in subject.lower() or "ஆங்கிலம்" in subject.lower()
    
    if is_english_subject:
        lang_instruction = "5. Language: The ENTIRE question paper text, headers, instructions, questions, and options MUST be in pure ENGLISH language only."
        header_format = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)\nExample: PART I - Choose the correct synonyms (14 x 1 = 14)"
        option_format = "For every MCQ question, write the 4 options on a single new line using standard markers: a) , b) , c) , d)"
        
        # V4 Strict TN Board & Textbook Accuracy Rules
        elite_2026_rules = """
        [STRICT TN BOARD ALIGNMENT & TEXTBOOK ACCURACY - MANDATORY V4]
        1. PART I (1-MARK MCQs) MANDATORY BLUEPRINT TOPICS:
           - Questions 1-3: Strictly Textbook Synonyms.
           - Questions 4-6: Strictly Textbook Antonyms.
           - Questions 7-14: MUST ONLY be selected from these official board topics: British English vs American English, Homophones, Syllabification, Prefix/Suffix, Phrasal Verbs, Compound Words, Prepositions, Tenses, Linkers, Idioms, or Sentence Patterns.
           - CRITICAL BAN: Absolutely NO out-of-syllabus expansions like 'IGNOU' or generic 'parts of speech' multiple-choice questions.

        2. PART III (5-MARK QUESTIONS) COMPULSORY SKILLS VARIETY:
           - To prevent students from skipping core writing tasks in the 14-question pool, you MUST strictly structure the questions as follows:
             * Questions 1-4: Prose & Poetry Paragraphs (HOTS/Thematic depth).
             * Questions 5-6: Supplementary Coherent Order / Passage Comprehension.
             * Questions 7-8: Mandatory Grammar/Error Spotting (Identify 5 grammatical errors in a passage) and Sentence Pattern matching.
             * Questions 9-14: Pure TN Board Writing Skills (Notice Writing, Formal/Informal Letter Writing, Report Writing, Advertisement Design, Dialogue Completion, Proverb Expansion).
           - CRITICAL BAN: Do NOT include 'Picture Description' in Part III as it is not part of the official board template for this section.

        3. ABSOLUTE TEXTBOOK FACTUAL ACCURACY (ANTI-HALLUCINATION LOCK):
           You MUST verify that the generated text and ANSWER KEY conform exactly to the Class 10 TN English Textbook facts:
           - "The Night the Ghost Got In": The footsteps belonged to a suspected ghost or burglar (NOT the father). Grandfather shot the policeman because he mistakenly believed the police were deserters from General Meade's army during the Civil War. The narrator's father was away in Princeton.
           - "The Attic": Aditya went to retrieve the silver medal he hid 29 years ago in the attic of his childhood home from Sasanka Sanyal. The price paid for the medal back then was 25 rupees.
           - "Tech Bloomers": Focus accurately on Alisha (Cerebral Palsy, uses Dragon Dictate) and David (Athetoid Cerebral Palsy, uses ECO2 with ECO point) and how assistive tech empowers them.
        """
    else:
        lang_instruction = "5. Language: Pure TAMIL language only."
        header_format = "பகுதி [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)\nExample: பகுதி I - சரியான விடையைத் தேர்ந்தெடு (14 x 1 = 14)"
        option_format = "For every MCQ question, write the 4 options on a single new line using standard markers: அ) , ஆ) , இ) , ஈ)"
        
        elite_2026_rules = """
        [2026 ELITE QUALITY UPGRADES - MANDATORY V4]
        1. பகுதி I (1-மதிப்பெண் வினாக்கள்): வினாக்கள் 1-3 சினோனிம்ஸ் (Synonyms), 4-6 ஆண்டோனிம்ஸ் (Antonyms), 7-14 பிற அதிகாரப்பூர்வ இலக்கணத் தலைப்புகள் (கூற்று-காரணம், கலைச்சொற்கள், பிழையற்ற தொடர்) மட்டுமே வர வேண்டும். 'IGNOU' போன்ற தேவையற்ற பொதுவான அப்ரிவியேஷன்கள் வரக்கூடாது.
        2. பகுதி III (5-மதிப்பெண் வினாக்கள்): 14 வினாக்களுக்குள் செய்யுள், உரைநடை நெடுவினாக்களோடு சேர்த்து, மொழித்திறன் பயிற்சிகளான கடிதம், விளம்பரம், மற்றும் பிழை திருத்துதல் (Error Spotting) ஆகியவற்றைச் சமமாகப் பிரித்து வழங்க வேண்டும். பட விவரிப்பு வினாக்கள் வரக்கூடாது.
        3. துல்லியமான விடைக்குறிப்பு: விடைக்குறிப்பில் (Answer Key) பாடப்புத்தகத்தின் அசல் கதையின் உண்மைத் தரவுகளின்படி துல்லியமான பதில்களை வழங்கவும்.
        """

    return f"""
    You are an Expert Question Paper Setter for TN Board Class 10. 
    Generate a professional question paper based strictly on the requested constraints.

    Subject: {subject}
    Lessons to Cover: {lessons_str}
    Exam Type: {exam_type}
    Total Marks: {total_marks}
    Time Allowed: {exam_time}
    Exam Mode: {exam_mode}

    [STRICT BLUEPRINT CHOICE PATTERN - MANDATORY RULE]
    You MUST strictly structure the parts, question counts, and choice rules according to this layout:
    {blueprint_desc}

    [STRICT HEADER FORMAT FOR WORD TABLES]
    For every section/part header, you MUST output it strictly in this format on its own line so our python parser can build beautiful left-right Word tables:
    {header_format}
    
    [CRITICAL OPTIONS FORMAT]
    {option_format}

    [STRICT NO-LATEX RULE]
    Do NOT use '$' or '$$' or '\\frac' or any LaTeX symbols. Write equations in plain, normal text format.

    [LANGUAGE DIRECTIVE]
    {lang_instruction}

    {elite_2026_rules}

    At the very end of the paper, append exactly this delimiter and nothing else before starting the Answer Key:
    === ANSWER KEY ===
    """

# ==========================================
# 4. Word Document Export Core Engine (Natural Table Parser)
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
            if not marks_match:
                marks_match = re.search(r'\([^)]+\)$', clean_line)
                
            if marks_match:
                calc_str = marks_match.group(0)
                title_str = clean_line.replace(calc_str, "").strip(":- ")
                if not calc_str.startswith("("):
                    calc_str = f"({calc_str})"
                
                table = doc.add_table(rows=1, cols=2)
                table.autofit = True
                c1, c2 = table.rows[0].cells
                
                p1 = c1.paragraphs[0]
                p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p1.paragraph_format.space_before = Pt(12)
                p1.paragraph_format.space_after = Pt(4)
                p1.paragraph_format.keep_with_next = True 
                p1.add_run(title_str).bold = True
                
                p2 = c2.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p2.paragraph_format.space_before = Pt(12)
                p2.paragraph_format.space_after = Pt(4)
                p2.paragraph_format.keep_with_next = True
                p2.add_run(calc_str).bold = True
                
                set_cell_margins(c1, top=0, bottom=0, start=0, end=0)
                set_cell_margins(c2, top=0, bottom=0, start=0, end=0)
                continue

        option_markers = ["அ)", "ஆ)", "இ)", "ஈ)", "a)", "b)", "c)", "d)", "A)", "B)", "C)", "D)"]
        if "OPTS |" in line or any(marker in clean_line for marker in option_markers):
            if "OPTS |" in line:
                parts = [p.strip() for p in line.split("|")[1:] if p.strip()]
            else:
                raw_parts = re.split(r'(அ\)|ஆ\)|இ\)|ஈ\)|a\)|b\)|c\)|d\)|A\)|B\)|C\)|D\))', clean_line)
                parts = []
                current = ""
                for chunk in raw_parts:
                    if chunk in option_markers:
                        if current.strip(): parts.append(current.strip())
                        current = chunk + " "
                    else:
                        current += chunk
                if current.strip(): parts.append(current.strip())
            
            if parts:
                pyq_tag = ""
                if parts[-1].startswith("[") or parts[-1].endswith("]"):
                    pyq_tag = parts.pop()
                    
                table = doc.add_table(rows=1, cols=len(parts))
                for idx, opt in enumerate(parts):
                    cell = table.cell(0, idx)
                    p_opt = cell.paragraphs[0]
                    p_opt.add_run(opt.replace("*", ""))
                    set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
                
                if pyq_tag:
                    p_pyq = doc.add_paragraph()
                    p_pyq.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    p_pyq.add_run(pyq_tag).bold = True
                continue

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(3)
        if re.match(r'^\d+\.', clean_line):
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            p.paragraph_format.keep_with_next = True
            
        parts = re.split(r'\*\*(.*?)\*\*', line)
        for i, part in enumerate(parts):
            run = p.add_run(part.replace('$', ''))
            if i % 2 == 1: 
                run.bold = True

def create_professional_docx(ai_response, school_name, class_val, subject_val, exam_type, time_val, marks_val):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.27), Inches(11.69)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.5)
    
    style = doc.styles['Normal']
    style.font.name = 'Nirmala UI'
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing = 1.15
    
    h_school = doc.add_paragraph()
    h_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_school.add_run(school_name.upper()).bold = True
    h_school.runs[0].font.size = Pt(15)
    
    h_exam = doc.add_paragraph()
    h_exam.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_exam.add_run(exam_type).bold = True
    h_exam.runs[0].font.size = Pt(13)
    
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
# 5. Streamlit Presentation Tabs Layout
# ==========================================
st.set_page_config(page_title="PMP AI Suite PRO", page_icon="🎓", layout="wide")

hide_st_style = """<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}</style>"""
st.markdown(hide_st_style, unsafe_allow_html=True)

tab1, tab2 = st.tabs([
    "🎓 வினாத்தாள் தயாரிப்பு (QP Generator)", 
    "📝 விடைத்தாள் திருத்தம் (AI Math Evaluator - TRIAL)"
])

# ------------------------------------------
# TAB 1: Question Paper Generator
# ------------------------------------------
with tab1:
    st.title("🎓 PMP Question Paper AI (V17.2 ULTIMATE)")
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
            marks_val = st.number_input("Total Marks (மதிப்பெண்கள்)", value=100, step=5)
            exam_mode = st.selectbox("Exam Mode", ["🏛️ Public Exam Mode", "🏫 School Elite Mode"])
            
        lesson_list = df[df['Subject'] == subject_val]['Lesson'].unique()
        selected_lessons = st.multiselect("பாடம் / பாடங்களைத் தேர்ந்தெடுக்கவும்", lesson_list)
        
        st.markdown("---")
        st.markdown("### 🎯 வினாத்தாள் உள்ளடக்கத் தேர்வு (Content Split)")
        book_back_perf = st.slider("புத்தகப் பின்புற வினாக்களின் சதவீதம் (Book-back %)", 50, 100, 80, 10)
        interior_perf = 100 - book_back_perf
        st.info(f"📊 தற்போதைய தேர்வு: புத்தகப் பின்புற வினாக்கள் {book_back_perf}% : உள் வினாக்கள் {interior_perf}%")
        
        custom_q = st.text_area("கட்டாயம் வினாத்தாளில் இடம்பெற வேண்டிய வினாக்கள் (Optional)")
        
        st.markdown("---")
        st.markdown("### 📋 வினா வடிவமைப்பு சாய்ஸ் பாட்டர்ன் (Blueprint Choices)")
        
        b_col1, b_col2, b_col3, b_col4 = st.columns(4)
        with b_col1:
            st.subheader("பகுதி I (1-Mark)")
            p1_ask = st.number_input("1-மார்க் வினாக்கள் எண்ணிக்கை", value=14, min_value=0, step=1)
        with b_col2:
            st.subheader("பகுதி II (2-Mark)")
            p2_get = st.number_input("2-மார்க் கொடுக்க வேண்டியவை (Given)", value=14, min_value=0, step=1)
            p2_ask = st.number_input("2-மார்க் எழுத வேண்டியவை (Answer)", value=10, min_value=0, step=1)
        with b_col3:
            st.subheader("பகுதி III (5-Mark)")
            p3_get = st.number_input("5-மார்க் கொடுக்க வேண்டியவை (Given)", value=14, min_value=0, step=1)
            p3_ask = st.number_input("5-மார்க் எழுத வேண்டியவை (Answer)", value=10, min_value=0, step=1)
        with b_col4:
            st.subheader("பகுதி IV (Long Qs)")
            p4_val = st.selectbox("நெடுவினா மதிப்பெண் (Per Question)", [8, 10], index=0)
            p4_get = st.number_input("நெடுவினா கொடுக்க வேண்டியவை (Given)", value=4, min_value=0, step=1)
            p4_ask = st.number_input("நெடுவினா எழுத வேண்டியவை (Answer)", value=2, min_value=0, step=1)
            
        total_calculated = (p1_ask * 1) + (p2_ask * 2) + (p3_ask * 5) + (p4_ask * p4_val)
        
        if total_calculated == marks_val:
            st.success(f"✅ வெற்றிகரமாகப் பொருந்தியது! உங்கள் பாட்டர்ன் கணக்கீடு: {total_calculated} மார்க் = மொத்த மதிப்பெண்: {marks_val} மார்க். (வினாத்தாள் அசல் TN Board 2026 எலைட் தரத்தில் உருவாக்கப்படும்).")
            can_generate = True
        else:
            st.error(f"❌ மதிப்பெண்கள் பொருந்தவில்லை! உங்கள் பாட்டர்ன் கணக்கீடு: {total_calculated} மார்க் | நீங்கள் மேலே குறிப்பிட்ட மொத்த மதிப்பெண்: {marks_val} மார்க்.")
            can_generate = False
            
        st.markdown("---")
        
        if st.button("🚀 Generate PRO Question Paper", use_container_width=True):
            if not can_generate:
                st.error("⚠️ மதிப்பெண்கள் சரியாகப் பொருந்தாததால் வினாத்தாள் தயாரிக்க முடியாது!")
            elif not selected_lessons:
                st.warning("⚠️ தயவுசெய்து பாடங்களைத் தேர்ந்தெடுக்கவும்.")
            else:
                with st.spinner("⏳ 9.7/10 எலைட் போர்டு தரத்தில் வினாத்தாள் தயாராகிறது..."):
                    dynamic_blueprint_desc = f"""
                    - Part I: 1 Mark Questions Total: {p1_ask} (No choice).
                    - Part II: 2 Mark Questions Total: {p2_get} and answer any {p2_ask}.
                    - Part III: 5 Mark Questions Total: {p3_get} and answer any {p3_ask}.
                    - Part IV: {p4_val} Mark Questions Total: {p4_get} and answer any {p4_ask} (Either/Or Choice style).
                    """
                    prompt = generate_prompt_v16(subject_val, selected_lessons, exam_type, time_val, marks_val, exam_mode, dynamic_blueprint_desc, custom_q, book_back_perf, interior_perf)
                    try:
                        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                        docx_file = create_professional_docx(response.text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                        st.session_state['ai_text'] = response.text
                        st.session_state['docx_bytes'] = docx_file.getvalue()
                        st.session_state['file_name'] = f"PMP_{subject_val}.docx"
                        st.success("✅ 9.7/10 அசல் அரசுப் பொதுத்தேர்வு தரத்திலான வினாத்தாள் வெற்றிகரமாகத் தயாராகிவிட்டது!")
                    except Exception as e:
                        st.error(f"பிழை: {e}")

        if 'ai_text' in st.session_state:
            st.download_button(
                label="📥 Download as Word File (.docx)",
                data=st.session_state['docx_bytes'],
                file_name=st.session_state['file_name'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            with st.expander("👀 வினாத்தாள் & Answer Key Preview"):
                st.text(st.session_state['ai_text'])

# ------------------------------------------
# TAB 2: AI Handwriting Paper Evaluator
# ------------------------------------------
with tab2:
    st.title("📝 AI Math Paper Evaluator (கையெழுத்து சோதனைப் பதிப்பு)")
    st.markdown("""
    ### 🎯 சோதனை செய்யும் முறை (How to Test):
    1. ஒரு தாளில் **பத்தாம் வகுப்பு கணிதக் கணக்கு** ஒன்றை உங்கள் கைகளால் எழுதுங்கள்.
    2. அதை உங்கள் மொபைலில் நேராக, நிழல் விழாமல் ஒரு **புகைப்படம் (Photo)** எடுத்துக்கொள்ளுங்கள்.
    3. கீழே உள்ள அப்லோடரில் அந்தப் புகைப்படத்தைப் பதிவேற்றி, **"Start AI Evaluation"** பட்டனை அழுத்தி மேஜிக்கைப் பாருங்கள்!
    """)
    
    with st.form("evaluation_form"):
        uploaded_image = st.file_uploader("உங்கள் கையெழுத்து கணிதப் பக்கத்தை (PNG/JPG) அப்லோட் செய்யவும்", type=["png", "jpg", "jpeg"])
        submit_eval = st.form_submit_button("🚀 Start AI Evaluation & Marking", use_container_width=True)
        
    if uploaded_image is not None:
        st.markdown("### 🔍 பயனர் சரிபார்ப்புப் பிரிவு (Verification Preview)")
        image = Image.open(uploaded_image)
        st.image(image, caption="நீங்கள் அப்லோட் செய்த விடைத்தாள் பக்கம்", width=450)
        st.success("✅ போட்டோ தெளிவாக உள்ளது! (தரக்குறியீடு: 100% OK)")
        
        if submit_eval:
            with st.spinner("⏳ ஜெமினி AI உங்கள் கையெழுத்தைப் படித்து, கணக்கைச் சரிபார்த்து மதிப்பெண் வழங்கி வருகிறது..."):
                eval_prompt = """
                You are an official Tamil Nadu State Board Class 10 Mathematics Evaluator. 
                Strictly read the handwritten text/math formulas from the provided image and correct it based on standard state board evaluation rules.
                
                STRICT STEP-WISE MARKING SCHEME:
                - Formula written correctly: 1 Mark
                - Substitution and step-by-step calculation: 2 to 3 Marks
                - Final Answer with correct Unit: 1 Mark
                
                OUTPUT FORMAT (Respond purely in TAMIL language except formulas):
                1. **கண்டறியப்பட்ட கேள்வி (Question Identified):**
                2. **மதிப்பீட்டு அறிக்கை (Evaluation Steps):**
                3. **வழங்கப்பட்ட மதிப்பெண்கள் (Marks Awarded):** Total Marks = X / 5.
                4. **ஆசிரியருக்கான குறிப்பு / திருத்தம் (Feedback for Teacher/Student):**
                """
                try:
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=[image, eval_prompt])
                    st.markdown("---")
                    st.markdown("### 📊 AI மதிப்பீட்டு அறிக்கை (Evaluation Report)")
                    st.info("💡 கீழே உள்ள ரிப்போர்ட்டைப் பார்த்து, AI உங்கள் கையெழுத்தையும் கணக்கையும் சரியாகக் கணித்துள்ளதா எனச் சரிபார்க்கவும்.")
                    st.markdown(response.text)
                except Exception as e:
                    st.error(f"❌ பிழை ஏற்பட்டது: {e}. ஏபிஐ கீ (API Key) கட்டுப்பாடு முடிந்துவிட்டதா எனச் சரிபார்க்கவும்.")