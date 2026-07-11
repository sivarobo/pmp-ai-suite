import streamlit as st
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
import datetime
import random
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import psycopg2
import psycopg2.extras
from streamlit_oauth import OAuth2Component
import requests

# ==========================================
# st.set_page_config - MUST BE FIRST
# ==========================================
st.set_page_config(page_title="PMP AI Suite PRO", page_icon="🎓", layout="centered")

# ==========================================
# CSS Styling
# ==========================================
st.markdown("""
<style>
    html, body, [data-testid="stMarkdownContainer"] p {
        font-size: 18px !important;
        font-weight: 500 !important;
    }
    .stSelectbox label, .stTextInput label, .stNumberInput label {
        font-size: 19px !important;
        font-weight: bold !important;
        color: #1E3A8A !important;
    }
    button { font-size: 18px !important; height: 50px !important; }

    /* Google Login Card */
    .google-login-wrapper {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 20px 0;
    }
    .google-btn-link {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 12px;
        background: #ffffff;
        color: #3c4043;
        border: 1.5px solid #dadce0;
        border-radius: 8px;
        padding: 14px 32px;
        font-size: 16px;
        font-weight: 600;
        text-decoration: none !important;
        width: 100%;
        box-shadow: 0 2px 8px rgba(0,0,0,0.10);
        transition: box-shadow 0.2s, background 0.2s;
        font-family: 'Segoe UI', sans-serif;
        cursor: pointer;
    }
    .google-btn-link:hover {
        box-shadow: 0 4px 16px rgba(0,0,0,0.16);
        background: #f8faff;
        text-decoration: none !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# DB Connection (Neon PostgreSQL)
# ==========================================
def get_db():
    try:
        conn = psycopg2.connect(
            st.secrets["NEON_DATABASE_URL"],
            cursor_factory=psycopg2.extras.RealDictCursor
        )
        return conn
    except Exception as e:
        st.error(f"DB Connection Error: {e}")
        return None

# ==========================================
# Google OAuth — streamlit-oauth library
# ==========================================
GOOGLE_CLIENT_ID     = st.secrets["google"]["client_id"]
GOOGLE_CLIENT_SECRET = st.secrets["google"]["client_secret"]
GOOGLE_REDIRECT_URI  = st.secrets["google"]["redirect_uri"]

oauth2 = OAuth2Component(
    client_id=GOOGLE_CLIENT_ID,
    client_secret=GOOGLE_CLIENT_SECRET,
    authorize_endpoint="https://accounts.google.com/o/oauth2/v2/auth",
    token_endpoint="https://oauth2.googleapis.com/token",
    refresh_token_endpoint="https://oauth2.googleapis.com/token",
    revoke_token_endpoint="https://oauth2.googleapis.com/revoke",
)

def _google_userinfo(access_token):
    r = requests.get(
        "https://www.googleapis.com/oauth2/v2/userinfo",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=10
    )
    return r.json() if r.status_code == 200 else None

# ==========================================
# User DB Functions (Google-based)
# ==========================================
def upsert_google_user(google_info):
    """
    Google login → DB-ல் user save/update.
    streamlit-google-auth provides: name, email, picture
    """
    try:
        conn = get_db()
        if not conn:
            return None
        cur = conn.cursor()

        # Table create (first run only - safe to run every time)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id          SERIAL PRIMARY KEY,
                google_id   TEXT UNIQUE,
                email       TEXT UNIQUE,
                name        TEXT,
                picture     TEXT,
                plan        TEXT DEFAULT 'free',
                daily_count INTEGER DEFAULT 0,
                last_used   DATE DEFAULT CURRENT_DATE,
                created_at  TIMESTAMP DEFAULT NOW(),
                last_login  TIMESTAMP DEFAULT NOW()
            )
        """)

        # Upsert by google_id
        cur.execute("""
            INSERT INTO users (google_id, email, name, picture, last_login)
            VALUES (%s, %s, %s, %s, NOW())
            ON CONFLICT (email) DO UPDATE
              SET name       = EXCLUDED.name,
                  picture    = EXCLUDED.picture,
                  last_login = NOW()
            RETURNING id, google_id, email, name, picture, plan, created_at
        """, (
            google_info.get("sub", google_info.get("id", "")),
            google_info["email"],
            google_info["name"],
            google_info.get("picture", ""),
        ))

        row = cur.fetchone()
        conn.commit()
        conn.close()
        return dict(row) if row else None

    except Exception as e:
        st.error(f"DB Error: {e}")
        return None

def get_today_usage(user_id):
    try:
        conn = get_db()
        if not conn:
            return 0
        cur = conn.cursor()
        cur.execute(
            "SELECT question_count FROM daily_usage WHERE user_id = %s AND usage_date = CURRENT_DATE",
            (str(user_id),)
        )
        row = cur.fetchone()
        conn.close()
        return row["question_count"] if row else 0
    except:
        return 0

def increment_usage(user_id):
    try:
        conn = get_db()
        if not conn:
            return False
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO daily_usage (user_id, usage_date, question_count)
               VALUES (%s, CURRENT_DATE, 1)
               ON CONFLICT (user_id, usage_date)
               DO UPDATE SET question_count = daily_usage.question_count + 1""",
            (str(user_id),)
        )
        conn.commit()
        conn.close()
        return True
    except:
        return False

# ==========================================
# QUESTION BANK (புத்தக கேள்வி வங்கி) — DB Functions
# ==========================================
def ensure_question_bank_table():
    try:
        conn = get_db()
        if not conn:
            return
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS question_bank (
                id            SERIAL PRIMARY KEY,
                subject       TEXT,
                lesson        TEXT,
                mark_type     TEXT,
                qtype         TEXT,
                question_text TEXT,
                answer_text   TEXT,
                created_at    TIMESTAMP DEFAULT NOW()
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Question Bank Table Error: {e}")

def fetch_bank_questions(subject, lesson, mark_type):
    try:
        conn = get_db()
        if not conn:
            return []
        cur = conn.cursor()
        cur.execute(
            "SELECT id, qtype, question_text, answer_text FROM question_bank "
            "WHERE subject=%s AND lesson=%s AND mark_type=%s ORDER BY id",
            (subject, lesson, mark_type)
        )
        rows = cur.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    except Exception as e:
        st.error(f"Bank Fetch Error: {e}")
        return []

MARK_TYPE_LABELS = {
    "1M":   "1-மார்க் (MCQ / குறுவினா)",
    "2M":   "2-மார்க் வினா",
    "5M":   "5-மார்க் வினா",
    "LONG": "நெடுவினா (Long Answer)",
}

def save_bank_questions(subject, lesson, mark_type, items):
    try:
        conn = get_db()
        if not conn:
            return
        cur = conn.cursor()
        for it in items:
            cur.execute(
                "INSERT INTO question_bank (subject, lesson, mark_type, qtype, question_text, answer_text) "
                "VALUES (%s, %s, %s, %s, %s, %s)",
                (subject, lesson, mark_type, it.get("qtype", "பயிற்சி"), it.get("question", "").strip(), it.get("answer", "").strip())
            )
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Bank Save Error: {e}")

def generate_bank_questions_ai(subject, lesson, mark_type, count=8):
    """Gemini மூலம் புத்தக பாணி (Book Back / Exercise / Example) கேள்விகளை உருவாக்குதல்"""
    mark_desc = MARK_TYPE_LABELS.get(mark_type, mark_type)
    prompt = f"""நீங்கள் Tamil Nadu Class 10 பாடநூல் நிபுணர். பாடம்: "{lesson}" | பாடப்பிரிவு: {subject}.

இந்த பாடத்திலிருந்து {count} கேள்விகளை உருவாக்கவும் — மதிப்பெண் வகை: {mark_desc}.
இவை TN Samacheer Kalvi பாடப்புத்தகத்தின் "பின்புற வினாக்கள்" (Book Back Exercise), "பயிற்சி கணக்குகள்" (Practice Problems),
மற்றும் "எடுத்துக்காட்டு கணக்குகள்" (Solved Examples) பாணியை நெருக்கமாக ஒத்திருக்க வேண்டும்.
qtype-ஐ மூன்றிலும் கலந்து (mix) கொடுக்கவும்.

பதில் STRICTLY இந்த JSON அணி (array) வடிவில் மட்டும் இருக்க வேண்டும், வேறு எந்த உரையும் (preamble, code fences) கூடாது:
[
  {{"qtype": "பின்புற வினா", "question": "தமிழில் முழு கேள்வி", "answer": "தமிழில் சுருக்கமான விடை/தீர்வு"}},
  ...
]
கணிதம் சின்னங்களுக்கு LaTeX பயன்படுத்தாதீர்கள் — plain Unicode (×, ÷, √, ², π, ∠ போன்றவை) மட்டும் பயன்படுத்தவும்.
"""
    try:
        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
        raw = response.text.strip()
        raw = re.sub(r'^```json\s*|\s*```$', '', raw.strip(), flags=re.MULTILINE).strip()
        raw = re.sub(r'^```|```$', '', raw.strip()).strip()
        items = json.loads(raw)
        if isinstance(items, list):
            return items
        return []
    except Exception as e:
        st.warning(f"⚠️ '{lesson}' ({mark_desc}) கேள்வி வங்கி உருவாக்கத்தில் பிழை: {e}")
        return []

def get_or_build_bank(subject, lesson, mark_type, min_count=8):
    """DB-ல் இருந்தால் fetch பண்ணும், இல்லைனா AI மூலம் generate பண்ணி DB-ல் save பண்ணி திருப்பும்"""
    existing = fetch_bank_questions(subject, lesson, mark_type)
    if len(existing) >= min_count:
        return existing
    needed = max(min_count - len(existing), 6)
    new_items = generate_bank_questions_ai(subject, lesson, mark_type, count=needed)
    if new_items:
        save_bank_questions(subject, lesson, mark_type, new_items)
        existing = fetch_bank_questions(subject, lesson, mark_type)
    return existing

def assemble_paper_from_bank(parts_cfg):
    """
    parts_cfg: list of dicts —
      {"label": "பகுதி I", "mark": 1, "given": N, "answer": N, "note": "", "items": [ {question_text, answer_text}, ... ]}
    தேர்ந்தெடுக்கப்பட்ட கேள்விகளை மட்டும் வைத்து AI-response போன்ற text உருவாக்குகிறது
    (create_professional_docx() அதே pipeline-ஐ மறுபயன்பாடு செய்ய).
    """
    q_lines = []
    a_lines = []
    counter = 1
    for part in parts_cfg:
        items = part["items"]
        if not items:
            continue
        total_marks = part["given"] * part["mark"]
        header = f'{part["label"]} ( {part["given"]} x {part["mark"]} = {total_marks} )'
        if part.get("note"):
            header += f'  [{part["note"]}]'
        q_lines.append(header)
        a_lines.append(part["label"])
        for it in items:
            q_lines.append(f'{counter}. {it["question_text"]}')
            a_lines.append(f'{counter}. {it.get("answer_text", "")}')
            counter += 1
        q_lines.append("")
        a_lines.append("")
    return "\n".join(q_lines) + "\n=== ANSWER KEY ===\n" + "\n".join(a_lines)

FREE_DAILY_LIMIT = 2

# ==========================================
# GOOGLE OAUTH CALLBACK HANDLER
# ==========================================
# ==========================================
# ACCESS GATE — streamlit-oauth
# ==========================================
if not st.session_state.get("logged_in_user"):
    st.markdown("""
    <div style='text-align:center; padding:30px 0 10px 0;'>
        <div style='font-size:56px;'>🎓</div>
        <h1 style='color:#1E3A8A; margin:8px 0 4px 0;'>PMP Master AI Suite</h1>
        <p style='color:#64748b; font-size:16px;'>AI-Powered Question Paper Generator</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style='background:#f0fdf4; border:1px solid #86efac; border-radius:10px;
                    padding:14px 18px; margin:16px 0;'>
            <b style='color:#166534;'>🎁 Free Plan:</b>
            <ul style='color:#166534; margin:6px 0 0 0; padding-left:20px;'>
                <li>தினமும் 2 Question Papers Free</li>
                <li>Always Free — No Expiry</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        result = oauth2.authorize_button(
            name="Google-ல் உள்நுழைக",
            icon="https://www.google.com/favicon.ico",
            redirect_uri=GOOGLE_REDIRECT_URI,
            scope="openid email profile",
            key="google_login",
            extras_params={"prompt": "select_account"},
            use_container_width=True,
        )

        if result and "token" in result:
            token = result["token"]
            access_token = token.get("access_token", "")
            guser = _google_userinfo(access_token)
            if guser and "email" in guser:
                db_user = upsert_google_user(guser)
                st.session_state["logged_in_user"] = db_user or {
                    "id":      0,
                    "email":   guser.get("email", ""),
                    "name":    guser.get("name", "User"),
                    "picture": guser.get("picture", ""),
                    "plan":    "free",
                }
                st.rerun()

        st.markdown("""
        <p style='text-align:center;color:#94a3b8;font-size:13px;margin-top:8px;'>
            OTP தேவையில்லை · Gmail account தேர்ந்தெடுக்கவும்
        </p>
        """, unsafe_allow_html=True)
    st.stop()

# ==========================================
# LOGGED IN — Get user & check usage
# ==========================================
current_user = st.session_state["logged_in_user"]
user_id   = current_user["id"]
user_name = current_user["name"]
user_email= current_user["email"]
user_plan = current_user.get("plan", "free")
user_pic  = current_user.get("picture", "")

today_usage  = get_today_usage(user_id)
is_premium   = user_plan in ["premium", "paid"]
limit_reached= (not is_premium) and (today_usage >= FREE_DAILY_LIMIT)

if limit_reached:
    show_limit_reached_page(current_user, today_usage)
    st.stop()

# ==========================================
# USAGE BANNER
# ==========================================
remaining = FREE_DAILY_LIMIT - today_usage
if not is_premium:
    if remaining <= 1:
        st.warning(f"⚠️ இன்று {remaining} question paper மட்டுமே மிச்சம் | 👤 {user_name}")
    else:
        st.info(f"🎁 இன்று {remaining}/{FREE_DAILY_LIMIT} free papers மிச்சம் | 👤 {user_name}")
else:
    st.success(f"⭐ Premium Plan | 👤 {user_name} | Unlimited Access")

# ==========================================
# API Configuration
# ==========================================
YOUR_API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=YOUR_API_KEY)

# ==========================================
# Database Loading
# ==========================================
@st.cache_data
def load_data():
    try:
        return pd.read_csv('lesson_master_v1_5.csv')
    except:
        return pd.DataFrame()

ensure_question_bank_table()

def get_math_dynamic_weightage(selected_lessons, part1_val, part2_val, part3_val):
    base_matrix = {
        "Relations and Functions":      {"1M": 1.5, "2M": 2, "5M": 1.5, "8M": 0},
        "Numbers and Sequences":        {"1M": 2.0, "2M": 2, "5M": 2.0, "8M": 0},
        "Algebra":                      {"1M": 2.0, "2M": 2, "5M": 2.0, "8M": 1},
        "Geometry":                     {"1M": 2.0, "2M": 1, "5M": 1.0, "8M": 1},
        "Coordinate Geometry":          {"1M": 1.5, "2M": 2, "5M": 2.0, "8M": 0},
        "Mensuration":                  {"1M": 1.5, "2M": 2, "5M": 2.0, "8M": 0},
        "Statistics and Probability":   {"1M": 2.0, "2M": 2, "5M": 2.0, "8M": 0}
    }
    rules = []
    for lesson in selected_lessons:
        if lesson in base_matrix:
            bm = base_matrix[lesson]
            rules.append(f"- From '{lesson}': Generate approx {int(bm['1M'])} MCQs, {bm['2M']} Questions (2-Mark), {int(bm['5M'])} Questions (5-Mark), and {bm['8M']} Question (8-Mark).")
    return "\n".join(rules)

def get_english_blueprint_rules():
    return """
    [STRICT MASTER ENGLISH BLUEPRINT LOCK]
    PART I (14 Marks): One Mark Questions
    - Q1-3: Synonyms strictly selected from the text prose context.
    - Q4-6: Antonyms strictly selected from the text prose context.
    - Q7-14: Textual Grammar Matrix: Plural Form, Suffix/Prefix, Abbreviations, Phrasal Verbs, Compound Words, Prepositions, Tense Forms, Linkers.
    PART II (20 Marks): Two Mark Questions (Answer any 10 out of 12)
    - Section 1 (Prose Qs): 3 Short answer questions from selected textbook lessons.
    - Section 2 (Poetry Appreciation): 3 Poetic line sets with internal appreciation questions.
    - Section 3 (Grammar Blocks): 5 Mandatory Core Grammar Questions: Voice Change, Reported Speech, Punctuation, Sentence Transformation, Rearrange jumbled words.
    - Section 4 (Roadmap): 1 Question on Road Map directions guide.
    PART III (35 Marks): Five Mark Questions (Answer any 7 out of 10)
    - Section 1 (Literature Paragraphs): Prose Paragraphs and Poetry Paragraphs covering textbook lessons.
    - Section 2 (Coherent Order & Comprehension): Rearrange sentences in coherent order, Supplementary passage comprehension.
    - Section 3 (Writing Skills Matrix): Advertisement, Formal Letter, Notice Writing, Picture Comprehension.
    PART IV (16 Marks): Eight Mark Questions (Answer both - Internal Choice)
    - Q37: Comprehensive Prose Passage reading or Poem reading with granularity questions.
    - Q38: Detailed Literature Essay from supplementary stories.
    """

def get_blueprint_defaults(total_marks, is_social=False, is_english=False):
    if is_english or is_social:
        return {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 10, "p3a": 7, "p4v": 8, "p4g": 2, "p4a": 2}
    defaults = {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    if total_marks == 106:
        defaults = {"p1": 20, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    elif total_marks == 50:
        defaults = {"p1": 10, "p2g": 8, "p2a": 6, "p3g": 6, "p3a": 4, "p4v": 8, "p4g": 2, "p4a": 1}
    elif total_marks == 25:
        defaults = {"p1": 5, "p2g": 6, "p2a": 5, "p3g": 3, "p3a": 2, "p4v": 8, "p4g": 0, "p4a": 0}
    return defaults

def _clean_diagram_label(label_text):
    """matplotlib Tamil render பண்ணாது - ASCII மட்டும் வைக்கிறோம்"""
    import unicodedata
    cleaned = label_text.replace("Angle", "∠").replace("angle", "∠")
    # Keep only ASCII + common math symbols (Tamil letters become boxes in matplotlib)
    safe = ''.join(c for c in cleaned if ord(c) < 0x0B80 or ord(c) > 0x0BFF)
    return safe.strip()

def generate_geometry_image(shape_type, label_text=""):
    fig, ax = plt.subplots(figsize=(2.8, 2.8))
    shape_upper = shape_type.upper()
    clean_label = _clean_diagram_label(label_text)

    if "THALES" in shape_upper or "BPT" in shape_upper:
        # தேல்ஸ் தேற்றம் / Basic Proportionality Theorem
        # Triangle ABC with D on AB, E on AC, DE parallel to BC
        A = np.array([2.0, 3.6]); B = np.array([0.2, 0.2]); C = np.array([3.8, 0.2])
        # D and E at 45% down from A
        t = 0.45
        D = A + t * (B - A)
        E = A + t * (C - A)
        # Main triangle
        tri = np.array([A, B, C, A])
        ax.plot(tri[:, 0], tri[:, 1], 'k-', lw=2)
        # DE parallel line
        ax.plot([D[0], E[0]], [D[1], E[1]], 'k-', lw=2)
        # Parallel marks on DE and BC (small double ticks)
        for seg_p1, seg_p2 in [(D, E), (B, C)]:
            mid = (seg_p1 + seg_p2) / 2
            ax.annotate('▸', xy=mid, fontsize=8, ha='center', va='center')
        # Labels
        ax.text(A[0], A[1]+0.15, 'A', fontsize=12, fontweight='bold', ha='center')
        ax.text(B[0]-0.2, B[1]-0.1, 'B', fontsize=12, fontweight='bold')
        ax.text(C[0]+0.1, C[1]-0.1, 'C', fontsize=12, fontweight='bold')
        ax.text(D[0]-0.25, D[1], 'D', fontsize=12, fontweight='bold')
        ax.text(E[0]+0.12, E[1], 'E', fontsize=12, fontweight='bold')
        # DE ∥ BC annotation
        ax.text(2.0, -0.35, 'DE ∥ BC', fontsize=11, ha='center', fontweight='bold')
        ax.set_xlim(-0.5, 4.4); ax.set_ylim(-0.7, 4.1)

    elif "EXT_BISECTOR" in shape_upper or "EXTERNAL_BISECTOR" in shape_upper:
        # வெளிப்புற கோண இருசமவெட்டி - AD meets BC EXTENSION at D
        A = np.array([1.4, 3.0]); B = np.array([0.3, 0.5]); C = np.array([2.9, 0.5])
        D = np.array([4.6, 0.5])  # on BC extension
        tri = np.array([A, B, C, A])
        ax.plot(tri[:, 0], tri[:, 1], 'k-', lw=2)
        # BC extension (dashed)
        ax.plot([C[0], D[0]], [C[1], D[1]], 'k--', lw=1.5)
        # AD line
        ax.plot([A[0], D[0]], [A[1], D[1]], 'k-', lw=1.8)
        ax.plot(*D, 'ko', markersize=4)
        ax.text(A[0], A[1]+0.15, 'A', fontsize=12, fontweight='bold', ha='center')
        ax.text(B[0]-0.25, B[1]-0.15, 'B', fontsize=12, fontweight='bold')
        ax.text(C[0]-0.1, C[1]-0.35, 'C', fontsize=12, fontweight='bold')
        ax.text(D[0]+0.08, D[1]-0.15, 'D', fontsize=12, fontweight='bold')
        ax.set_xlim(-0.3, 5.3); ax.set_ylim(-0.4, 3.6)

    elif "ANGLE_BISECTOR" in shape_upper or "BISECTOR" in shape_upper:
        # உட்புற கோண இருசமவெட்டி - AD bisects angle A, D on BC
        A = np.array([2.0, 3.4]); B = np.array([0.3, 0.4]); C = np.array([3.9, 0.4])
        D = np.array([2.35, 0.4])  # on BC (bisector foot, slightly right of midpoint)
        tri = np.array([A, B, C, A])
        ax.plot(tri[:, 0], tri[:, 1], 'k-', lw=2)
        ax.plot([A[0], D[0]], [A[1], D[1]], 'k-', lw=1.8)
        ax.plot(*D, 'ko', markersize=4)
        # angle bisector arcs at A (two small equal-angle marks)
        ax.annotate('', xy=(1.75, 2.85), xytext=(1.9, 3.0),
                    arrowprops=dict(arrowstyle='-', lw=1))
        ax.annotate('', xy=(2.25, 2.85), xytext=(2.1, 3.0),
                    arrowprops=dict(arrowstyle='-', lw=1))
        ax.text(A[0], A[1]+0.15, 'A', fontsize=12, fontweight='bold', ha='center')
        ax.text(B[0]-0.25, B[1]-0.15, 'B', fontsize=12, fontweight='bold')
        ax.text(C[0]+0.1, C[1]-0.15, 'C', fontsize=12, fontweight='bold')
        ax.text(D[0]+0.05, D[1]-0.35, 'D', fontsize=12, fontweight='bold')
        ax.set_xlim(-0.3, 4.5); ax.set_ylim(-0.5, 4.0)

    elif "TWO_TANGENT" in shape_upper or "TANGENTS" in shape_upper:
        # வெளிப்புள்ளியில் இருந்து இரண்டு தொடுகோடுகள் PA, PB
        O = np.array([1.5, 2.0]); r = 1.1
        P = np.array([4.3, 2.0])
        circle = plt.Circle(O, r, fill=False, color='black', lw=2)
        ax.add_patch(circle)
        ax.plot(*O, 'ko', markersize=4)
        ax.text(O[0]-0.05, O[1]+0.15, 'O', fontsize=11, fontweight='bold')
        # Tangent points A (top) and B (bottom)
        d = np.linalg.norm(P - O)
        ang = np.arccos(r / d)
        base = np.arctan2(P[1]-O[1], P[0]-O[0])
        Apt = O + r * np.array([np.cos(base + ang), np.sin(base + ang)])
        Bpt = O + r * np.array([np.cos(base - ang), np.sin(base - ang)])
        ax.plot([P[0], Apt[0]], [P[1], Apt[1]], 'k-', lw=1.8)
        ax.plot([P[0], Bpt[0]], [P[1], Bpt[1]], 'k-', lw=1.8)
        # Radii OA, OB (dashed)
        ax.plot([O[0], Apt[0]], [O[1], Apt[1]], 'k--', lw=1)
        ax.plot([O[0], Bpt[0]], [O[1], Bpt[1]], 'k--', lw=1)
        ax.plot(*Apt, 'ko', markersize=4)
        ax.plot(*Bpt, 'ko', markersize=4)
        ax.plot(*P, 'ko', markersize=4)
        ax.text(Apt[0]-0.1, Apt[1]+0.18, 'A', fontsize=12, fontweight='bold')
        ax.text(Bpt[0]-0.1, Bpt[1]-0.32, 'B', fontsize=12, fontweight='bold')
        ax.text(P[0]+0.1, P[1]-0.05, 'P', fontsize=12, fontweight='bold')
        ax.set_xlim(0, 5.2); ax.set_ylim(0.2, 3.8)
        ax.set_aspect('equal')

    elif "PYTHAGORAS" in shape_upper or "RIGHT" in shape_upper:
        # செங்கோண முக்கோணம் - right angle at B
        A = np.array([0.3, 3.4]); B = np.array([0.3, 0.3]); C = np.array([3.9, 0.3])
        tri = np.array([A, B, C, A])
        ax.plot(tri[:, 0], tri[:, 1], 'k-', lw=2)
        # Right angle square at B
        ax.plot([0.3, 0.7], [0.7, 0.7], 'k-', lw=1)
        ax.plot([0.7, 0.7], [0.3, 0.7], 'k-', lw=1)
        ax.text(A[0]-0.05, A[1]+0.15, 'A', fontsize=12, fontweight='bold', ha='center')
        ax.text(B[0]-0.25, B[1]-0.15, 'B', fontsize=12, fontweight='bold')
        ax.text(C[0]+0.1, C[1]-0.15, 'C', fontsize=12, fontweight='bold')
        ax.set_xlim(-0.4, 4.4); ax.set_ylim(-0.6, 4.0)

    elif "TANGENT" in shape_upper:
        # வட்டத்துக்கு தொடுகோடு
        circle = plt.Circle((1.8, 2.0), 1.2, fill=False, color='black', lw=2)
        ax.add_patch(circle)
        ax.plot(1.8, 2.0, 'ko', markersize=4)
        ax.text(1.85, 2.1, 'O', fontsize=11, fontweight='bold')
        # Tangent point P at right side of circle
        P = np.array([3.0, 2.0])
        ax.plot(*P, 'ko', markersize=4)
        ax.text(P[0]+0.1, P[1]+0.1, 'P', fontsize=11, fontweight='bold')
        # Radius OP
        ax.plot([1.8, P[0]], [2.0, P[1]], 'k-', lw=1.2)
        # Tangent line (vertical through P)
        ax.plot([P[0], P[0]], [0.4, 3.6], 'k-', lw=2)
        # Right angle mark at P
        ax.plot([P[0]-0.22, P[0]-0.22], [2.0, 2.22], 'k-', lw=1)
        ax.plot([P[0]-0.22, P[0]], [2.22, 2.22], 'k-', lw=1)
        ax.set_xlim(0, 4.2); ax.set_ylim(0, 4.0)

    elif "TRIANGLE" in shape_upper:
        points = np.array([[0.3, 0.3], [4.1, 0.3], [2.2, 3.3], [0.3, 0.3]])
        ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
        ax.text(0.1, 0.05, 'A', fontsize=12, fontweight='bold')
        ax.text(4.2, 0.05, 'B', fontsize=12, fontweight='bold')
        ax.text(2.2, 3.45, 'C', fontsize=12, fontweight='bold', ha='center')
        ax.set_xlim(-0.3, 4.7); ax.set_ylim(-0.6, 3.9)

    elif "SQUARE" in shape_upper or "RECTANGLE" in shape_upper:
        w = 4 if "RECTANGLE" in shape_upper else 3
        points = np.array([[0.3, 0.3], [w+0.3, 0.3], [w+0.3, 3.3], [0.3, 3.3], [0.3, 0.3]])
        ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
        ax.text(0.05, 0.05, 'A', fontsize=11, fontweight='bold')
        ax.text(w+0.4, 0.05, 'B', fontsize=11, fontweight='bold')
        ax.text(w+0.4, 3.4, 'C', fontsize=11, fontweight='bold')
        ax.text(0.05, 3.4, 'D', fontsize=11, fontweight='bold')
        ax.set_xlim(-0.4, w+1); ax.set_ylim(-0.6, 3.9)

    elif "SEMICIRCLE" in shape_upper or "SEMI" in shape_upper:
        theta = np.linspace(0, np.pi, 100)
        ax.plot(2 + 1.6*np.cos(theta), 1 + 1.6*np.sin(theta), 'k-', lw=2)
        ax.plot([0.4, 3.6], [1, 1], 'k-', lw=2)
        ax.plot(2, 1, 'ko', markersize=4)
        ax.text(2.05, 0.78, 'O', fontsize=11, fontweight='bold')
        ax.set_xlim(0, 4); ax.set_ylim(0, 3.2)

    elif "CIRCLE" in shape_upper:
        circle = plt.Circle((2, 2), 1.6, fill=False, color='black', lw=2)
        ax.add_patch(circle)
        ax.plot(2, 2, 'ko', markersize=4)
        ax.text(2.1, 2.1, 'O', fontsize=11, fontweight='bold')
        ax.plot([2, 3.6], [2, 2], 'k-', lw=1.2)
        ax.text(2.8, 2.14, 'r', fontsize=10, style='italic')
        ax.set_xlim(0, 4); ax.set_ylim(0, 4)

    # Label (ASCII-safe only)
    if clean_label:
        ax.text(0.5, 0.01, clean_label, fontsize=9, ha='center',
                fontweight='bold', color='blue', transform=ax.transAxes)

    ax.set_aspect('equal')
    ax.axis('off')
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=110)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

def generate_prompt_v18(subject, lessons_list, exam_type, exam_time, total_marks, exam_mode, blueprint_desc, part1_val, part2_val, part3_val, diff_level):
    lessons_str = ", ".join(lessons_list)
    sub_lower = subject.lower()
    is_english = "english" in sub_lower or "ஆங்கிலம்" in sub_lower
    is_tamil   = "tamil"   in sub_lower or "தமிழ்"    in sub_lower
    is_social  = "social"  in sub_lower or "சமூக"     in sub_lower
    is_math    = "math"    in sub_lower or "கணிதம்"   in sub_lower

    if diff_level == "எளிமை (Easy)":
        difficulty_directive = "DIFFICULTY CRITERIA: Focus 80% on direct textbook back questions and formulas."
    elif diff_level == "நடுத்தரம் (Medium)":
        difficulty_directive = "DIFFICULTY CRITERIA: Balanced public paper structure. 60% Direct, 30% Application, 10% HOTS."
    else:
        difficulty_directive = "DIFFICULTY CRITERIA: High-level revision standard with indirect grammatical transformations."

    math_weightage_directive = ""
    if is_math:
        math_weightage_directive = f"[STRICT LESSON-WISE MARKS WEIGHTAGE MATRIX]\n{get_math_dynamic_weightage(lessons_list, part1_val, part2_val, part3_val)}"

    if is_english:
        lang_instruction      = "5. Language: Pure ENGLISH only."
        header_format         = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)"
        option_format         = "Options marker: a) , b) , c) , d)"
        subject_blueprint_rules = f"[STRICT TN BOARD ENGLISH BLUEPRINT LOCK]\n{get_english_blueprint_rules()}"
    elif is_tamil:
        lang_instruction      = "5. Language: Pure TAMIL only."
        header_format         = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (வினாக்கள் எண்ணிக்கை x மதிப்பெண் = மொத்த மதிப்பெண்கள்)"
        option_format         = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = "[அசல் தமிழ் பாடத்திட்ட ப்ளூபிரின்ட்] சொல்வளம், இலக்கணம் லாக்."
    elif is_social:
        lang_instruction      = "5. Language: Pure TAMIL only."
        header_format         = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்]"
        option_format         = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = "[MANDATORY CRITICAL SOCIAL SCIENCE BLUEPRINT] Assertion-Reason, Map locked."
    elif is_math:
        lang_instruction      = "5. Language: Pure TAMIL only."
        header_format         = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (No_of_Qs x Marks = Total_Marks)"
        option_format         = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = f"""[MANDATORY CRITICAL MATHEMATICS CORE EMBEDDED LOCK]
1. ABSOLUTE BAN ON AI DISCLAIMERS.
2. GEOMETRY DIAGRAM TAGS - STRICT RULES:
   - NEVER draw diagrams using text characters (/, \\, -, |). ASCII art is ABSOLUTELY BANNED.
   - Instead, write the FULL question text first, then on the NEXT separate line add ONE diagram tag.
   - The question text must always be complete. NEVER replace question text with a tag.
   - TAG SELECTION MATRIX (match question content EXACTLY - wrong diagram = INVALID paper):
     * Question mentions "DE" and "BC" or "DE ∥ BC" or "மிகைவிகித"/"விகிதசம" → [DRAW_THALES] (NEVER plain TRIANGLE)
     * தேல்ஸ் தேற்றம் / Thales / BPT → [DRAW_THALES]
     * உட்புற கோண இருசமவெட்டி / internal angle bisector / "AD bisects ∠A" → [DRAW_ANGLE_BISECTOR]
     * வெளிப்புற கோண இருசமவெட்டி / external bisector / "D on BC extension"/"BC-ன் நீட்டிப்பு" → [DRAW_EXT_BISECTOR]
     * ஒரு தொடுகோடு / single tangent → [DRAW_TANGENT]
     * இரண்டு தொடுகோடுகள் / two tangents PA and PB from external point → [DRAW_TWO_TANGENTS]
     * பிதாகரஸ் தேற்றம் / Pythagoras / right-angle triangle → [DRAW_PYTHAGORAS]
   - CRITICAL: If question text mentions points D, E on triangle sides → NEVER use [DRAW_TRIANGLE].
     The diagram MUST show every point named in the question (A,B,C,D,E,P etc).
     A diagram missing a point mentioned in the question is a CRITICAL ERROR.
   - GENERIC shape tags: [DRAW_TRIANGLE], [DRAW_SQUARE], [DRAW_RECTANGLE], [DRAW_CIRCLE], [DRAW_SEMICIRCLE]
   - Labels inside tags must be ENGLISH/numbers only (e.g. [DRAW_CIRCLE: r=7cm]). NO Tamil text inside tags.
   - Example CORRECT:
     35. தேல்ஸ் தேற்றத்தை எழுதி நிரூபிக்கவும்.
     [DRAW_THALES]
3. GRAPH PAPER COORDINATES.
{math_weightage_directive}"""
    else:
        lang_instruction      = "5. Language: Pure TAMIL language only."
        header_format         = "பகுதி [ROMAN_NUM]"
        option_format         = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = ""

    no_latex_rule = """
[STRICT NO-LATEX OUTPUT RULE - CRITICAL]
- NEVER output LaTeX syntax. This is a printable exam paper, not a LaTeX document.
- BANNED: \\rightarrow, \\frac{}{}, \\{ \\}, \\times, \\sqrt{}, x^{2}, $...$, \\in, \\triangle etc.
- USE INSTEAD plain Unicode: → for arrows, { } for sets, × for multiply,
  √ for root, x² x³ superscripts, ∈ for element-of, ∠ for angle, △ for triangle,
  a/b format for fractions, θ π α β for Greek letters.
- Example CORRECT: f: A → B, f(x) = 3x - 1, A = {1, 2, 3, 4}, x² + y² = r²
- Example WRONG: f: A \\rightarrow B, A = \\{1, 2, 3\\}, x^{2}
"""

    theorem_proof_rule = """
[ANSWER KEY COMPLETENESS RULE - CRITICAL]
- "Refer Textbook" ([பாடநூல் பார்க்கவும்]) is allowed ONLY for:
  * Theorem PROOFS (தேற்றத்தை நிரூபிக்கவும்)
  * Definitions (வரையறு)
  * Compass/ruler CONSTRUCTION steps (வரைக/அமைக்க)
- For ALL other questions - numerical problems, equations, statistics,
  probability, coordinate geometry, mensuration, graph calculations -
  the Answer Key MUST contain COMPLETE STEP-BY-STEP working with the final answer.
  This is especially MANDATORY for compulsory questions and Part III/IV
  high-mark questions. A numerical question answered with Refer-Textbook is INVALID.
- For graph questions: Answer Key must include the table of values (x, y points)
  and the final answer read from the graph.

[ARITHMETIC ACCURACY RULE - CRITICAL]
- Before writing any Answer Key calculation, VERIFY every arithmetic step.
- For STATISTICS questions (mean, standard deviation, variance):
  * Choose data values so that the sum, sum of squares, and mean are CLEAN numbers.
  * Show working explicitly: n = ..., sum = ..., mean = ..., variance = ..., SD = ...
    with each intermediate value actually computed (not just formula with symbols).
  * The final SD should be a clean value (whole number or one decimal).
- If you cannot verify a calculation is correct, redesign the question with simpler numbers.

[OUTPUT HYGIENE RULE - ABSOLUTELY CRITICAL]
- This output is a FINAL PRINTED EXAM DOCUMENT given directly to students and teachers.
- NEVER include any meta-commentary, thinking process, or self-correction text such as:
  "Let me try a different approach", "Re-checking", "My initial analysis was correct",
  "This is complex", "Re-design the question", "நான் ... கணக்கிடுகிறேன்",
  "குறிப்பு: ... தெளிவாக குறிப்பிடப்படவில்லை" or ANY similar draft/verification notes.
- Output ONLY the polished final content. Any visible reasoning = INVALID output.

[QUESTION-ANSWER CONSISTENCY RULE - ABSOLUTELY CRITICAL]
- The question printed in the paper and the question solved in the Answer Key
  MUST be 100% IDENTICAL. NEVER solve a modified version in the Answer Key.
- DESIGN QUESTIONS BACKWARDS: first choose the clean final answer, then build
  the question from it. Examples:
  * Perfect-square polynomial question: FIRST pick a quadratic like (2x²-3x+2),
    square it, and use THAT expansion as the question polynomial. This guarantees
    the question is solvable.
  * Square root / factorisation: construct from known factors.
  * Statistics: pick data whose mean and SD are clean by design.
- If while writing the Answer Key you discover the question is unsolvable or messy,
  you MUST go back and replace the QUESTION itself with the corrected version,
  so the printed question matches its solution perfectly.

[STATISTICS TWO-DATASET RULE]
- If a statistics question gives two datasets (e.g. income x and expenditure y),
  the question must state EXPLICITLY what to compute:
  e.g. "இரு தரவுத்தொகுப்புகளின் மாறுபாட்டுக் கெழுக்களை (C.V.) கணக்கிட்டு ஒப்பிடுக" -
  and the Answer Key must then compute BOTH datasets fully.
- Never give two datasets and compute only one.

[MENSURATION VALUES RULE]
- When π = 22/7 is used, choose radius/height values that are multiples of 7
  so answers come out as whole numbers (e.g. r=7, r=14, r=21, h=7).
- Express intermediate fractions exactly (33000/7) only if they cancel later;
  final answers must be clean numbers.
"""

    quality_rules = """
[QUESTION PAPER QUALITY RULES - MANDATORY]
1. THEOREM DIAGRAMS ARE COMPULSORY:
   - EVERY theorem proof question (தேற்றம் எழுதி நிரூபிக்கவும்) MUST have its diagram tag.
   - தேல்ஸ் தேற்றம் → [DRAW_THALES] is NON-NEGOTIABLE. Never skip it.
   - A theorem question WITHOUT its diagram tag is an INVALID output.

2. CONSTRUCTION QUESTIONS - STRICT SCOPE:
   - The note "(உரிய அளவுகளுடன் வரைபடம் வரைந்து காட்டவும்)" must be added ONLY to
     COMPASS-AND-RULER CONSTRUCTION questions - i.e. questions whose main verb is
     "வரைக" or "அமைக்க" (construct/draw).
   - NEVER add this note to: prove questions (நிரூபி), calculation questions (காண்க),
     theorem questions, graph questions, or any non-construction question.
   - EVERY construction question MUST state ALL required measurements explicitly
     in the question text (side lengths, radius, scale factor etc). A construction
     question without complete measurements is INVALID - redesign it with numbers.

3. QUESTION CLARITY:
   - NEVER write vague questions like "வடிவொத்த முக்கோணங்கள் வரைய பயன்படும் தேற்றத்தை எழுதுக".
   - Always name the EXACT theorem: e.g. "அடிப்படை விகிதசம தேற்றத்தை (தேல்ஸ் தேற்றம்) எழுதி நிரூபிக்கவும்".
   - If a question asks about angle bisector, specify internal (உட்புற) or external (வெளிப்புற)
     AND state clearly where the intersection point lies (e.g. "D என்பது BC-ன் நீட்டிப்பில் உள்ளது").

4. MULTI-PART QUESTIONS:
   - If a question has two tasks (e.g. "நிரூபிக்கவும்" + "மதிப்பு காண்க"),
     split into (அ) and (ஆ) sub-parts explicitly.

5. ANSWER KEY CONSISTENCY:
   - The ratio/expression in the Answer Key MUST be IDENTICAL to the one asked in the question.
   - If question asks AO/BO = CO/DO, answer key must use AO/BO = CO/DO. Never switch to a different
     equivalent form. Never write "(அல்லது ...)" alternate forms.

6. REALISTIC VALUES:
   - Design numerical questions so answers are clean values (whole numbers or simple decimals
     like 4.5). Avoid awkward answers like 0.4 or 9.6 in geometry length problems.

7. MCQ QUALITY STANDARD:
   - Avoid trivial rote-memory MCQs (e.g. asking just the count of elements in a 3x2 matrix).
   - Each MCQ should require at least one step of thinking or calculation.
   - Mix: 40% concept application, 40% one-step calculation, maximum 20% pure recall.

8. BILINGUAL TECHNICAL TERMS:
   - For key technical concepts in Tamil medium papers, add the English term in
     brackets on first use: "முடிவிலா எண்ணிக்கையிலான தீர்வுகள் (infinitely many solutions)",
     "திட்ட விலக்கம் (Standard Deviation)". This matches TN Board pattern.
"""

    pyq_tagging_rule = """
[PREVIOUS YEAR QUESTION (PYQ) TAGGING RULE - TN BOARD PUBLIC EXAMS]
- Scope: Apply ONLY to 2-Mark, 5-Mark, and 8-Mark/Long-answer questions.
- Do NOT apply to 1-Mark MCQ questions under any circumstance.
- If a generated question appeared in a Tamil Nadu State Board PUBLIC EXAM
  in the last 4 years (2022, 2023, 2024, 2025), append the exam session tag
  at the END of that question in brackets.
- Tag format (Tamil papers):  (செப் 2024) / (ஏப்ரல் 2023) / (ஜூன் 2022)
- Tag format (English papers): (Sep 2024) / (Apr 2023) / (Jun 2022)
- Multiple appearances: list all, comma-separated: (ஏப்ரல் 2023, செப் 2024)
- CRITICAL ACCURACY RULE: Tag ONLY if you are HIGHLY CONFIDENT the question
  appeared in that specific exam session. If uncertain about the year or month,
  DO NOT tag at all. An untagged repeated question is acceptable;
  a WRONGLY tagged question is NOT acceptable.
- Do NOT invent or guess exam sessions. Do NOT tag more than 40% of questions.
"""

    no_latex_rule = """
[ABSOLUTE NO-LATEX FORMATTING RULE]
- NEVER use LaTeX syntax in the output. This is a Word document, not LaTeX.
- BANNED: \\{ \\} \\rightarrow \\leftarrow \\times \\div \\le \\ge \\ne \\in \\subset \\cup \\cap \\sqrt{} \\frac{}{} \\pi \\theta \\alpha \\beta \\angle \\triangle \\cdot $...$ etc.
- USE plain Unicode symbols instead:
  * Sets: {1, 2, 3} — plain curly braces WITHOUT backslash
  * Arrow: → (f: A → B)
  * Multiply: × | Divide: ÷ | Not equal: ≠ | Less/greater equal: ≤ ≥
  * Element of: ∈ | Subset: ⊂ | Union: ∪ | Intersection: ∩
  * Square root: √2, √(x+1) | Fraction: a/b or (x+1)/(x-2)
  * Pi: π | Theta: θ | Degree: 45° | Angle: ∠ABC | Triangle: △ABC
  * Squared: x² | Cubed: x³ | Subscript: a₁, a₂, aₙ
- Example CORRECT: "A = {1, 2, 3, 4}, B = {2, 5, 8} மற்றும் f: A → B ஆனது f(x) = 3x - 1"
- Example WRONG: "A = \\{1, 2, 3, 4\\}, f: A \\rightarrow B"
"""
    return f"Subject: {subject}\nLessons: {lessons_str}\nExam Type: {exam_type}\nTotal Marks: {total_marks}\nTime: {exam_time}\nMode: {exam_mode}\n{difficulty_directive}\n{blueprint_desc}\n{header_format}\n{option_format}\n{lang_instruction}\n{subject_blueprint_rules}\n{no_latex_rule}\n{quality_rules}\n{theorem_proof_rule}\n{pyq_tagging_rule}\n{no_latex_rule}\n=== ANSWER KEY ==="


# ==========================================
# LaTeX → Normal Text Converter
# ==========================================
LATEX_REPLACEMENTS = [
    (r'\\rightarrow', '→'), (r'\\to\b', '→'), (r'\\leftarrow', '←'),
    (r'\\Rightarrow', '⇒'), (r'\\leftrightarrow', '↔'),
    (r'\\in\b', '∈'), (r'\\notin', '∉'), (r'\\subseteq', '⊆'), (r'\\subset', '⊂'),
    (r'\\cup', '∪'), (r'\\cap', '∩'), (r'\\emptyset', '∅'), (r'\\phi', 'φ'),
    (r'\\times', '×'), (r'\\div', '÷'), (r'\\pm', '±'),
    (r'\\cdot', '·'), (r'\\neq', '≠'), (r'\\ne\b', '≠'),
    (r'\\leq', '≤'), (r'\\le\b', '≤'), (r'\\geq', '≥'), (r'\\ge\b', '≥'),
    (r'\\approx', '≈'), (r'\\equiv', '≡'), (r'\\infty', '∞'),
    (r'\\angle', '∠'), (r'\\triangle', '△'),
    (r'\\perp', '⊥'), (r'\\parallel', '∥'), (r'\\cong', '≅'), (r'\\sim\b', '∼'),
    (r'\\degree', '°'), (r'\\circ\b', '°'),
    (r'\\alpha', 'α'), (r'\\beta', 'β'), (r'\\gamma', 'γ'), (r'\\delta', 'δ'),
    (r'\\theta', 'θ'), (r'\\lambda', 'λ'), (r'\\mu\b', 'μ'), (r'\\pi\b', 'π'),
    (r'\\sigma', 'σ'), (r'\\omega', 'ω'), (r'\\Delta', 'Δ'), (r'\\Sigma', 'Σ'),
    (r'\\mathbb\{R\}', 'ℝ'), (r'\\mathbb\{N\}', 'ℕ'), (r'\\mathbb\{Z\}', 'ℤ'), (r'\\mathbb\{Q\}', 'ℚ'),
]

SUPERSCRIPT_MAP = {'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','n':'ⁿ','-':'⁻','+':'⁺'}
SUBSCRIPT_MAP   = {'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','n':'ₙ','-':'₋','+':'₊'}

def latex_to_normal(text):
    """LaTeX symbols → normal Unicode text"""
    text = re.sub(r'\\d?frac\{([^{}]+)\}\{([^{}]+)\}', r'(\1/\2)', text)
    text = re.sub(r'\\sqrt\{([^{}]+)\}', r'√(\1)', text)
    text = re.sub(r'\\sqrt\s*', '√', text)
    for pattern, repl in LATEX_REPLACEMENTS:
        text = re.sub(pattern, repl, text)
    text = re.sub(r'\^\{([^{}]+)\}', lambda m: ''.join(SUPERSCRIPT_MAP.get(c, c) for c in m.group(1)), text)
    text = re.sub(r'\^(\d)', lambda m: SUPERSCRIPT_MAP.get(m.group(1), m.group(1)), text)
    text = re.sub(r'_\{([^{}]+)\}', lambda m: ''.join(SUBSCRIPT_MAP.get(c, c) for c in m.group(1)), text)
    text = re.sub(r'_(\d)', lambda m: SUBSCRIPT_MAP.get(m.group(1), m.group(1)), text)
    text = text.replace('\\{', '{').replace('\\}', '}')
    text = text.replace('$', '')
    text = re.sub(r'\\(?=[a-zA-Z])', '', text)  # strip leftover backslashes
    return text

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
    text = latex_to_normal(text)
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        draw_match = re.search(r'\[DRAW_(THALES|BPT|EXT_BISECTOR|EXTERNAL_BISECTOR|ANGLE_BISECTOR|BISECTOR|TWO_TANGENTS|TWO_TANGENT|PYTHAGORAS|RIGHT_TRIANGLE|TANGENT|TRIANGLE|SQUARE|RECTANGLE|CIRCLE|SEMICIRCLE)[:\s]*(.*?)\]', line, re.IGNORECASE)
        if draw_match:
            shape_type = draw_match.group(1)
            label_text = draw_match.group(2)
            try:
                img_buf = generate_geometry_image(shape_type, label_text)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_buf, width=Inches(2.5))
            except Exception as e:
                doc.add_paragraph(f"[Error loading diagram: {e}]")
            continue
        clean_line = line.replace('*', '').replace('$', '').strip()
        if "பகுதி" in clean_line or "PART" in clean_line.upper():
            marks_match = re.search(r'\(?\d+\s*[xX*]\s*\d+\s*=\s*\d+\)?', clean_line)
            if marks_match:
                calc_str  = marks_match.group(0)
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
                    if current.strip():
                        parts.append(current.strip())
                    current = chunk + " "
                else:
                    current += chunk
            if current.strip():
                parts.append(current.strip())
            if parts:
                table = doc.add_table(rows=1, cols=len(parts))
                for idx, opt in enumerate(parts):
                    cell = table.cell(0, idx)
                    cell.paragraphs[0].add_run(opt.replace("*", ""))
                    set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
                continue
        p = doc.add_paragraph()
        if re.match(r'^\d+\.', clean_line):
            p.paragraph_format.left_indent       = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)
        parts = re.split(r'\*\*(.*?)\*\*', line)
        for i, part in enumerate(parts):
            run = p.add_run(part.replace('$', ''))
            if i % 2 == 1:
                run.bold = True

def create_professional_docx(ai_response, school_name, class_val, subject_val, exam_type, time_val, marks_val):
    doc = Document()
    section = doc.sections[0]
    section.page_width = section.page_height = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.5)
    style = doc.styles['Normal']
    style.font.name = 'Nirmala UI'
    h_school = doc.add_paragraph(style='Normal')
    h_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_school.add_run(school_name.upper()).bold = True
    h_school.runs[0].font.size = Pt(15)
    table = doc.add_table(rows=2, cols=2)
    def format_cell(cell, text, align_right=False):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align_right else WD_ALIGN_PARAGRAPH.LEFT
        p.add_run(text).bold = True
        set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
    format_cell(table.cell(0, 0), f"Class : {class_val}")
    format_cell(table.cell(0, 1), f"Time : {time_val}", align_right=True)
    format_cell(table.cell(1, 0), f"Subject : {subject_val}")
    format_cell(table.cell(1, 1), f"Marks : {marks_val}", align_right=True)
    add_solid_line(doc)
    parts = ai_response.split("=== ANSWER KEY ===")
    write_markdown_to_word(doc, parts[0].strip())
    return doc

# ==========================================
# MAIN APP TABS
# ==========================================
tab1, tab2 = st.tabs(["🎓 வினாத்தாள் தயாரிப்பு", "📝 விடைத்தாள் திருத்தம்"])

with tab1:
    st.title("🎓 PMP Master AI Engine (V20.0)")
    df = load_data()
    if not df.empty:
        col1, col2 = st.columns(2)
        with col1:
            school_name    = st.text_input("School Name", value="ABC SCHOOL")
            class_val      = st.selectbox("Class", ["10"])
            subject_list   = df['Subject'].unique()
            subject_val    = st.selectbox("Subject", subject_list)
        with col2:
            exam_type  = st.selectbox("Exam Type", ["Unit Test", "Quarterly Exam", "Half-Yearly Exam", "Annual Exam"])
            time_val   = st.selectbox("Time (நேரம்)", ["1.00 Hour", "2.00 Hours", "3.00 Hours"], index=2)
            marks_val  = st.number_input("Total Marks", value=100, step=1)

        exam_mode        = st.selectbox("Exam Mode", ["🏛️ Public Exam Mode", "🏫 School Elite Mode"])
        lesson_list      = df[df['Subject'] == subject_val]['Lesson'].unique()
        selected_lessons = st.multiselect("பாடங்களைத் தேர்ந்தெடுக்கவும்", lesson_list)

        st.markdown("---")
        diff_level = st.selectbox("வினாத்தாள் கடினத்தன்மை", ["எளிமை (Easy)", "நடுத்தரம் (Medium)", "கடினம் (Hard)"], index=1)
        st.markdown("---")

        is_eng = "english" in subject_val.lower() or "ஆங்கிலம்" in subject_val.lower()
        is_soc = "social"  in subject_val.lower() or "சமூக"     in subject_val.lower()
        bp     = get_blueprint_defaults(marks_val, is_social=is_soc, is_english=is_eng)

        st.markdown("### 📋 வினா வடிவமைப்பு பிரிவு")
        show_p1 = st.checkbox("பகுதி I (1-Mark Questions)",   value=True)
        show_p2 = st.checkbox("பகுதி II (2-Mark Questions)",  value=True)
        show_p3 = st.checkbox("பகுதி III (5-Mark Questions)", value=True)
        show_p4 = st.checkbox("பகுதி IV (Long Questions)",    value=True)

        st.markdown("#### ⚙️ மதிப்பெண் விவரங்கள்")
        b1, b2 = st.columns(2)
        with b1:
            p1_ask  = st.number_input("1-மார்க் வினாக்கள்",       value=int(bp["p1"])  if show_p1 else 0, step=1, disabled=not show_p1)
            p2_get  = st.number_input("2-மார்க் கொடுக்க (Given)", value=int(bp["p2g"]) if show_p2 else 0, step=1, disabled=not show_p2)
            p2_ask  = st.number_input("2-மார்க் எழுத (Answer)",   value=int(bp["p2a"]) if show_p2 else 0, step=1, disabled=not show_p2)
        with b2:
            p3_get  = st.number_input("5-மார்க் கொடுக்க (Given)", value=int(bp["p3g"]) if show_p3 else 0, step=1, disabled=not show_p3)
            p3_ask  = st.number_input("5-மார்க் எழுத (Answer)",   value=int(bp["p3a"]) if show_p3 else 0, step=1, disabled=not show_p3)
            p4_val  = st.selectbox("நெடுவினா மதிப்பெண்", [5, 8, 10], index=1 if is_eng or is_soc or marks_val==100 else 0, disabled=not show_p4)
            p4_get  = st.number_input("நெடுவினா கொடுக்க (Given)", value=int(bp["p4g"]) if show_p4 else 0, step=1, disabled=not show_p4)
            p4_ask  = st.number_input("நெடுவினா எழுத (Answer)",   value=int(bp["p4a"]) if show_p4 else 0, step=1, disabled=not show_p4)

        total_calculated = (p1_ask * 1) + (p2_ask * 2) + (p3_ask * 5) + (p4_ask * p4_val)
        can_generate     = total_calculated == marks_val

        if can_generate:
            st.success(f"✅ மதிப்பெண்கள் சரியாகப் பொருந்தியது: {total_calculated} மார்க்.")
        else:
            st.warning(f"⚠️ கணக்கீடு: {total_calculated} | மொத்தம்: {marks_val} (சமப்படுத்தவும்)")

        st.markdown("---")
        gen_mode = st.radio(
            "வினாத்தாள் உருவாக்கும் முறை",
            ["🤖 AI Auto-Generate (தற்போதைய முறை)", "📚 கேள்வி வங்கியில் இருந்து தேர்ந்தெடு (Book Back / Exercise / Example)"],
            horizontal=False,
        )
        bank_mode = gen_mode.startswith("📚")

        parts_meta = [
            {"key": "p1", "label": "பகுதி I",   "mark_type": "1M",   "mark": 1,     "given": int(p1_ask), "answer": int(p1_ask), "show": show_p1, "note": ""},
            {"key": "p2", "label": "பகுதி II",  "mark_type": "2M",   "mark": 2,     "given": int(p2_get), "answer": int(p2_ask), "show": show_p2, "note": (f"ஏதேனும் {int(p2_ask)} கேள்விகளுக்கு மட்டும் விடையளிக்கவும்" if p2_ask != p2_get else "")},
            {"key": "p3", "label": "பகுதி III", "mark_type": "5M",   "mark": 5,     "given": int(p3_get), "answer": int(p3_ask), "show": show_p3, "note": (f"ஏதேனும் {int(p3_ask)} கேள்விகளுக்கு மட்டும் விடையளிக்கவும்" if p3_ask != p3_get else "")},
            {"key": "p4", "label": "பகுதி IV",  "mark_type": "LONG", "mark": int(p4_val), "given": int(p4_get), "answer": int(p4_ask), "show": show_p4, "note": (f"ஏதேனும் {int(p4_ask)} கேள்விகளுக்கு மட்டும் விடையளிக்கவும்" if p4_ask != p4_get else "")},
        ]

        if bank_mode:
            st.markdown("### 📚 கேள்வி வங்கி — தேர்ந்தெடுக்கவும்")
            if not selected_lessons:
                st.warning("⚠️ முதலில் மேலே பாடங்களைத் தேர்ந்தெடுக்கவும்!")
            else:
                if st.button("🔄 கேள்வி வங்கியை ஏற்று / புதுப்பி (Load Question Bank)", use_container_width=True):
                    with st.spinner("⏳ புத்தக கேள்விகள் தயார் செய்யப்படுகிறது... (முதல் முறை கொஞ்சம் நேரம் ஆகும்)"):
                        pool = {}
                        for part in parts_meta:
                            if not part["show"] or part["given"] <= 0:
                                continue
                            merged = []
                            per_lesson_min = max(4, (part["given"] // max(len(selected_lessons), 1)) + 3)
                            for lesson in selected_lessons:
                                merged.extend(get_or_build_bank(subject_val, lesson, part["mark_type"], min_count=per_lesson_min))
                            pool[part["key"]] = merged
                        st.session_state["bank_pool"] = pool
                        st.session_state["bank_pool_subject"] = subject_val
                        st.success("✅ கேள்வி வங்கி தயார்!")

                if "bank_pool" in st.session_state and st.session_state.get("bank_pool_subject") == subject_val:
                    for part in parts_meta:
                        if not part["show"] or part["given"] <= 0:
                            continue
                        pool = st.session_state["bank_pool"].get(part["key"], [])
                        if not pool:
                            continue
                        with st.expander(f'{part["label"]} — {MARK_TYPE_LABELS.get(part["mark_type"], part["mark_type"])} (தேவை: {part["given"]})', expanded=False):
                            col_left, col_right = st.columns([3, 2])
                            selected_items = []
                            with col_left:
                                st.caption("⬅️ கிடைக்கும் கேள்விகள் — tick செய்யவும்")
                                for it in pool:
                                    chk_key = f'chk_{part["key"]}_{it["id"]}'
                                    tag = f'[{it.get("qtype","")}] '
                                    st.checkbox(f'{tag}{it["question_text"]}', key=chk_key)
                            with col_right:
                                st.caption("➡️ தேர்ந்தெடுக்கப்பட்டவை")
                                for it in pool:
                                    chk_key = f'chk_{part["key"]}_{it["id"]}'
                                    if st.session_state.get(chk_key):
                                        selected_items.append(it)
                                        st.markdown(f'✅ {it["question_text"][:60]}{"..." if len(it["question_text"])>60 else ""}')
                                got = len(selected_items)
                                need = part["given"]
                                if got == need:
                                    st.success(f'{got}/{need} தேர்ந்தெடுக்கப்பட்டது ✅')
                                elif got > need:
                                    st.error(f'{got}/{need} — {got-need} அதிகமாக உள்ளது, சிலவற்றை நீக்கவும்')
                                else:
                                    st.warning(f'{got}/{need} தேர்ந்தெடுக்கப்பட்டது')

        if st.button("🚀 Generate PRO Question Paper", use_container_width=True, type="primary"):
            if bank_mode:
                if not selected_lessons:
                    st.warning("⚠️ பாடங்களைத் தேர்ந்தெடுக்கவும்!")
                elif "bank_pool" not in st.session_state:
                    st.warning("⚠️ முதலில் 'கேள்வி வங்கியை ஏற்று' பொத்தானை அழுத்தவும்!")
                else:
                    parts_cfg = []
                    all_ok = True
                    for part in parts_meta:
                        if not part["show"] or part["given"] <= 0:
                            continue
                        pool = st.session_state["bank_pool"].get(part["key"], [])
                        chosen = [it for it in pool if st.session_state.get(f'chk_{part["key"]}_{it["id"]}')]
                        if len(chosen) != part["given"]:
                            all_ok = False
                        parts_cfg.append({
                            "label": part["label"], "mark": part["mark"],
                            "given": part["given"], "answer": part["answer"],
                            "note": part["note"], "items": chosen,
                        })
                    if not all_ok:
                        st.error("⚠️ ஒவ்வொரு பகுதியிலும் தேவையான எண்ணிக்கை கேள்விகளை சரியாக தேர்ந்தெடுக்கவும் (மேலே உள்ள எண்களைப் பார்க்கவும்).")
                    else:
                        with st.spinner("⏳ தேர்ந்தெடுக்கப்பட்ட கேள்விகளுடன் வினாத்தாள் தயாராகிறது..."):
                            assembled_text = assemble_paper_from_bank(parts_cfg)
                            increment_usage(user_id)
                            doc    = create_professional_docx(assembled_text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            st.session_state['docx_bytes'] = doc_io.getvalue()
                            st.success("✅ வினாத்தாள் தயாராகிவிட்டது! (கேள்வி வங்கியில் இருந்து)")
            elif can_generate and selected_lessons:
                with st.spinner("⏳ வினாத்தாள் தயாராகிறது..."):
                    blueprint_desc = f"- Part I: {p1_ask} Qs. - Part II: Given {p2_get}, Answer {p2_ask}. - Part III: Given {p3_get}, Answer {p3_ask}. - Part IV: Given {p4_get}, Answer {p4_ask}."
                    prompt = generate_prompt_v18(subject_val, selected_lessons, exam_type, time_val, marks_val, exam_mode, blueprint_desc, p1_ask, p2_ask, p3_ask, diff_level)
                    response = None
                    for attempt in range(4):
                        try:
                            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                            break
                        except Exception as api_err:
                            if ("429" in str(api_err) or "503" in str(api_err)) and attempt < 3:
                                time.sleep((attempt + 1) * 4)
                            else:
                                st.error(f"சர்வர் பிழை: {api_err}")
                    if response:
                        increment_usage(user_id)
                        doc    = create_professional_docx(response.text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        st.session_state['docx_bytes'] = doc_io.getvalue()
                        st.success("✅ வினாத்தாள் தயாராகிவிட்டது!")
            elif not selected_lessons:
                st.warning("⚠️ பாடங்களைத் தேர்ந்தெடுக்கவும்!")

        if 'docx_bytes' in st.session_state:
            st.download_button(
                label="📥 Download Word File (.docx)",
                data=st.session_state['docx_bytes'],
                file_name=f"PMP_{subject_val}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    else:
        st.warning("⚠️ lesson_master_v1_5.csv கோப்பு கிடைக்கவில்லை.")

with tab2:
    st.title("📝 AI Answer Sheet Evaluator")
    uploaded_file = st.file_uploader("விடைத்தாளைத் தேர்ந்தெடுக்கவும் (Image / PDF)", type=["png", "jpg", "jpeg", "pdf"])
    if uploaded_file is not None:
        file_type   = uploaded_file.name.split(".")[-1].lower()
        eval_payload= []
        if file_type == "pdf":
            st.info("📊 PDF கோப்பு கண்டறியப்பட்டது.")
            file_data = uploaded_file.read()
            pdf_part  = types.Part.from_bytes(data=file_data, mime_type="application/pdf")
            eval_payload.append(pdf_part)
        else:
            image = Image.open(uploaded_file)
            st.image(image, caption="Uploaded Handwriting Page", width=420)
            eval_payload.append(image)

        if st.button("🚀 Start AI Evaluation", use_container_width=True):
            with st.spinner("⏳ AI திருத்தி வருகிறது..."):
                eval_prompt = "You are an official TN Board Math Evaluator. Read handwriting or PDF pages and correct step-by-step. Write fractions as $\\frac{a}{b}$ inside single dollar signs. Respond in Tamil."
                eval_payload.append(eval_prompt)
                response = None
                for attempt in range(3):
                    try:
                        response = client.models.generate_content(model='gemini-2.5-flash', contents=eval_payload)
                        break
                    except Exception as eval_err:
                        if "429" in str(eval_err) and attempt < 2:
                            time.sleep(4)
                        else:
                            st.error(f"மதிப்பீட்டு சர்வர் பிழை: {eval_err}")
                if response:
                    st.markdown(response.text)

    st.markdown("---")
    col_l, col_m, col_r = st.columns([1, 2, 1])
    with col_m:
        # User info with profile picture
        if user_pic:
            st.markdown(f"""
            <div style='text-align:center; margin-bottom:8px;'>
                <img src="{user_pic}" width="48" height="48"
                     style='border-radius:50%; border:2px solid #3b82f6;'/>
            </div>
            """, unsafe_allow_html=True)
        st.markdown(f"**👤 {user_name}** ({user_email})")
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.pop("logged_in_user", None)
            st.rerun()
