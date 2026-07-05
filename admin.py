import streamlit as st
import psycopg2
import psycopg2.extras
import datetime
import resend

st.set_page_config(page_title="PMP Admin Dashboard", page_icon="🔧", layout="wide")

# ==========================================
# ADMIN CONFIG
# ==========================================
ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", "sivarobo@gmail.com")

# ==========================================
# DB Connection
# ==========================================
def get_db():
    try:
        conn = psycopg2.connect(
            st.secrets["NEON_DATABASE_URL"],
            cursor_factory=psycopg2.extras.RealDictCursor
        )
        return conn
    except Exception as e:
        st.error(f"DB Error: {e}")
        return None

# ==========================================
# Admin Auth Check
# ==========================================
def check_admin_access():
    if st.session_state.get("admin_logged_in"):
        return True

    st.markdown("""
    <div style='text-align:center; padding:40px 0 10px 0;'>
        <h1 style='color:#1E3A8A;'>🔧 PMP Admin Dashboard</h1>
        <p style='color:#64748b;'>Admin Access Only</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("#### Admin Email உள்ளிடுங்கள்:")
        admin_email = st.text_input("Email", placeholder="admin@gmail.com", key="admin_email_input")

        if st.button("🔑 Login", use_container_width=True, type="primary"):
            if admin_email.strip().lower() == ADMIN_EMAIL:
                import random, time
                otp = str(random.randint(100000, 999999))
                st.session_state["admin_otp"] = otp
                st.session_state["admin_otp_expiry"] = time.time() + 600
                st.session_state["admin_email_input_val"] = admin_email.strip().lower()

                # Send OTP
                try:
                    resend.api_key = st.secrets["RESEND_API_KEY"]
                    resend.Emails.send({
                        "from": "PMP AI Suite <onboarding@resend.dev>",
                        "to": [ADMIN_EMAIL],
                        "subject": "🔧 PMP Admin OTP",
                        "html": f"<h2>Admin OTP: <b style='color:#1e3a8a;font-size:32px;'>{otp}</b></h2><p>10 நிமிடம் valid.</p>"
                    })
                    st.success(f"✅ OTP அனுப்பப்பட்டது! {ADMIN_EMAIL} check பண்ணுங்கள்.")
                    st.session_state["show_admin_otp"] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"OTP அனுப்ப முடியவில்லை: {e}")
            else:
                st.error("❌ நீங்கள் Admin இல்லை!")
        return False

def check_admin_otp():
    import time
    st.markdown("""
    <div style='text-align:center; padding:20px;'>
        <h2>📧 Admin OTP Verification</h2>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        otp_input = st.text_input("6-digit OTP:", max_chars=6, placeholder="123456")
        if st.button("✅ Verify", use_container_width=True, type="primary"):
            if time.time() > st.session_state.get("admin_otp_expiry", 0):
                st.error("OTP காலாவதி! மீண்டும் try பண்ணுங்கள்.")
                for k in ["admin_otp", "admin_otp_expiry", "show_admin_otp"]:
                    st.session_state.pop(k, None)
                st.rerun()
            elif otp_input.strip() == st.session_state.get("admin_otp", ""):
                st.session_state["admin_logged_in"] = True
                for k in ["admin_otp", "admin_otp_expiry", "show_admin_otp"]:
                    st.session_state.pop(k, None)
                st.rerun()
            else:
                st.error("❌ தவறான OTP!")

        if st.button("← Back", use_container_width=True):
            for k in ["admin_otp", "admin_otp_expiry", "show_admin_otp"]:
                st.session_state.pop(k, None)
            st.rerun()

# ==========================================
# ACCESS GATE
# ==========================================
if st.session_state.get("show_admin_otp"):
    check_admin_otp()
    st.stop()

if not check_admin_access():
    st.stop()

# ==========================================
# MAIN ADMIN DASHBOARD
# ==========================================
st.markdown("""
<style>
    .metric-card {
        background: var(--background-color);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# Header
col_h1, col_h2 = st.columns([4, 1])
with col_h1:
    st.title("🔧 PMP Admin Dashboard")
    st.caption(f"Logged in as: {ADMIN_EMAIL}")
with col_h2:
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.pop("admin_logged_in", None)
        st.rerun()

st.markdown("---")

# ==========================================
# STATS
# ==========================================
conn = get_db()
if conn:
    cur = conn.cursor()

    # Total users
    cur.execute("SELECT COUNT(*) as total FROM users")
    total_users = cur.fetchone()["total"]

    # Today's active users
    cur.execute("SELECT COUNT(DISTINCT user_id) as active FROM daily_usage WHERE usage_date = CURRENT_DATE")
    today_active = cur.fetchone()["active"]

    # Today's total questions
    cur.execute("SELECT COALESCE(SUM(question_count), 0) as total FROM daily_usage WHERE usage_date = CURRENT_DATE")
    today_questions = cur.fetchone()["total"]

    # Premium users
    cur.execute("SELECT COUNT(*) as premium FROM users WHERE plan = 'premium'")
    premium_users = cur.fetchone()["premium"]

    # This month registrations
    cur.execute("SELECT COUNT(*) as monthly FROM users WHERE created_at >= date_trunc('month', NOW())")
    monthly_reg = cur.fetchone()["monthly"]

    # Total questions all time
    cur.execute("SELECT COALESCE(SUM(question_count), 0) as total FROM daily_usage")
    total_questions = cur.fetchone()["total"]

    conn.close()

    # Metric Cards
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.metric("👥 Total Users", total_users)
    with c2:
        st.metric("🟢 Today Active", today_active)
    with c3:
        st.metric("📄 Today Questions", today_questions)
    with c4:
        st.metric("⭐ Premium Users", premium_users)
    with c5:
        st.metric("📅 This Month New", monthly_reg)
    with c6:
        st.metric("📊 All Time Questions", total_questions)

st.markdown("---")

# ==========================================
# TABS
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["👥 All Users", "📊 Usage Stats", "✉️ Send Email", "⚙️ Manage Users"])

# ---- TAB 1: ALL USERS ----
with tab1:
    st.subheader("👥 Registered Users")

    conn = get_db()
    if conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT u.id, u.email, u.name, u.plan, u.created_at,
                   COALESCE(SUM(d.question_count), 0) as total_questions,
                   MAX(d.usage_date) as last_active
            FROM users u
            LEFT JOIN daily_usage d ON u.id::text = d.user_id::text
            GROUP BY u.id, u.email, u.name, u.plan, u.created_at
            ORDER BY u.created_at DESC
        """)
        users = cur.fetchall()
        conn.close()

        if users:
            import pandas as pd
            df = pd.DataFrame([dict(u) for u in users])
            df["created_at"] = pd.to_datetime(df["created_at"]).dt.strftime("%Y-%m-%d %H:%M")
            df["last_active"] = df["last_active"].astype(str)
            df.columns = ["ID", "Email", "Name", "Plan", "Registered", "Total Questions", "Last Active"]
            df = df.drop(columns=["ID"])
            st.dataframe(df, use_container_width=True, height=400)
            st.caption(f"மொத்தம் {len(df)} users")
        else:
            st.info("இன்னும் users இல்லை.")

# ---- TAB 2: USAGE STATS ----
with tab2:
    st.subheader("📊 Daily Usage (Last 30 Days)")

    conn = get_db()
    if conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT usage_date,
                   COUNT(DISTINCT user_id) as active_users,
                   SUM(question_count) as total_questions
            FROM daily_usage
            WHERE usage_date >= CURRENT_DATE - INTERVAL '30 days'
            GROUP BY usage_date
            ORDER BY usage_date DESC
        """)
        usage_data = cur.fetchall()
        conn.close()

        if usage_data:
            import pandas as pd
            df = pd.DataFrame([dict(u) for u in usage_data])
            df.columns = ["Date", "Active Users", "Questions Generated"]

            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**📈 Daily Questions**")
                st.bar_chart(df.set_index("Date")["Questions Generated"])
            with col_b:
                st.markdown("**👥 Daily Active Users**")
                st.bar_chart(df.set_index("Date")["Active Users"])

            st.dataframe(df, use_container_width=True)
        else:
            st.info("Usage data இல்லை.")

# ---- TAB 3: SEND EMAIL ----
with tab3:
    st.subheader("✉️ Users-க்கு Email அனுப்புங்கள்")

    email_type = st.selectbox("Email Type:", [
        "📢 Announcement (எல்லாருக்கும்)",
        "⭐ Premium Upgrade Offer",
        "🎉 Custom Message"
    ])

    subject = st.text_input("Subject:", value="PMP AI Suite - Important Update")
    message = st.text_area("Message (HTML OK):", height=150,
        value="<p>வணக்கம்!</p><p>PMP Master AI Suite-ல் புதிய features வந்துள்ளன.</p>")

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        target = st.selectbox("யாருக்கு அனுப்பணும்?", ["All Users", "Free Users Only", "Premium Users Only"])
    with col_s2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📧 Send Email", use_container_width=True, type="primary"):
            conn = get_db()
            if conn:
                cur = conn.cursor()
                if target == "All Users":
                    cur.execute("SELECT email, name FROM users")
                elif target == "Free Users Only":
                    cur.execute("SELECT email, name FROM users WHERE plan = 'free'")
                else:
                    cur.execute("SELECT email, name FROM users WHERE plan = 'premium'")
                recipients = cur.fetchall()
                conn.close()

                resend.api_key = st.secrets["RESEND_API_KEY"]
                success_count = 0
                with st.spinner(f"⏳ {len(recipients)} users-க்கு அனுப்புகிறோம்..."):
                    for r in recipients:
                        try:
                            html_body = f"""
                            <div style='font-family:Arial; max-width:500px; margin:auto; padding:20px;
                                        border:2px solid #3b82f6; border-radius:12px;'>
                                <h2 style='color:#1e3a8a;'>🎓 PMP Master AI Suite</h2>
                                <p>வணக்கம் {r['name']}!</p>
                                {message}
                                <hr>
                                <p style='color:#94a3b8; font-size:12px;'>PMP Master AI Suite</p>
                            </div>
                            """
                            resend.Emails.send({
                                "from": "PMP AI Suite <onboarding@resend.dev>",
                                "to": [r["email"]],
                                "subject": subject,
                                "html": html_body
                            })
                            success_count += 1
                        except:
                            pass

                st.success(f"✅ {success_count}/{len(recipients)} users-க்கு email அனுப்பப்பட்டது!")

# ---- TAB 4: MANAGE USERS ----
with tab4:
    st.subheader("⚙️ User Management")

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.markdown("#### 🔄 Plan மாற்றம்")
        search_email = st.text_input("User Email:", placeholder="user@gmail.com", key="manage_email")
        new_plan = st.selectbox("புதிய Plan:", ["free", "premium"])

        if st.button("✅ Plan Update பண்ணு", use_container_width=True):
            if search_email:
                conn = get_db()
                if conn:
                    cur = conn.cursor()
                    cur.execute("UPDATE users SET plan = %s WHERE email = %s RETURNING email", 
                               (new_plan, search_email.strip().lower()))
                    updated = cur.fetchone()
                    conn.commit()
                    conn.close()
                    if updated:
                        st.success(f"✅ {search_email} → {new_plan} plan update ஆச்சு!")
                    else:
                        st.error("❌ User கண்டறியவில்லை!")

    with col_m2:
        st.markdown("#### 🗑️ User Delete")
        del_email = st.text_input("Delete Email:", placeholder="user@gmail.com", key="del_email")

        if st.button("🗑️ User Delete பண்ணு", use_container_width=True, type="secondary"):
            if del_email:
                conn = get_db()
                if conn:
                    cur = conn.cursor()
                    # Delete usage first
                    cur.execute("""
                        DELETE FROM daily_usage WHERE user_id = 
                        (SELECT id::text FROM users WHERE email = %s)
                    """, (del_email.strip().lower(),))
                    cur.execute("DELETE FROM users WHERE email = %s RETURNING email",
                               (del_email.strip().lower(),))
                    deleted = cur.fetchone()
                    conn.commit()
                    conn.close()
                    if deleted:
                        st.success(f"✅ {del_email} delete ஆச்சு!")
                    else:
                        st.error("❌ User கண்டறியவில்லை!")

    st.markdown("---")
    st.markdown("#### 📋 Today's Active Users")
    conn = get_db()
    if conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT u.name, u.email, u.plan, d.question_count
            FROM daily_usage d
            JOIN users u ON d.user_id::text = u.id::text
            WHERE d.usage_date = CURRENT_DATE
            ORDER BY d.question_count DESC
        """)
        today_users = cur.fetchall()
        conn.close()

        if today_users:
            import pandas as pd
            df = pd.DataFrame([dict(u) for u in today_users])
            df.columns = ["Name", "Email", "Plan", "Questions Today"]
            st.dataframe(df, use_container_width=True)
        else:
            st.info("இன்று யாரும் use பண்ணவில்லை.")
