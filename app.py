import streamlit as st

# 1. எளிய லாகின் லாஜிக் (எந்த லைப்ரரியும் தேவையில்லை)
def check_password():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.title("🔐 PMP Master AI Login")
        user = st.text_input("User ID")
        pwd = st.text_input("Password", type="password")
        if st.button("Login"):
            if user == "admin" and pwd == "pmp123":
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ தவறான ID அல்லது Password!")
        return False
    return True

# 2. மெயின் ஆப் (லாகின் ஆன பிறகு மட்டும் வேலை செய்யும்)
if check_password():
    st.sidebar.button("Logout", on_click=lambda: st.session_state.update({"logged_in": False}))
    
    # ----------------------------------------------------
    # இங்கிருந்து உங்கள் பழைய V20.8 வினாத்தாள் ஜெனரேட்டர் கோட் தொடரும்
    # ----------------------------------------------------
    st.title("🎓 PMP Master AI Engine (V22.0)")
    # ... உங்கள் வினாத்தாள் தயாரிப்பு கோட் அனைத்தும் இங்கே ...
