import streamlit as st
from .database import get_connection, init_db
from .utils import check_password
from .recovery import recovery_screen

def login_screen():
    init_db()
    col1, col2, col3 = st.columns([7,7,5])
    with col2:
        st.image(r"C:\Users\310287618\OneDrive - Philips\EMR\strategy\TESTES_COPILOT\Edwiges1.png", width=200)
    st.title("Login")    
    st.markdown('**User (email) <span style="color:red;">*</span>**', unsafe_allow_html=True)
    email = st.text_input("User (email)", max_chars=40, key="email", label_visibility="collapsed")
    st.markdown('**Password <span style="color:red;">*</span>**', unsafe_allow_html=True)
    password = st.text_input("Password", type="password", key="password", label_visibility="collapsed")
    st.markdown("""
        <style>
        div.stButton > button:first-child {
            background-color: #4169E1;
            color: white;
            font-weight: bold;
            border-radius: 6px;
            border: none;
            height: 3em;
            width: 100%;
        }
        </style>
    """, unsafe_allow_html=True)

    if st.button("**Login**"):
        conn = get_connection()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE email = ?", (email,))
        user = c.fetchone()
        conn.close()
        if user and check_password(password, user[5]):
            st.session_state["user"] = user
            st.session_state["__rerun__"] = True
            st.write("")  # Força Streamlit a reprocessar, mas não recarrega a página
        else:
            st.error("Invalid email or password. Please try again.")

    st.markdown(
        """
        <style>
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background: white;
            text-align: center;
            font-size: 11px;
            z-index: 9999;
            padding: 8px 0 4px 0;
            box-shadow: 0 -2px 8px rgba(0,0,0,0.03);
        }
        </style>
        <div class="footer">
            Copyrights and all other proprietary rights in any software and related documentation ("Software") made available to you rest exclusively with Philips or its licensors.
            No title or ownership in the Software is conferred to you. Use of the Software is subject to the end user license conditions as are available on request.<br>
            Version 1.0.0 - 2025
        </div>
        """,
        unsafe_allow_html=True
    )

