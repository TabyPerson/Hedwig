import streamlit as st
from .database import get_connection
from .utils import hash_password, is_valid_email, password_strength

def register_screen():
    st.title("User Registration")
    st.markdown('**Full Name <span style="color:red;">*</span>**', unsafe_allow_html=True)
    name = st.text_input("", max_chars=40, key="name", label_visibility="collapsed")
    st.markdown('**Email <span style="color:red;">*</span>**', unsafe_allow_html=True)
    email = st.text_input("", max_chars=40, key="email", label_visibility="collapsed")
    company = st.text_input("**Company**", max_chars=20)
    role = st.text_input("**Role**", max_chars=20)
    st.markdown('**Password <span style="color:red;">*</span>**', unsafe_allow_html=True)
    password = st.text_input("", type="password", key="password", label_visibility="collapsed")
    st.markdown('**Confirm password <span style="color:red;">*</span>**', unsafe_allow_html=True)
    confirm = st.text_input("", type="password", key="confirm", label_visibility="collapsed")
    access = st.selectbox(
        "Access permission*",
        [
            "verification",  # access to verification comparisons
            "validation",    # access to validation comparisons
            "registration",  # access to user registration screen
            "all"            # full access
        ],
        format_func=lambda x: {
            "verification": "Verification",
            "validation": "Validation",
            "registration": "User Registration",
            "all": "All"
        }[x]
    )
    password = st.text_input("Password*", type="password")
    confirm = st.text_input("Confirm password*", type="password")
    st.write(f"Password strength: {password_strength(password)}")
    if st.button("Register"):
        if not name or not email or not password or not confirm:
            st.error("Please fill in all required fields.")
        elif not is_valid_email(email):
            st.error("Invalid email.")
        elif password != confirm:
            st.error("Passwords do not match.")
        elif password_strength(password) == "Weak":
            st.error("Weak password. Use uppercase, lowercase, numbers and special characters.")
        else:
            conn = get_connection()
            c = conn.cursor()
            try:
                c.execute("INSERT INTO users (name, email, company, role, password, access) VALUES (?, ?, ?, ?, ?, ?)",
                          (name, email, company, role, hash_password(password), access))
                conn.commit()
                st.session_state["show_register"] = False  # Close registration screen
                st.success("User registered successfully! Please log in.")
                st.experimental_rerun()
            except Exception as e:
                st.error("Error registering user or email already exists.")
            conn.close()