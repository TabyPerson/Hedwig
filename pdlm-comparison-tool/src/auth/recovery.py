import streamlit as st

def recovery_screen():
    st.title("Recuperação de Senha")
    email = st.text_input("Digite seu email cadastrado")
    if st.button("Enviar link de redefinição"):
        # Aqui você implementaria o envio de email real
        st.success(f"Se o email {email} estiver cadastrado, um link de redefinição será enviado.")