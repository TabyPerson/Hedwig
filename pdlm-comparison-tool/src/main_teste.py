import streamlit as st
from auth.login import login_screen
from auth.register import register_screen
import importlib

def main():
    # Checa se já está logado
    if "user" not in st.session_state or not st.session_state["user"]:
        user = login_screen()
        if not user:
            st.stop()
    else:
        user = st.session_state["user"]

    st.set_page_config(
        page_title="Hedwig",
        page_icon="🦉",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # Sidebar navigation
    with st.sidebar:
        menu = st.radio(
            "Menu",
            options=["**✅Verification**", "**☑️Validation**"] + (["**🕵️User Registration**"] if user[-1] == "all" else []),
            index=0
        )

        if menu == "**✅Verification**":
            verification_option = st.selectbox(
                "Choose the analysis to run:",
                [
                    "Select a option",
                    "PRS DOC x Requirements TM Comparison",
                    "Verification TM Requirements x Test Protocol Comparison",
                    "Verification TM APP x Test Protocol Comparison",
                    "Verification Test Protocol x Records Comparison",
                    "Verification TM APP x Test Records Comparison",
                    "Verification Test Records x PDSR Comparison",
                    "Product Verification Report",
                    "Verification Check Video"
                ],
                key="verification_option"
            )
            st.session_state["selected_analysis"] = ("verification", verification_option)

        if menu == "**☑️Validation**":
            validation_option = st.selectbox(
                "Choose the analysis to run:",
                [
                    "Select a option",
                    "URS DOC x TM APP Comparison",
                    "TM APP x Validation Test Protocol Comparison",
                    "Validation Test Protocol x Records Comparison",
                    "TM APP x Validation Test Records Comparison",
                    "Validation Test Records x PDSR Comparison",
                    "Product Validation Report",
                    "Validation Check Video"
                ],
                key="validation_option"
            )
            st.session_state["selected_analysis"] = ("validation", validation_option)

        # Espaço para empurrar o botão para baixo
        st.markdown("<br><br><br><br><br><br><br><br><br>", unsafe_allow_html=True)

        # CSS global: apenas tamanho e borda (sem cor!)
        st.markdown("""
        <style>
        div.stButton > button:first-child {
            font-weight: bold;
            border-radius: 6px;
            border: none;
            height: 3em;
            width: 50%;
        }
        </style>
        """, unsafe_allow_html=True)
        
        if st.button("**Logout**", key="logout_btn_sidebar"):
            st.session_state["show_logout_confirm"] = True

        # POPUP Streamlit acima do botão de logout
        if st.session_state.get("show_logout_confirm"):
            # CSS para colorir cada botão individualmente do popup
            st.markdown("""
                <style>
                /* Só afeta os botões do popup */
                div[data-testid="column"] div.stButton > button:first-child {
                    height: 3em;
                    width: 90%;
                    font-weight: bold;
                    border-radius: 6px;
                }
                /* Botão Yes (primeira coluna) */
                div[data-testid="column"]:nth-child(1) div.stButton > button:first-child {
                    background-color: #4169E1 !important;
                    color: white !important;
                    border: none !important;
                }
                /* Botão No (segunda coluna) */
                div[data-testid="column"]:nth-child(2) div.stButton > button:first-child {
                    background-color: #fff !important;
                    color: #4169E1 !important;
                    border: 2px solid #4169E1 !important;
                }
                </style>
            """, unsafe_allow_html=True)

            with st.container():
                st.markdown(
                    "<div style='text-align:left; font-size:18px; font-weight:bold;'>Do you really want to logout?</div>",
                    unsafe_allow_html=True
                )
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Yes", key="logout_yes_popup"):
                        del st.session_state["user"]
                        st.session_state["show_logout_confirm"] = False
                        st.info("Logged out! Please refresh the page.")
                with col2:
                    if st.button("No", key="logout_no_popup"):
                        st.session_state["show_logout_confirm"] = False    

    # Main area: show selected analysis or registration
    col1, col2, col3 = st.columns([9,7,5])
    with col2:
        st.image(r"C:\Users\310287618\OneDrive - Philips\EMR\strategy\TESTES_COPILOT\Edwiges1.png", width=200)
    st.markdown(
        """
        <div style="text-align: center;">
            <p style="font-size: 18px;">
                Welcome! Hedwig helps you compare requirements, protocols, records, defects, and videos across your documentation.<br>
                Select a comparison type and follow the prompts below.
            </p>
        </div>
        """ ,
        unsafe_allow_html=True
    )

    if menu == "**🕵️User Registration**":
        register_screen()
    else:
        selected = st.session_state.get("selected_analysis")
        if selected:
            analysis_type, option = selected
            # Só executa se não for a opção padrão
            if option and option != "Select a option":
                module_name = (
                    option.lower()
                    .replace(" x ", "_")
                    .replace(" ", "_")
                    .replace("-", "_")
                )
                try:
                    with st.spinner("Loading module..."):
                        comparison_module = importlib.import_module(f"options.{module_name}")
                        comparison_module.run_comparison()
                except ModuleNotFoundError:
                    st.error("The selected comparison option is not available.")    
  

if __name__ == "__main__":
    main()

