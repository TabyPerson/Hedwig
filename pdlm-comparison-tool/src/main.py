import streamlit as st
import importlib

def main():
    st.set_page_config(
        page_title="Hedwig",
        page_icon="🦉",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # Configuração de tema para todos os componentes Streamlit
    st.markdown("""
    <style>
    /* Estilos básicos para melhorar a aparência geral */
    .stButton > button {
        font-weight: bold !important;
        border-radius: 6px !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Sidebar navigation
    with st.sidebar:
        st.markdown("""
        <style>
        /* Força todos botões na sidebar a ficarem azul royal */
        [data-testid="stSidebar"] button {
            background-color: #4169E1 !important;
            color: white !important;
            font-weight: bold !important;
            border-radius: 6px !important;
            width: 100% !important;
            margin-bottom: 16px !important;
            border: none !important;
        }
        [data-testid="stSidebar"] button:hover {
            background-color: #27408B !important;
        }
        </style>
        """, unsafe_allow_html=True)
        # Estado dos menus
        if 'show_menu' not in st.session_state:
            st.session_state['show_menu'] = False
        if 'show_cmdk_menu' not in st.session_state:
            st.session_state['show_cmdk_menu'] = False
        # Botão TASY
        tasy_clicked = st.button("TASY", key="tasy_btn", help="Show TASY options", use_container_width=True)
        if tasy_clicked:
            st.session_state['show_menu'] = True
            st.session_state['show_cmdk_menu'] = False
        # Botão CMDK
        cmdk_clicked = st.button("CMDK", key="cmdk_btn", help="Show CMDK options", use_container_width=True)
        if cmdk_clicked:
            st.session_state['show_cmdk_menu'] = True
            st.session_state['show_menu'] = False
        # Botão Java/Delphi
        java_delphi_clicked = st.button("Java/Delphi", key="java_delphi_btn", help="Show Java/Delphi options", use_container_width=True)
        if java_delphi_clicked:
            st.session_state['show_cmdk_menu'] = False
            st.session_state['show_menu'] = False

        # Menu TASY
        if st.session_state['show_menu']:
            menu = st.radio(
                "Menu",
                options=["**✅Verification**", "**☑️Validation**"],
                index=0
            )
            if menu == "**✅Verification**":
                verification_option = st.selectbox(
                    "Choose the analysis to run:",
                    [
                        "Select a option",
                        "PRS DOC x Requirements TM Comparison",
                        "PRS DOC Comparison",
                        "Product Verification Plan",
                        "Verification TM Requirements x Test Protocol Comparison",
                        "Verification TM APP x Test Protocol Comparison",
                        "Verification Test Protocol Revision Comparison",
                        "Verification Test Protocol x Records Comparison",
                        "Verification Test Protocol x TSVR Comparison",
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
            st.markdown("<br><br><br><br><br><br><br><br><br>", unsafe_allow_html=True)
        
         # Menu Java/Delphi
        if not st.session_state['show_cmdk_menu'] and not st.session_state['show_menu']:
            java_delphi_option = st.selectbox(
                "Choose the Java/Delphi analysis to run:",
                [
                    "Select a option",
                    "Java Delphi PRS DOC x Requirements TM Comparison",
                    "Java Delphi TM Requirements x Test Protocol Comparison",
                    "Java/Delphi Option B"
                ],
                key="java_delphi_option"
            )
            st.session_state["selected_analysis"] = ("java_delphi", java_delphi_option)
            st.markdown("<br><br><br><br><br><br><br><br><br>", unsafe_allow_html=True)

        # Menu CMDK
        if st.session_state['show_cmdk_menu']:
            menu_cmdk = st.radio(
                "Menu CMDK",
                options=["**🔵Verification**", "**🟢Validation**"],
                index=0
            )
            if menu_cmdk == "**🔵Verification**":
                verification_cmdk_option = st.selectbox(
                    "Choose the CMDK analysis to run:",
                    [
                        "Select a option",
                        "CMDK PRS DOC x Requirements TM Comparison",
                        "CMDK Verification TM requirements x Test Protocol Comparison",
                        "CMDK Verification Test Protocol x TSVR Comparison",
                        "CMDK Verification Test Protocol x Records Comparison",
                        "CMDK Records x Evidences Comparison"                       
                    ],
                    key="verification_cmdk_option"
                )
                st.session_state["selected_analysis"] = ("cmdk_verification", verification_cmdk_option)
            if menu_cmdk == "**🟢Validation**":
                validation_cmdk_option = st.selectbox(
                    "Choose the CMDK analysis to run:",
                    [
                        "Select a option",
                        "CMDK TM Validation x Test Protocol Comparison",
                        "CMDK Option B"
                    ],
                    key="validation_cmdk_option"
                )
                st.session_state["selected_analysis"] = ("cmdk_validation", validation_cmdk_option)
            st.markdown("<br><br><br><br><br><br><br><br><br>", unsafe_allow_html=True)

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

