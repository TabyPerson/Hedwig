import streamlit as st
import streamlit as st
import pandas as pd
import io
import re


def get_sheet_with_fallback(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return preferred_name
    else:
        st.warning(f"Worksheet named '{preferred_name}' not found. Please select the correct sheet.")
        return st.selectbox("Select the correct worksheet:", xl.sheet_names, key=preferred_name)


def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r"\s+", " ", x.strip()))


def force_to_strings(item):
    if item is None:
        return ""
    if isinstance(item, pd.DataFrame):
        try:
            if len(item) == 0:
                return pd.DataFrame({"Nenhum dado": ["Nenhum dado encontrado"]})
            result_df = item.copy()
            for col in result_df.columns:
                try:
                    result_df[col] = result_df[col].apply(lambda x: "" if pd.isna(x) else str(x).strip())
                except Exception:
                    try:
                        result_df[col] = result_df[col].fillna("").astype(str)
                    except Exception:
                        result_df[col] = ["[Erro de Conversão]"] * len(result_df)
            return result_df
        except Exception:
            return pd.DataFrame({"Erro": ["Erro ao processar DataFrame"]})
    if isinstance(item, (list, tuple, set)):
        result = []
        for subitem in item:
            try:
                if subitem is None or pd.isna(subitem):
                    continue
                text = str(subitem).strip()
                if text and text.lower() not in ["nan", "none", "null", ""]:
                    result.append(text)
            except Exception:
                pass
        return result
    try:
        if pd.isna(item):
            return ""
        return str(item).strip()
    except Exception:
        return ""


# Utilitários para PRS
def get_prs_from_csv(df, col):
    try:
        if col in df.columns:
            return df[col].dropna().astype(str).str.strip().tolist()
    except Exception:
        pass
    return []


def get_prs_from_xls(xl, col):
    prs_ids = []
    try:
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            if col in df.columns:
                prs_ids.extend(df[col].dropna().astype(str).str.strip().tolist())
    except Exception:
        pass
    return prs_ids


def get_prs_from_sheet(xl, sheet_name, prs_col):
    try:
        if sheet_name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            if prs_col in df.columns:
                return df[prs_col].dropna().astype(str).str.strip().tolist()
    except Exception:
        pass
    return []


def display_prs_section(title, prs_list, description):
    total = len(prs_list)
    st.markdown(f"**{title}**")
    st.write(f"({total}) {description}")
    cols = st.columns(3)
    col_len = (len(prs_list) + 2) // 3
    split_lists = [prs_list[i * col_len:(i + 1) * col_len] for i in range(3)]
    for idx, col in enumerate(cols):
        with col:
            st.write(split_lists[idx] if split_lists[idx] else "-")


def run_comparison():
    st.markdown("### Product Verification Plan")
    st.info(
        "Upload your Product Requirement (.xlsx), ADO Impacted Spreadsheet (.csv), PRSTS Spreadsheet (.xlsx) e Requirements Traceability Matrix (.xlsx)."
    )

    cols = st.columns(4)
    with cols[0]:
        prs_file = st.file_uploader("Upload Product Requirement (.xlsx)", type=["xlsx"], key="dvp_prs")
    with cols[1]:
        ado_file = st.file_uploader("Upload ADO Impacted Spreadsheet (.csv)", type=["csv"], key="dvp_ADO")
    with cols[2]:
        prsts_file = st.file_uploader("Upload PRSTS Spreadsheet (.xlsx)", type=["xlsx"], key="dvp_PRSTS")
    with cols[3]:
        rtm_file = st.file_uploader(
            "Upload Requirements Traceability Matrix (.xlsx)", type=["xlsx"], key="dvp_RTM"
        )

    if prs_file and ado_file and prsts_file:
        if st.button("🔍 Run Comparison", key="run_comparison"):
            # carregar arquivos
            try:
                prs_xl = pd.ExcelFile(prs_file)
                ado_df = pd.read_csv(ado_file)
                prsts_xl = pd.ExcelFile(prsts_file)
            except Exception as e:
                st.error(f"Erro ao carregar arquivos: {e}")
                return

            ado_prs = get_prs_from_csv(ado_df, "Requirements")
            prsts_prs = get_prs_from_xls(prsts_xl, "Cd prs id")

            def prs_from_product(sheet, prs_col="PRS ID"):
                return get_prs_from_sheet(prs_xl, sheet, prs_col)

            def intersect_prs(prs_a, prs_b, prs_sheet):
                return sorted(set(prs_a + prs_b) & set(prs_sheet))

            def prs_dataframe_3cols(prs_list):
                prs_list = sorted(set(str(prs).strip() for prs in prs_list if prs and str(prs).strip()))
                n = len(prs_list)
                if n == 0:
                    return pd.DataFrame({"PRS 1": [], "PRS 2": [], "PRS 3": []})
                col_len = (n + 2) // 3
                cols_data = [prs_list[i * col_len:(i + 1) * col_len] for i in range(3)]
                max_len = max(len(c) for c in cols_data)
                for i in range(3):
                    cols_data[i] += [""] * (max_len - len(cols_data[i]))
                df = pd.DataFrame({f"PRS {i+1}": cols_data[i] for i in range(3)})
                return df

            export_dfs = {}
            export_texts = {}

            # Sessões e textos
            sessions = [
                (
                    "Labeling and Learning Materials",
                    "Labeling and Learning Materials Requirements",
                    "PRS's listed below will be retested to ensure that labeling specifications are met for this release",
                ),
                ("Interface Requirements", "Interface Requirements", "PRS´s listed below were impacted and for this reason are part of the scope of this release."),
                (
                    "Risk Management Matrix",
                    "Safety Mitigation Requirements (part of the Risk Management Matrix)",
                    "PRS's listed below were impacted and for this reason are part of the scope of this release.",
                ),
                ("Security and Privacy", "Security and Privacy Requirements", "PRS's listed below were impacted and for this reason are part of the scope of this release."),
                ("Cloud Design", "Cloud Requirements", "PRS's listed below were impacted and for this reason are part of the scope of this release."),
                ("AI Requirements", "AI Requirements", "PRS's listed below were impacted and for this reason are part of the scope of this release."),
            ]

            # RTM - read Design Verification sheet (collect PRS marked as Not verified)
            rtm_not_verified = []
            if rtm_file:
                try:
                    rtm_xl = pd.ExcelFile(rtm_file)
                    # prefer a sheet named exactly 'Design Verification' if present
                    rtm_sheet = "Design Verification" if "Design Verification" in rtm_xl.sheet_names else rtm_xl.sheet_names[0]
                    rtm_df_full = pd.read_excel(rtm_xl, sheet_name=rtm_sheet)
                    if "Verification Test Result (Pass/Fail)" in rtm_df_full.columns and "PRS Requirement ID" in rtm_df_full.columns:
                        rtm_not_verified = rtm_df_full.loc[
                            rtm_df_full["Verification Test Result (Pass/Fail)"].astype(str).str.strip().str.lower() == "not verified",
                            "PRS Requirement ID",
                        ].dropna().astype(str).str.strip().tolist()
                    else:
                        rtm_not_verified = []
                except Exception as e:
                    st.warning(f"Erro ao processar RTM (Design Verification): {e}")

            all_special_prs = set()
            for sheet, tab_name, text in sessions:
                prs = prs_from_product(sheet)
                prs_found = intersect_prs(ado_prs, prsts_prs, prs)
                df = prs_dataframe_3cols(prs_found)
                export_dfs[tab_name] = df
                # Adiciona o total de PRSs antes das frases específicas
                if (
                    "have test cases that have not been tested and for this reason are part of the scope of this release." in text
                    or "were impacted and for this reason are part of the scope of this release." in text
                    or "will be retested to ensure that labeling specifications are met for this release" in text
                ):
                    text = f"({len(prs_found)}) {text}"
                export_texts[tab_name] = text
                all_special_prs.update(prs)
                st.markdown(f"**{tab_name}**")
                st.markdown(text)
                st.dataframe(df, use_container_width=True)

                # Se houver PRSs marcados como 'Not verified' na aba Design Verification do RTM,
                # verifica se esses PRS pertencem a essa sessão e exibe os PRS em 3 colunas com o texto pedido.
                try:
                    rtm_prs_in_session = sorted(set(rtm_not_verified) & set(prs)) if rtm_not_verified else []
                    if rtm_prs_in_session:
                        rtm_df_3cols = prs_dataframe_3cols(rtm_prs_in_session)
                        rtm_text_session = f"({len(rtm_prs_in_session)}) PRS´s listed below and that are part of the scope have test cases that have not been and for this reason are part of the scope of this release"
                        export_dfs[f"{tab_name} - RTM Not Verified"] = rtm_df_3cols
                        export_texts[f"{tab_name} - RTM Not Verified"] = rtm_text_session
                        st.markdown(rtm_text_session)
                        st.dataframe(rtm_df_3cols, use_container_width=True)
                except Exception:
                    # não falha a pipeline por causa do RTM
                    pass

            # Functional Requirements
            functional_prs = prs_from_product("Functional Requirements")
            functional_prs_found = sorted((set(ado_prs + prsts_prs) & set(functional_prs)) - all_special_prs)
            functional_df = prs_dataframe_3cols(functional_prs_found)
            export_dfs["Functional Requirements"] = functional_df
            functional_text = f"({len(functional_prs_found)}) PRS´s listed below were impacted and for this reason are part of the scope of this release."
            export_texts["Functional Requirements"] = functional_text
            st.markdown("**Functional Requirements**")
            st.markdown(functional_text)
            st.dataframe(functional_df, use_container_width=True)

            # Incluir PRSs do RTM (Design Verification) que estejam na aba Functional Requirements
            # e que não pertençam a outras sessões especiais
            try:
                rtm_prs_in_functional = []
                if rtm_not_verified:
                    # pega PRSs do RTM que estão na aba Functional Requirements
                    rtm_prs_in_functional = sorted((set(rtm_not_verified) & set(functional_prs)) - all_special_prs - set(functional_prs_found))
                if rtm_prs_in_functional:
                    rtm_df_3cols = prs_dataframe_3cols(rtm_prs_in_functional)
                    rtm_text_session = f"({len(rtm_prs_in_functional)}) PRS´s listed below and that are part of the scope have test cases that have not been and for this reason are part of the scope of this release"
                    export_dfs["Functional Requirements - RTM Not Verified"] = rtm_df_3cols
                    export_texts["Functional Requirements - RTM Not Verified"] = rtm_text_session
                    st.markdown(rtm_text_session)
                    st.dataframe(rtm_df_3cols, use_container_width=True)
            except Exception:
                pass

            # RTM already integrated into session displays above (Design Verification sheet)

            # Exportar Excel com todas as abas
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                for sheet_name, df in export_dfs.items():
                    text = export_texts.get(sheet_name, "")
                    # escreve texto na primeira linha
                    txt_df = pd.DataFrame({sheet_name: [text]})
                    txt_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=0)
                    # escreve tabela a partir da linha 3
                    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
            buffer.seek(0)
            st.download_button(
                label="⬇️ Download Comparisson (Excel)",
                data=buffer.getvalue(),
                file_name="prs_por_sessao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
