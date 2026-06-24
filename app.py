"""
XLcomparator — Application Streamlit de comparaison de fichiers Excel.

Fonctionnalités :
- Upload d'un fichier de référence (xls / xlsx)
- Upload d'un fichier à comparer (xls / xlsx)
- Options : colonne clé, normalisation des types, casse, espaces
- Résumé par feuille et tableau interactif des différences
- Téléchargement des résultats en XML, DOCX ou XLSX
"""

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

from comparator import (
    CompareOptions,
    compare_workbooks,
    differences_to_dataframe,
    differences_to_docx,
    differences_to_summary,
    differences_to_xlsx,
    differences_to_xml,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="XLcomparator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.4rem;
        font-weight: 700;
        color: #2E4057;
        text-align: center;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #6c757d;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.2rem;
        font-weight: 600;
        color: #2E4057;
        margin-bottom: 0.5rem;
    }
    .diff-count-ok {
        font-size: 1.1rem;
        color: #198754;
        font-weight: 600;
    }
    .diff-count-warn {
        font-size: 1.1rem;
        color: #dc3545;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
h_left, h_center, h_right = st.columns([2, 5, 2])

with h_left:
    st.image("logo.png", width=130)

with h_center:
    st.markdown(
        '<p class="main-title">📊 XLcomparator</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="sub-title">Comparez deux fichiers Excel'
        ' et exportez les différences</p>',
        unsafe_allow_html=True,
    )

with h_right:
    st.image("LOL.jpg", width=130)

st.divider()

# ---------------------------------------------------------------------------
# File upload section
# ---------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.markdown(
        '<p class="section-header">📁 Fichier de référence</p>',
        unsafe_allow_html=True,
    )
    ref_file = st.file_uploader(
        "Choisissez le fichier de référence",
        type=["xls", "xlsx"],
        key="ref_upload",
        label_visibility="collapsed",
    )
    if ref_file:
        st.success(
            f"✔ **{ref_file.name}** chargé ({ref_file.size:,} octets)"
        )

with col2:
    st.markdown(
        '<p class="section-header">📂 Fichier à comparer</p>',
        unsafe_allow_html=True,
    )
    cmp_file = st.file_uploader(
        "Choisissez le fichier à comparer",
        type=["xls", "xlsx"],
        key="cmp_upload",
        label_visibility="collapsed",
    )
    if cmp_file:
        st.success(
            f"✔ **{cmp_file.name}** chargé ({cmp_file.size:,} octets)"
        )

st.divider()

# ---------------------------------------------------------------------------
# Options de comparaison
# ---------------------------------------------------------------------------
with st.expander("⚙️ Options de comparaison", expanded=False):
    opt_col1, opt_col2, opt_col3 = st.columns(3)

    with opt_col1:
        normalize_types = st.checkbox(
            "Normaliser les types",
            value=True,
            help=(
                'Évite les faux positifs : "1.0" = "1", '
                '"01/01/2024" = "2024-01-01".'
            ),
        )
    with opt_col2:
        ignore_case = st.checkbox(
            "Ignorer la casse",
            value=False,
            help='Traite "ABC" et "abc" comme identiques.',
        )
    with opt_col3:
        ignore_whitespace = st.checkbox(
            "Ignorer les espaces",
            value=False,
            help="Ignore les espaces en début/fin de cellule.",
        )

    st.markdown("---")

    key_col: str | None = None
    if ref_file:
        try:
            ref_file.seek(0)
            peek = pd.read_excel(
                ref_file, header=None, nrows=1, dtype=str
            )
            ref_file.seek(0)
            first_row = (
                peek.iloc[0].fillna("").tolist()
                if not peek.empty else []
            )
        except Exception:
            ref_file.seek(0)
            first_row = []

        col_opts = ["— Aucune (comparaison positionnelle) —"] + [
            f"{get_column_letter(i + 1)} — {v}"
            for i, v in enumerate(first_row)
        ]
        key_choice = st.selectbox(
            "Colonne clé (optionnel)",
            col_opts,
            help=(
                "Aligne les lignes par valeur unique avant "
                "comparaison (ex. N° de commande, ID élève). "
                "Laissez vide pour une comparaison positionnelle."
            ),
        )
        if key_choice and not key_choice.startswith("—"):
            key_col = key_choice.split(" — ")[0].strip()
    else:
        st.caption(
            "Chargez le fichier de référence pour activer "
            "la sélection d'une colonne clé."
        )

# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------
if ref_file and cmp_file:
    if st.button(
        "🔍 Lancer la comparaison",
        use_container_width=True,
        type="primary",
    ):
        with st.spinner("Comparaison en cours…"):
            try:
                opts = CompareOptions(
                    key_col=key_col,
                    ignore_case=ignore_case,
                    ignore_whitespace=ignore_whitespace,
                    normalize_types=normalize_types,
                )
                result = compare_workbooks(ref_file, cmp_file, opts)
                st.session_state["result"] = result
                st.session_state["opts"] = opts
                st.session_state["ref_name"] = ref_file.name
                st.session_state["cmp_name"] = cmp_file.name
            except Exception as exc:
                st.error(f"Erreur lors de la comparaison : {exc}")
                st.session_state.pop("result", None)
elif ref_file or cmp_file:
    st.info(
        "⬆ Veuillez charger les **deux fichiers**"
        " pour lancer la comparaison."
    )
else:
    st.info(
        "⬆ Chargez un fichier de référence"
        " et un fichier à comparer pour commencer."
    )

# ---------------------------------------------------------------------------
# Results display
# ---------------------------------------------------------------------------
if "result" in st.session_state:
    result = st.session_state["result"]
    opts: CompareOptions = st.session_state.get(
        "opts", CompareOptions()
    )
    differences = result.differences
    ref_name = st.session_state.get("ref_name", "référence")
    cmp_name = st.session_state.get("cmp_name", "comparé")

    st.subheader("📋 Résultats de la comparaison")

    # --- Summary table ---
    summary_df = differences_to_summary(result)
    st.markdown("**Résumé par feuille**")
    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    if not differences:
        st.markdown(
            '<p class="diff-count-ok">✅ Les deux fichiers sont'
            ' identiques — aucune différence trouvée.</p>',
            unsafe_allow_html=True,
        )
    else:
        n = len(differences)
        st.markdown(
            f'<p class="diff-count-warn">⚠️ {n}'
            f' différence(s) trouvée(s) entre'
            f' <b>{ref_name}</b> et <b>{cmp_name}</b></p>',
            unsafe_allow_html=True,
        )

        df = differences_to_dataframe(differences)
        ligne_label = (
            f"Clé (col. {opts.key_col})"
            if opts.key_col else "Ligne"
        )
        df.columns = [
            "Feuille", ligne_label, "Colonne",
            "Valeur référence", "Valeur comparée",
        ]

        # Optional sheet filter
        sheets = sorted(df["Feuille"].unique().tolist())
        if len(sheets) > 1:
            selected_sheets = st.multiselect(
                "Filtrer par feuille",
                options=sheets,
                default=sheets,
            )
            df_display = df[df["Feuille"].isin(selected_sheets)]
        else:
            df_display = df

        st.dataframe(
            df_display, use_container_width=True, hide_index=True
        )

        # -------------------------------------------------------------------
        # Download section
        # -------------------------------------------------------------------
        st.divider()
        st.subheader("⬇️ Télécharger les résultats")

        dl_col1, dl_col2, dl_col3 = st.columns(3)

        with dl_col1:
            xml_bytes = differences_to_xml(differences)
            st.download_button(
                label="📄 Télécharger en XML",
                data=xml_bytes,
                file_name="differences.xml",
                mime="application/xml",
                use_container_width=True,
            )
            st.caption("Format XML pour injection en base de données")

        with dl_col2:
            docx_bytes = differences_to_docx(differences)
            st.download_button(
                label="📝 Télécharger en DOCX",
                data=docx_bytes,
                file_name="differences.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                ),
                use_container_width=True,
            )
            st.caption("Rapport Word (.docx)")

        with dl_col3:
            xlsx_bytes = differences_to_xlsx(differences)
            st.download_button(
                label="📊 Télécharger en XLSX",
                data=xlsx_bytes,
                file_name="differences.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
            st.caption("Tableau Excel (.xlsx)")

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.divider()
st.markdown(
    "<p style='text-align:center; color:#6c757d; font-size:0.85rem;'>"
    "XLcomparator — Comparez, analysez, exportez"
    " vos fichiers Excel en toute simplicité."
    "</p>",
    unsafe_allow_html=True,
)
