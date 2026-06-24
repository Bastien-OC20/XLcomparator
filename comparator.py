"""
Comparison engine for XLcomparator.
Compares two Excel workbooks cell-by-cell and returns a list of differences.
"""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import List, Dict, Any

import pandas as pd
from openpyxl.utils import get_column_letter


def _read_workbook(file_obj) -> Dict[str, pd.DataFrame]:
    """Read all sheets of an Excel file into a dict of DataFrames."""
    file_obj.seek(0)
    xl = pd.ExcelFile(file_obj)
    sheets: Dict[str, pd.DataFrame] = {}
    for name in xl.sheet_names:
        df = xl.parse(name, header=None, dtype=str)
        df = df.fillna("")
        df.columns = list(range(len(df.columns)))
        sheets[str(name)] = df
    return sheets


def compare_workbooks(ref_file, cmp_file) -> List[Dict[str, Any]]:
    """
    Compare two Excel workbooks and return a list of differences.

    Each difference is a dict with keys:
        feuille      : sheet name
        ligne        : 1-based row number
        colonne      : Excel-style column letter (A, B, …)
        valeur_ref   : value in the reference file
        valeur_cmp   : value in the file being compared
    """
    ref_sheets = _read_workbook(ref_file)
    cmp_sheets = _read_workbook(cmp_file)

    differences: List[Dict[str, Any]] = []

    all_sheets = sorted(set(list(ref_sheets.keys()) + list(cmp_sheets.keys())))

    for sheet in all_sheets:
        ref_df = ref_sheets.get(sheet, pd.DataFrame())
        cmp_df = cmp_sheets.get(sheet, pd.DataFrame())

        # Align to the same shape
        max_rows = max(len(ref_df), len(cmp_df))
        max_cols = max(
            ref_df.shape[1] if not ref_df.empty else 0,
            cmp_df.shape[1] if not cmp_df.empty else 0,
        )

        ref_df = ref_df.reindex(
            index=range(max_rows), columns=range(max_cols), fill_value=""
        )
        cmp_df = cmp_df.reindex(
            index=range(max_rows), columns=range(max_cols), fill_value=""
        )

        for row_idx in range(max_rows):
            for col_idx in range(max_cols):
                ref_val = str(ref_df.iloc[row_idx, col_idx])
                cmp_val = str(cmp_df.iloc[row_idx, col_idx])

                if ref_val != cmp_val:
                    differences.append(
                        {
                            "feuille": sheet,
                            "ligne": row_idx + 1,
                            "colonne": get_column_letter(col_idx + 1),
                            "valeur_ref": ref_val,
                            "valeur_cmp": cmp_val,
                        }
                    )

    return differences


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------


def differences_to_dataframe(
    differences: List[Dict[str, Any]]
) -> pd.DataFrame:
    """Convert the list of differences to a pandas DataFrame."""
    if not differences:
        return pd.DataFrame(
            columns=["feuille", "ligne", "colonne", "valeur_ref", "valeur_cmp"]
        )
    return pd.DataFrame(differences)


def differences_to_xml(differences: List[Dict[str, Any]]) -> bytes:
    """Serialise differences to an XML byte string for DB injection."""
    root = ET.Element("differences")

    for diff in differences:
        elem = ET.SubElement(root, "difference")
        ET.SubElement(elem, "feuille").text = str(diff["feuille"])
        ET.SubElement(elem, "ligne").text = str(diff["ligne"])
        ET.SubElement(elem, "colonne").text = str(diff["colonne"])
        ET.SubElement(elem, "valeur_ref").text = str(diff["valeur_ref"])
        ET.SubElement(elem, "valeur_cmp").text = str(diff["valeur_cmp"])

    raw = ET.tostring(root, encoding="unicode")
    pretty = minidom.parseString(raw).toprettyxml(
        indent="  ", encoding="UTF-8"
    )
    return pretty


def differences_to_xlsx(differences: List[Dict[str, Any]]) -> bytes:
    """Serialise differences to an XLSX byte string."""
    df = differences_to_dataframe(differences)
    df.columns = [
        "Feuille", "Ligne", "Colonne",
        "Valeur référence", "Valeur comparée",
    ]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Différences")

        # Auto-fit column widths
        worksheet = writer.sheets["Différences"]
        for col in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max_length + 4

    return buf.getvalue()


def differences_to_docx(differences: List[Dict[str, Any]]) -> bytes:
    """Serialise differences to a DOCX byte string."""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # Title
    title = doc.add_heading("Rapport de comparaison XLcomparator", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not differences:
        doc.add_paragraph(
            "Aucune différence trouvée entre les deux fichiers."
        )
    else:
        doc.add_paragraph(f"Nombre total de différences : {len(differences)}")
        doc.add_paragraph("")

        headers = [
            "Feuille", "Ligne", "Colonne",
            "Valeur référence", "Valeur comparée",
        ]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"

        # Header row
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            run = hdr_cells[i].paragraphs[0].runs[0]
            run.bold = True
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            para_fmt = hdr_cells[i].paragraphs[0].paragraph_format
            para_fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Dark background via XML shading
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
            tc = hdr_cells[i]._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "2E4057")
            tcPr.append(shd)

        # Data rows
        for diff in differences:
            row_cells = table.add_row().cells
            row_cells[0].text = str(diff["feuille"])
            row_cells[1].text = str(diff["ligne"])
            row_cells[2].text = str(diff["colonne"])
            row_cells[3].text = str(diff["valeur_ref"])
            row_cells[4].text = str(diff["valeur_cmp"])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
