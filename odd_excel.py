"""
ESONO ODD Excel Generator — Fills the OVO ODD template
ODD_template.xlsx

Sheets to fill:
1. "Evaluation cibles ODD avec Circ" — E1 (project name), F/G/H/I per cible (0/1), M/N (questions)
2. "INDICATEURS d'IMPACT" — H/I per cible (conditional), J (focus), K (baseline), M (target)

NEVER write to cols J, K, L in Evaluation sheet (formulas).
NEVER write to Objectif header rows (row 9, 12, 16, etc.) — only cible rows.
"""

import io
import base64
from typing import Optional
from fastapi import HTTPException, Header, Request
from fastapi.responses import Response
import openpyxl


# ═══════════════════════════════════════════════════════════════
# CIBLE ROW MAPPING — Exact rows for each SDG target in the template
# Key = cible ID as string, Value = row number
# ═══════════════════════════════════════════════════════════════

# Evaluation sheet cible rows
EVAL_CIBLE_ROWS = {
    "1.1": 10, "1.4": 11,
    "2.1": 13, "2.4": 14, "2.a": 15,
    "3.3": 17, "3.9": 18,
    "4.1": 20,
    "5.1": 22,
    "6.3": 24, "6.4": 25, "6.a": 26,
    "7.1": 28, "7.2a": 29, "7,2 a": 29, "7.2 a": 29,
    "7.2b": 30, "7,2 b": 30, "7.2 b": 30,
    "7.3": 31,
    "8.2": 33, "8.3": 34, "8.5": 35, "8.6": 36, "8.8": 37, "8.9": 38,
    "9.3": 40, "9.4": 41, "9.b": 42, "9.c": 43,
    "10.2": 45,
    "11.6": 47, "11.c": 48,
    "12.2": 50, "12.5": 51,
    "13.3": 53,
    "14.1": 55,
    "15.2": 58, "15.3": 59,
    "16.5": 62, "16.7": 63,
    "17.1": 66, "17.3": 67, "17.16": 69,
}

# Indicators sheet cible rows (slightly different numbering)
INDIC_CIBLE_ROWS = {
    "1.1": 11, "1.4": 12,
    "2.1": 14, "2.4": 15, "2.a": 16,
    "3.3": 18, "3.9": 19,
    "4.1": 21,
    "5.1": 23,
    "6.3": 25, "6.4": 26, "6.a": 27,
    "7.1": 29, "7.2a": 30, "7,2 a": 30, "7.2 a": 30,
    "7.2b": 31, "7,2 b": 31, "7.2 b": 31,
    "7.3": 32,
    "8.2": 34, "8.3": 35, "8.5": 36, "8.6": 37, "8.8": 38, "8.9": 39,
    "9.3": 41, "9.4": 42, "9.b": 43, "9.c": 44,
    "10.2": 46,
    "11.6": 48, "11.c": 49,
    "12.2": 51, "12.5": 52,
    "13.3": 54,
    "14.1": 56,
    "15.2": 58, "15.3": 59,
    "16.5": 61, "16.7": 62,
    "17.1": 65, "17.3": 66, "17.16": 68,
}


def sw(ws, row, col, value):
    """Safe write: skip formulas."""
    cell = ws[f"{col}{row}"]
    if isinstance(cell.value, str) and cell.value.startswith("="):
        return
    cell.value = value


def normalize_cible_id(cible_id):
    """Normalize cible ID to match template format."""
    s = str(cible_id).strip().lower()
    # Try exact match first, then common variations
    return s


def find_row(cible_id, row_map):
    """Find row for a cible ID, trying multiple formats."""
    cid = str(cible_id).strip()
    # Direct match
    if cid in row_map:
        return row_map[cid]
    # Try lowercase
    if cid.lower() in row_map:
        return row_map[cid.lower()]
    # Try without spaces
    no_space = cid.replace(" ", "")
    if no_space in row_map:
        return row_map[no_space]
    # Try with period instead of comma
    dot_ver = cid.replace(",", ".")
    if dot_ver in row_map:
        return row_map[dot_ver]
    return None


def fill_odd(wb, data):
    """Fill the ODD template with evaluation data."""

    eval_ws = wb["Evaluation cibles ODD avec Circ"]
    indic_ws = wb["INDICATEURS d'IMPACT"]

    # ── Project name (E1) ──
    project_name = data.get("project_name", data.get("nom_entreprise", ""))
    if project_name:
        sw(eval_ws, 1, "E", project_name)

    # ── Evaluation cibles ──
    cibles = data.get("cibles", [])
    for cible in cibles:
        cible_id = cible.get("cible_id", cible.get("cible", cible.get("cible_number", "")))
        row = find_row(cible_id, EVAL_CIBLE_ROWS)
        if not row:
            print(f"[odd-excel] Warning: cible '{cible_id}' not found in template")
            continue

        # F = positif (1/0), G = neutre (1/0), H = négatif (1/0), I = besoin aide (1/0)
        positif = cible.get("positif", cible.get("positive", 0))
        neutre = cible.get("neutre", cible.get("neutral", 0))
        negatif = cible.get("negatif", cible.get("negative", 0))
        aide = cible.get("besoin_aide", cible.get("need_help", 0))

        # Auto-derive from alignment string if individual flags not provided
        alignment = str(cible.get("alignement", cible.get("alignment", ""))).lower()
        if not any([positif, neutre, negatif]):
            if "positif" in alignment or "positive" in alignment or "fort" in alignment or "direct" in alignment:
                positif = 1
            elif "négatif" in alignment or "negative" in alignment or "negatif" in alignment:
                negatif = 1
            elif "neutre" in alignment or "neutral" in alignment or "indirect" in alignment:
                neutre = 1

        sw(eval_ws, row, "F", positif or 0)
        sw(eval_ws, row, "G", neutre or 0)
        sw(eval_ws, row, "H", negatif or 0)
        sw(eval_ws, row, "I", aide or 0)

        # M = question 1, N = question 2
        q1 = cible.get("question_1", cible.get("justification", ""))
        q2 = cible.get("question_2", cible.get("recommandation", cible.get("action", "")))
        if q1:
            sw(eval_ws, row, "M", str(q1)[:1000])
        if q2:
            sw(eval_ws, row, "N", str(q2)[:1000])

    # ── Indicateurs d'impact ──
    indicateurs = data.get("indicateurs", [])
    for ind in indicateurs:
        cible_id = ind.get("cible_id", ind.get("cible", ""))
        row = find_row(cible_id, INDIC_CIBLE_ROWS)
        if not row:
            continue

        # H = positif envers cible (conditional formatting in template)
        # I = à améliorer pendant le projet
        # J = focus OVO (1 or empty)
        # K = baseline au début du projet
        # M = cible à la fin du projet

        focus = ind.get("focus_ovo", ind.get("focus", None))
        if focus is not None:
            sw(indic_ws, row, "J", 1 if focus else None)

        baseline = ind.get("baseline", ind.get("valeur_actuelle", ""))
        if baseline:
            sw(indic_ws, row, "K", str(baseline))

        target = ind.get("target", ind.get("cible_fin", ind.get("objectif", "")))
        if target:
            sw(indic_ws, row, "M", str(target))

    # ── Date in INDICATEURS sheet (K7) ──
    date_str = data.get("date", "")
    if date_str:
        sw(indic_ws, 7, "K", f"DATE: {date_str}")

    return wb


def register_odd_endpoints(app):
    """Register ODD Excel generation endpoint."""

    @app.post("/generate-odd-excel")
    async def generate_odd_excel(request: Request, authorization: Optional[str] = Header(None)):
        import os
        PK = os.getenv("PARSER_API_KEY", "esono-parser-dev-key")
        if authorization != f"Bearer {PK}":
            raise HTTPException(status_code=401, detail="Invalid API key")

        body = await request.json()
        data = body.get("data", {})
        t64 = body.get("template_base64")

        if not t64:
            raise HTTPException(status_code=400, detail="template_base64 required")

        try:
            tb = base64.b64decode(t64)
            wb = openpyxl.load_workbook(io.BytesIO(tb), data_only=False)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid template: {e}")

        nc = len(data.get("cibles", []))
        ni = len(data.get("indicateurs", []))
        print(f"[odd-excel] {nc} cibles, {ni} indicateurs")

        try:
            wb = fill_odd(wb, data)
        except Exception as e:
            import traceback
            print(f"[odd-excel] Error: {traceback.format_exc()}")
            raise HTTPException(status_code=500, detail=f"Fill error: {e}")

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        print(f"[odd-excel] Done: {out.getbuffer().nbytes} bytes")

        return Response(
            content=out.read(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Analyse_ODD.xlsx"}
        )
