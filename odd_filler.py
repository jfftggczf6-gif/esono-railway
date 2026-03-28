"""
odd_filler.py — Remplissage du template ODD depuis le JSON odd_analysis.

Source: odd_analysis JSON (evaluation_cibles_odd, indicateurs_impact, synthese, etc.)
Target: ODD template.xlsx
"""

import openpyxl

# ═══════════════════════════════════════════════════════════════
# MAPPING: target_id → row dans "Evaluation cibles ODD avec Circ"
# ═══════════════════════════════════════════════════════════════

EVAL_TARGET_ROWS = {
    "1.1": 10, "1.4": 11,
    "2.1": 13, "2.4": 14, "2.a": 15,
    "3.3": 17, "3.9": 18,
    "4.1": 20,
    "5.1": 22,
    "6.3": 24, "6.4": 25, "6.a": 26,
    "7.1": 28, "7.2a": 29, "7,2 a": 29, "7.2b": 30, "7,2 b": 30, "7.3": 31,
    "8.2": 33, "8.3": 34, "8.5": 35, "8.6": 36, "8.8": 37, "8.9": 38,
    "9.3": 40, "9.4": 41, "9.b": 42, "9.c": 43,
    "10.2": 45,
    "11.6": 47, "11.c": 48,
    "12.2": 50, "12.5": 51,
    "13.3": 53,
    "14.1": 55,
    "15.2": 58, "15.3": 59,
    "16.5": 62, "16.7": 63,
    "17.1": 66, "17.3": 67, "17.16": 69, "17.17": 70,
}

# MAPPING: target_id → row dans "INDICATEURS d'IMPACT"
INDIC_TARGET_ROWS = {
    "1.1": 11, "1.4": 12,
    "2.1": 14, "2.4": 15, "2.a": 16,
    "3.3": 18, "3.9": 19,
    "4.1": 21,
    "5.1": 23,
    "6.3": 25, "6.4": 26, "6.a": 27,
    "7.1": 29, "7.2a": 30, "7,2 a": 30, "7.2b": 31, "7,2 b": 31, "7.3": 32,
    "8.2": 34, "8.3": 35, "8.5": 36, "8.6": 37, "8.8": 38, "8.9": 39,
    "9.3": 41, "9.4": 42, "9.b": 43, "9.c": 44,
    "10.2": 46,
    "11.6": 48, "11.c": 49,
    "12.2": 51, "12.5": 52,
    "13.3": 54,
    "14.1": 56,
    "15.2": 58, "15.3": 59,
    "16.5": 61, "16.7": 62,
    "17.1": 65, "17.3": 66, "17.16": 68, "17.17": 69,
}

# ODD headers → row dans "Evaluation cibles ODD avec Circ"
ODD_HEADER_ROWS = {
    "1": 9, "2": 12, "3": 16, "4": 19, "5": 21, "6": 23, "7": 27,
    "8": 32, "9": 39, "10": 44, "11": 46, "12": 49, "13": 52,
    "14": 54, "15": 56, "16": 60, "17": 64,
}

# Aperçu sheet: ODD number → row (rows 5-21)
APERCU_ODD_ROWS = {str(i): 4 + i for i in range(1, 18)}


def fw(ws, row, col, value):
    """Force write."""
    ws[f"{col}{row}"].value = value


def fill_odd(wb, data):
    """
    Remplit le workbook ODD à partir du JSON odd_analysis.
    """
    eval_ws = wb["Evaluation cibles ODD avec Circ"]
    indic_ws = wb["INDICATEURS d'IMPACT"]
    apercu_ws = wb["Aperçu"]

    # ═══ 1. En-tête: nom projet + pays ═══
    info = data.get("informations_projet", data.get("metadata", {}))
    nom = info.get("nom_entreprise", data.get("metadata", {}).get("nom_entreprise", ""))
    pays = info.get("pays", data.get("metadata", {}).get("pays", ""))
    fw(eval_ws, 1, "E", f"{nom} en {pays}" if nom and pays else nom or "")

    # Aussi dans INDICATEURS
    fw(indic_ws, 2, "C", f"{nom} — {pays}" if nom else "")

    # ═══ 2. Evaluation des cibles ODD ═══
    cibles = data.get("evaluation_cibles_odd", {}).get("cibles", [])

    for cible in cibles:
        target_id = str(cible.get("target_id", "")).strip()
        evaluation = cible.get("evaluation", "").lower()
        justification = cible.get("justification", "")
        info_add = cible.get("info_additionnelle", "")

        row = EVAL_TARGET_ROWS.get(target_id)
        if not row:
            # Essayer sans le point (ex: "2a" → "2.a")
            for key, r in EVAL_TARGET_ROWS.items():
                if key.replace(".", "").replace(",", "").replace(" ", "") == target_id.replace(".", "").replace(",", "").replace(" ", ""):
                    row = r
                    break
        if not row:
            print(f"[odd_filler] Target {target_id} not found in template mapping")
            continue

        # Colonnes F=positif, G=neutre, H=négatif, I=besoin d'aide
        fw(eval_ws, row, "F", 1 if evaluation == "positif" else None)
        fw(eval_ws, row, "G", 1 if evaluation == "neutre" else None)
        fw(eval_ws, row, "H", 1 if evaluation in ("negatif", "négatif", "risque") else None)
        fw(eval_ws, row, "I", 1 if evaluation in ("inconnu", "besoin_aide", "???") else None)

        # Justification dans M, info additionnelle dans N
        if justification:
            fw(eval_ws, row, "M", justification[:500])
        if info_add:
            fw(eval_ws, row, "N", info_add[:500])

    print(f"[odd_filler] {len(cibles)} cibles évaluées")

    # ═══ 3. Indicateurs d'impact ═══
    indicateurs = data.get("indicateurs_impact", {}).get("indicateurs", [])

    for indic in indicateurs:
        target_id = str(indic.get("target_id", "")).strip()
        row = INDIC_TARGET_ROWS.get(target_id)
        if not row:
            for key, r in INDIC_TARGET_ROWS.items():
                if key.replace(".", "").replace(",", "").replace(" ", "") == target_id.replace(".", "").replace(",", "").replace(" ", ""):
                    row = r
                    break
        if not row:
            continue

        valeur = indic.get("valeur", "")
        indic_ovo = indic.get("indicateur_ovo", "")
        cible_val = indic.get("cible", indic.get("target", ""))

        # H = positif (vert), I = à améliorer (orange)
        # Chercher l'évaluation correspondante dans les cibles
        eval_match = next((c for c in cibles if str(c.get("target_id", "")).strip() == target_id), None)
        if eval_match:
            evaluation = eval_match.get("evaluation", "").lower()
            if evaluation == "positif":
                fw(indic_ws, row, "H", 1)
                fw(indic_ws, row, "I", None)
            elif evaluation in ("neutre", "negatif", "négatif", "risque"):
                fw(indic_ws, row, "H", None)
                fw(indic_ws, row, "I", 1)

        # J = focus OVO (1 si indicateur important)
        if indic.get("focus") or indic.get("prioritaire"):
            fw(indic_ws, row, "J", 1)

        # K = baseline
        if valeur:
            fw(indic_ws, row, "K", str(valeur)[:200])

        # M = cible
        if cible_val:
            fw(indic_ws, row, "M", str(cible_val)[:200])

    print(f"[odd_filler] {len(indicateurs)} indicateurs remplis")

    # ═══ 4. Aperçu ═══
    synthese = data.get("synthese", {})
    odd_prioritaires = synthese.get("odd_prioritaires", [])

    # Collecter les ODD positifs, négatifs et besoin d'aide
    odd_positifs = {}
    odd_negatifs = {}
    odd_aide = {}

    for cible in cibles:
        odd_num = str(cible.get("odd_parent", "")).strip()
        if not odd_num:
            # Extraire depuis target_id (ex: "8.3" → "8")
            tid = str(cible.get("target_id", ""))
            odd_num = tid.split(".")[0].split(",")[0].strip()

        evaluation = cible.get("evaluation", "").lower()
        target_id = str(cible.get("target_id", ""))

        if evaluation == "positif":
            odd_positifs.setdefault(odd_num, []).append(target_id)
        elif evaluation in ("negatif", "négatif", "risque"):
            odd_negatifs.setdefault(odd_num, []).append(target_id)
        elif evaluation in ("inconnu", "besoin_aide", "???"):
            odd_aide.setdefault(odd_num, []).append(target_id)

    # Remplir l'aperçu
    for odd_num in range(1, 18):
        odd_str = str(odd_num)
        row = APERCU_ODD_ROWS.get(odd_str)
        if not row:
            continue

        # D = contribution positive
        positifs = odd_positifs.get(odd_str, [])
        if positifs:
            fw(apercu_ws, row, "D", f"Cibles: {', '.join(positifs)}")

        # F = contribution négative
        negatifs = odd_negatifs.get(odd_str, [])
        if negatifs:
            fw(apercu_ws, row, "F", f"Risques: {', '.join(negatifs)}")

        # H = besoin d'aide
        aide = odd_aide.get(odd_str, [])
        if aide:
            fw(apercu_ws, row, "H", f"À clarifier: {', '.join(aide)}")

    # Contribution globale dans la cellule A2
    contrib = synthese.get("contribution_globale", "")
    if contrib:
        fw(apercu_ws, 2, "A", contrib[:500])

    print(f"[odd_filler] Aperçu: {len(odd_positifs)} ODD positifs, {len(odd_negatifs)} ODD négatifs")

    return wb


def register_odd_endpoints(app):
    """Register the ODD Excel endpoint in FastAPI."""
    import io
    import base64
    from typing import Optional
    from fastapi import HTTPException, Header, Request
    from fastapi.responses import Response

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
            wb = openpyxl.load_workbook(io.BytesIO(tb))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid template: {e}")

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

        return Response(content=out.read(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=ODD_Evaluation.xlsx"})
