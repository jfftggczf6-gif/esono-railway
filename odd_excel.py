"""
ESONO ODD Excel Generator v2 — Uses odd_filler.py for actual filling.
Endpoint: POST /generate-odd-excel
"""

import io
import base64
from typing import Optional
from fastapi import HTTPException, Header, Request
from fastapi.responses import Response
import openpyxl

from odd_filler import fill_odd


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
            wb = openpyxl.load_workbook(io.BytesIO(tb))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid template: {e}")

        nc = len(data.get("evaluation_cibles_odd", {}).get("cibles", []))
        ni = len(data.get("indicateurs_impact", {}).get("indicateurs", []))
        print(f"[odd-excel] v2 filler: {nc} cibles, {ni} indicateurs")

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
