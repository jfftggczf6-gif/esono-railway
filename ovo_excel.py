"""
ESONO OVO Excel Generator — Complete cell mapping verified against actual template
251022-PlanFinancierOVO-Template5Ans-v0210-EMPTY.xlsm

CRITICAL RULES (verified by reading the template with openpyxl):
- Col S in FinanceData = ALWAYS FORMULA (=Q+R) — NEVER WRITE
- Col R in RevenueData = FORMULA (=1-P-Q) — NEVER WRITE
- Col Z,AA,AB in RevenueData = FORMULA (=1-W, =1-X, =1-Y) — NEVER WRITE
- Col AI in RevenueData rows 9-12 = FORMULA (=sum quarters), rows 13-16 = VALUE
- Row 216+ (social security) T-X = FORMULA — only write O,P,Q,R
- Row 773 (cash) = O,P only, S+ = FORMULA
- J25-J33 in InputsData = FORMULA cascade — only write J24
- Percentages as decimals (0.18 not 18)
- Never modify col C (VLOOKUP keys)
- Never add/delete rows
- Save as .xlsm with keep_vba=True
"""

import io
import base64
from typing import Optional
from datetime import datetime
from fastapi import HTTPException, Header, Request
from fastapi.responses import Response
import openpyxl


# Product block = 16 rows (8 volume + 8 revenue calc)
PRODUCT_BLOCK = 16
PRODUCT_VOL_START = 9   # First product volume row (PRODUCT 01, YEAR-2)
SERVICE_VOL_START = 329  # First service volume row (after 20 products × 16 = 320 + 9)

# Staff rows in FinanceData
STAFF = {
    0: {"eft":213,"sal":214,"allow":215,"ss":216},
    1: {"eft":220,"sal":221,"allow":222,"ss":223},
    2: {"eft":227,"sal":228,"allow":229,"ss":230},
    3: {"eft":234,"sal":235,"allow":236,"ss":237},
    4: {"eft":241,"sal":242,"allow":243,"ss":244},
    5: {"eft":248,"sal":249,"allow":250,"ss":251},
    6: {"eft":255,"sal":256,"allow":257,"ss":258},
    7: {"eft":262,"sal":263,"allow":264,"ss":265},
    8: {"eft":269,"sal":270,"allow":271,"ss":272},
    9: {"eft":276,"sal":277,"allow":278,"ss":279},
}

CAPEX = {
    "office":     {"start":408,"count":7,"amort":0.25},
    "production": {"start":450,"count":7,"amort":0.10},
    "other":      {"start":462,"count":7,"amort":0.10},
    "intangible": {"start":486,"count":5,"amort":0.20},
    "startup":    {"start":493,"count":5,"amort":0.20},
}

# Year key → FinanceData column
YR_COL = {
    "year_minus_2":"O","year_minus_1":"P","h1":"Q","h2":"R",
    "current_year":"S","year2":"T","year3":"U","year4":"V","year5":"W","year6":"X",
}

# Safe columns (skip S which is always formula)
SAFE_COLS = {"year_minus_2":"O","year_minus_1":"P","h1":"Q","h2":"R",
             "year2":"T","year3":"U","year4":"V","year5":"W","year6":"X"}

# Revenue year offsets within a product block (0-7)
REV_OFFSETS = {"year_minus_2":0,"year_minus_1":1,"current_year":2,"year2":4,"year3":5,"year4":6,"year5":7}

# Deno expander uses these keys in per_year arrays
DENO_YEAR_MAP = {
    "YEAR-2": "year_minus_2", "YEAR-1": "year_minus_1",
    "CURRENT YEAR": "current_year", "CURRENT YEAR H1": "h1", "CURRENT YEAR H2": "h2",
    "YEAR2": "year2", "YEAR3": "year3", "YEAR4": "year4", "YEAR5": "year5", "YEAR6": "year6",
}
# Index in 10-value arrays → year key
IDX_YEAR = ["year_minus_2","year_minus_1","h1","h2","current_year","year2","year3","year4","year5","year6"]


def normalize_per_year(per_year):
    """Convert per_year from Deno array format to dict format.
    Deno: [{year: "YEAR-2", unit_price_r1: 1500, ...}, ...]
    Python expects: {"year_minus_2": {unit_price_r1: 1500, ...}, ...}
    Also handles already-dict format."""
    if isinstance(per_year, dict):
        return per_year  # Already in dict format
    if isinstance(per_year, list):
        result = {}
        for entry in per_year:
            if not isinstance(entry, dict):
                continue
            year_label = entry.get("year", "")
            year_key = DENO_YEAR_MAP.get(year_label, year_label.lower().replace("-", "_minus_").replace(" ", "_"))
            result[year_key] = entry
        return result
    return {}


def normalize_finance_values(val):
    """Convert finance values from various formats to dict {year_key: number}.
    Handles: 10-element array, dict with Deno keys, dict with Python keys, or single number."""
    if isinstance(val, list) and len(val) >= 8:
        # Array of 10 values [O,P,Q,R,S,T,U,V,W,X] = [ym2,ym1,h1,h2,cy,y2,y3,y4,y5,y6]
        result = {}
        for i, v in enumerate(val[:10]):
            if i < len(IDX_YEAR):
                result[IDX_YEAR[i]] = v or 0
        return result
    if isinstance(val, dict):
        # Check if Deno keys (YEAR-2, etc.) and convert
        result = {}
        for k, v in val.items():
            mapped = DENO_YEAR_MAP.get(k, k)
            result[mapped] = v
        return result
    if isinstance(val, (int, float)):
        return {"current_year": val}
    return {}


def sw(ws, row, col, value):
    """Safe write: skip if cell is a formula."""
    cell = ws[f"{col}{row}"]
    if isinstance(cell.value, str) and cell.value.startswith("="):
        return
    cell.value = value


def wfr(ws, row, d, skip_s=True):
    """Write finance row from dict or array format."""
    d = normalize_finance_values(d)
    if not isinstance(d, dict):
        return
    for yk, col in (SAFE_COLS if skip_s else YR_COL).items():
        v = d.get(yk)
        if v is not None:
            sw(ws, row, col, v)


def fill_ovo(wb, data):
    # ── Safety: recursively ensure all nested objects are dicts where expected ──
    def safe_dict(val):
        """Convert to dict if not already. Single-element list [{...}] → unwrap."""
        if isinstance(val, dict):
            return val
        if isinstance(val, list):
            if len(val) == 1 and isinstance(val[0], dict):
                return val[0]
            if len(val) == 0:
                return {}
            # List of non-dict items — can't convert
            return {}
        return {}

    def sanitize(obj):
        """Recursively sanitize: any dict value that should be a dict but is a list gets converted."""
        if not isinstance(obj, dict):
            return obj
        result = {}
        for k, v in obj.items():
            if isinstance(v, dict):
                result[k] = sanitize(v)
            elif isinstance(v, list):
                # If list of dicts (like products, services, staff, capex items) → keep as list but sanitize each
                if v and isinstance(v[0], dict):
                    result[k] = [sanitize(item) if isinstance(item, dict) else item for item in v]
                else:
                    result[k] = v
            else:
                result[k] = v
        return result

    # Sanitize top-level keys that must be dicts
    for key in ["years", "ranges", "channels", "loans", "scenarios", "opex",
                "cogs", "startup_costs_detail", "funding", "metadata",
                "working_capital", "capex"]:
        if key in data and not isinstance(data[key], dict):
            print(f"[fill_ovo] WARNING: '{key}' is {type(data[key]).__name__}, converting to dict")
            data[key] = safe_dict(data[key])

    # Deep sanitize the entire opex tree (most complex nesting)
    if "opex" in data and isinstance(data["opex"], dict):
        data["opex"] = sanitize(data["opex"])

    # Deep sanitize loans, scenarios, working_capital
    for key in ["loans", "scenarios", "working_capital"]:
        if key in data and isinstance(data[key], dict):
            data[key] = sanitize(data[key])

    # Special handling for capex: AI might send a flat list instead of categorized dict
    capex_raw = data.get("capex", {})
    if isinstance(capex_raw, list):
        print(f"[fill_ovo] WARNING: 'capex' is a list ({len(capex_raw)} items), converting to categorized dict")
        categorized = {"office": [], "production": [], "other": [], "intangible": [], "startup": []}
        for item in capex_raw:
            if not isinstance(item, dict):
                continue
            cat = item.get("category", item.get("type", "other")).lower()
            if cat in categorized:
                categorized[cat].append(item)
            else:
                categorized["other"].append(item)
        data["capex"] = categorized
    elif isinstance(capex_raw, dict):
        data["capex"] = sanitize(capex_raw)

    inp = wb["InputsData"]
    rev = wb["RevenueData"]
    fin = wb["FinanceData"]
    rdm = wb["ReadMe"]

    # ── ReadMe ──
    sw(rdm, 3, "L", "French")

    # ── InputsData Section 1: Company ──
    sw(inp, 5, "J", data.get("company",""))
    sw(inp, 6, "J", data.get("country","CÔTE D'IVOIRE"))
    sw(inp, 8, "J", data.get("currency","XOF"))
    sw(inp, 9, "J", data.get("exchange_rate_eur",655.957))
    sw(inp, 10,"J", datetime.now())
    sw(inp, 12,"J", data.get("vat_rate",0.18))
    sw(inp, 14,"J", data.get("inflation_rate",0.03))
    sw(inp, 17,"J", data.get("tax_regime_1",0.04))
    sw(inp, 18,"J", data.get("tax_regime_2",0.25))

    # ── Section 2: Years (ONLY J24) ──
    sw(inp, 24, "J", data.get("years",{}).get("year_minus_2",2022))

    # ── Section 3: Products (36-55) ──
    prods = data.get("products",[])
    for i in range(20):
        r = 36+i
        if i < len(prods) and prods[i].get("active",True):
            sw(inp,r,"H",prods[i].get("name","-"))
            sw(inp,r,"I",1)
        else:
            sw(inp,r,"H","-")
            sw(inp,r,"I",0)

    # ── Section 3: Services (58-67) ──
    svcs = data.get("services",[])
    for i in range(10):
        r = 58+i
        if i < len(svcs) and svcs[i].get("active",True):
            sw(inp,r,"H",svcs[i].get("name","-"))
            sw(inp,r,"I",1)
        else:
            sw(inp,r,"H","-")
            sw(inp,r,"I",0)

    # ── Section 4: Ranges & Channels ──
    rng = data.get("ranges",{})
    sw(inp,70,"J",rng.get("range1","Entry level"))
    sw(inp,71,"J",rng.get("range2","Advanced level"))
    sw(inp,72,"J",rng.get("range3","Professional level"))
    ch = data.get("channels",{})
    sw(inp,75,"J",ch.get("channel1","Business to Business"))
    sw(inp,76,"J",ch.get("channel2","Business to Consumer"))

    # ── Product × Range/Channel matrix (79-98) ──
    for i,p in enumerate(prods[:20]):
        r = 79+i
        rf = p.get("range_flags",[1,0,0])
        cf = p.get("channel_flags",[1,0])
        sw(inp,r,"F",rf[0] if len(rf)>0 else 1)
        sw(inp,r,"G",rf[1] if len(rf)>1 else 0)
        sw(inp,r,"H",rf[2] if len(rf)>2 else 0)
        sw(inp,r,"I",cf[0] if len(cf)>0 else 1)
        sw(inp,r,"J",cf[1] if len(cf)>1 else 0)

    for i,s in enumerate(svcs[:10]):
        r = 101+i
        rf = s.get("range_flags",[1,0,0])
        cf = s.get("channel_flags",[1,0])
        sw(inp,r,"F",rf[0] if len(rf)>0 else 1)
        sw(inp,r,"G",rf[1] if len(rf)>1 else 0)
        sw(inp,r,"H",rf[2] if len(rf)>2 else 0)
        sw(inp,r,"I",cf[0] if len(cf)>0 else 1)
        sw(inp,r,"J",cf[1] if len(cf)>1 else 0)

    # ── Section 5: Staff (113-122) ──
    staff = data.get("staff",[])
    for i in range(10):
        r = 113+i
        if i < len(staff):
            c = staff[i]
            sw(inp,r,"H",c.get("category",""))
            sw(inp,r,"I",c.get("department",""))
            sw(inp,r,"J",c.get("social_security_rate",0.1645))
        else:
            sw(inp,r,"J",0)

    # ── Section 6: Loans (125-127) ──
    loans = data.get("loans",{})
    ovo = loans.get("ovo",{})
    sw(inp,125,"I",ovo.get("interest_rate",0.07))
    sw(inp,125,"J",ovo.get("duration_years",5))
    fam = loans.get("family",{})
    sw(inp,126,"I",fam.get("interest_rate",0.10))
    sw(inp,126,"J",fam.get("duration_years",3))
    bnk = loans.get("bank",{})
    sw(inp,127,"I",bnk.get("interest_rate",0.20))
    sw(inp,127,"J",bnk.get("duration_years",2))

    # ── Section 7: Scenarios (130-142) — only if provided ──
    scen = data.get("scenarios",{})
    SCEN_ROWS = {"revenue_products":130,"cogs_products":131,"revenue_services":132,
        "cogs_services":133,"marketing":134,"staff_salaries":135,"taxes_staff":136,
        "office":137,"other_expenses":138,"travel":139,"insurance":140,
        "maintenance":141,"third_parties":142}
    for k,r in SCEN_ROWS.items():
        if k in scen:
            s = scen[k]
            if s.get("worst") is not None: sw(inp,r,"H",s["worst"])
            if s.get("typical") is not None: sw(inp,r,"I",s["typical"])
            if s.get("best") is not None: sw(inp,r,"J",s["best"])

    # ═══ RevenueData: Products ═══
    def fill_rev_item(base_row, item):
        py = normalize_per_year(item.get("per_year", {}))
        active = item.get("active", True)
        for yk, off in REV_OFFSETS.items():
            r = base_row + off
            yr = py.get(yk,{})
            if not active:
                for c in ["L","M","N","P","Q","S","T","U","W","X","Y","AE","AF","AG","AH"]:
                    sw(rev,r,c,0)
                if off >= 4: sw(rev,r,"AI",0)
                continue
            # Prices
            sw(rev,r,"L",yr.get("unit_price_r1",0))
            sw(rev,r,"M",yr.get("unit_price_r2",0))
            sw(rev,r,"N",yr.get("unit_price_r3",0))
            # Mix volume (P,Q only — R=formula)
            sw(rev,r,"P",yr.get("mix_r1",1.0))
            sw(rev,r,"Q",yr.get("mix_r2",0))
            # COGS
            sw(rev,r,"S",yr.get("cogs_r1",0))
            sw(rev,r,"T",yr.get("cogs_r2",0))
            sw(rev,r,"U",yr.get("cogs_r3",0))
            # Channel mix (W,X,Y only — Z,AA,AB=formula)
            sw(rev,r,"W",yr.get("mix_r1_ch1",1.0))
            sw(rev,r,"X",yr.get("mix_r2_ch1",1.0))
            sw(rev,r,"Y",yr.get("mix_r3_ch1",1.0))
            # Volumes
            if off <= 2:
                sw(rev,r,"AE",round(yr.get("q1",yr.get("volume_q1",yr.get("volume_h1",0)))))
                sw(rev,r,"AF",round(yr.get("q2",yr.get("volume_q2",yr.get("volume_h2",0)))))
                sw(rev,r,"AG",round(yr.get("q3",yr.get("volume_q3",0))))
                sw(rev,r,"AH",round(yr.get("q4",yr.get("volume_q4",0))))
            else:
                sw(rev,r,"AI",round(yr.get("total",yr.get("volume_total",0))))

    for i,p in enumerate(prods[:20]):
        fill_rev_item(PRODUCT_VOL_START + i*PRODUCT_BLOCK, p)
    for i,s in enumerate(svcs[:10]):
        fill_rev_item(SERVICE_VOL_START + i*PRODUCT_BLOCK, s)

    # ═══ FinanceData: Stock Variation (188-189) ═══
    sv = data.get("stock_variation",{})
    wfr(fin,188,sv.get("increase",{}))
    wfr(fin,189,sv.get("decrease",{}))

    # ═══ OPEX Marketing (201-210) ═══
    opex = data.get("opex",{})
    mkt = opex.get("marketing",{})
    for k,r in {"research":201,"studies":202,"receptions":203,"documentation":204,"advertising":205}.items():
        wfr(fin,r,mkt.get(k,{}))
    for i,x in enumerate(mkt.get("extra",[])[:5]):
        wfr(fin,206+i,x)

    # ═══ Staff in FinanceData ═══
    for i,c in enumerate(staff[:10]):
        rows = STAFF.get(i)
        if not rows: continue
        py = normalize_per_year(c.get("per_year",{}))
        for yk,col in SAFE_COLS.items():
            yr = py.get(yk,{})
            if not isinstance(yr, dict): continue
            sw(fin,rows["eft"],col,round(yr.get("headcount",0)))
            sw(fin,rows["sal"],col,yr.get("gross_salary",yr.get("gross_monthly_salary_per_person",0)))
            sw(fin,rows["allow"],col,yr.get("allowances",yr.get("annual_allowances_per_person",0)))
        # Social security: O,P,Q,R ONLY (T-X = formula)
        for yk in ["year_minus_2","year_minus_1","h1","h2"]:
            col = SAFE_COLS.get(yk)
            if col:
                yr = py.get(yk,{})
                if not isinstance(yr, dict): continue
                sw(fin,rows["ss"],col,yr.get("social_security",0))

    # ═══ Taxes on Staff (283-291) ═══
    tx = opex.get("taxes_staff",{})
    wfr(fin,283,tx.get("salaries_tax",{}))
    wfr(fin,284,tx.get("apprenticeship",{}))
    wfr(fin,285,tx.get("training",{}))
    wfr(fin,286,tx.get("other",{}))
    for i,x in enumerate(tx.get("extra",[])[:5]): wfr(fin,287+i,x)

    # ═══ Office (294-308) ═══
    of = opex.get("office",{})
    for k,r in {"rent":294,"internet":295,"telecom":296,"supplies":297,"equipment":298,
                "packaging":299,"fuel":300,"water":301,"electricity":302,"cleaning":303}.items():
        wfr(fin,r,of.get(k,{}))
    for i,x in enumerate(of.get("extra",[])[:5]): wfr(fin,304+i,x)

    # ═══ Other (311-319) ═══
    ot = opex.get("other",{})
    for k,r in {"health":311,"directors":312,"donations":313,"sponsorship":314}.items():
        wfr(fin,r,ot.get(k,{}))
    for i,x in enumerate(ot.get("extra",[])[:5]): wfr(fin,315+i,x)

    # ═══ Travel (322-323) ═══
    tr = opex.get("travel",{})
    wfr(fin,322,tr.get("headcount",{}))
    wfr(fin,323,tr.get("cost_per_person",{}))

    # ═══ Insurance (326-332) ═══
    ins = opex.get("insurance",{})
    wfr(fin,326,ins.get("building",{}))
    wfr(fin,327,ins.get("company",{}))
    for i,x in enumerate(ins.get("extra",[])[:5]): wfr(fin,328+i,x)

    # ═══ Maintenance (335-342) ═══
    mn = opex.get("maintenance",{})
    wfr(fin,335,mn.get("movable",{}))
    wfr(fin,336,mn.get("demolition",{}))
    wfr(fin,337,mn.get("other",{}))
    for i,x in enumerate(mn.get("extra",[])[:5]): wfr(fin,338+i,x)

    # ═══ Third Parties (345-357) ═══
    tp = opex.get("third_parties",{})
    for k,r in {"legal_startup":345,"loan_followup":346,"other_startup":347,
                "transport":348,"delivery":349,"commissions":350,
                "legal_adviser":351,"accounting":352}.items():
        wfr(fin,r,tp.get(k,{}))
    for i,x in enumerate(tp.get("extra",[])[:5]): wfr(fin,353+i,x)

    # ═══ CAPEX (408-497) ═══
    capex = data.get("capex",{})
    for ck,cc in CAPEX.items():
        for i,item in enumerate(capex.get(ck,[])[:cc["count"]]):
            r = cc["start"]+i
            sw(fin,r,"J",item.get("label",""))
            sw(fin,r,"K",item.get("year",0))
            sw(fin,r,"L",item.get("value",0))
            sw(fin,r,"M",item.get("amort_rate",cc["amort"]))

    # ═══ Working Capital (693,697,701) — days only ═══
    wc = data.get("working_capital",{})
    wfr(fin,693,wc.get("stock_days",{}))
    wfr(fin,697,wc.get("receivable_days",{}))
    wfr(fin,701,wc.get("payable_days",{}))

    # ═══ P&L Adjustments (724-741) ═══
    adj = data.get("adjustments",{})
    wfr(fin,724,adj.get("inventory_change",{}))
    wfr(fin,725,adj.get("receivables_change",{}))
    wfr(fin,726,adj.get("debts_change",{}))
    # Bank charges O,P only
    bc = adj.get("bank_charges",{})
    sw(fin,729,"O",bc.get("year_minus_2",0))
    sw(fin,729,"P",bc.get("year_minus_1",0))
    # Historical interest O,P only
    for rw,lk in [(730,"interest_ovo"),(731,"interest_family"),(732,"interest_bank")]:
        d = adj.get(lk,{})
        sw(fin,rw,"O",d.get("year_minus_2",0))
        sw(fin,rw,"P",d.get("year_minus_1",0))
    # Tax regimes
    taxes = adj.get("taxes",{})
    wfr(fin,737,taxes.get("regime_1",{}))
    wfr(fin,738,taxes.get("regime_2",{}))

    # ═══ Cash Position (773) — O,P only ═══
    cash = data.get("cash_position",{})
    sw(fin,773,"O",cash.get("year_minus_2",0))
    sw(fin,773,"P",cash.get("year_minus_1",0))

    # ═══ Loans (785-816) ═══
    LC = {"current_year":"S","year2":"T","year3":"U","year4":"V","year5":"W","year6":"X"}
    ovo_amt = ovo.get("amount_by_year",{})
    fam_amt = fam.get("amount_by_year",{})
    bnk_amt = bnk.get("amount_by_year",{})
    for yk,col in LC.items():
        sw(fin,785,col,ovo_amt.get(yk,0))
        sw(fin,786,col,fam_amt.get(yk,0))
        sw(fin,787,col,bnk_amt.get(yk,0))

    # Shareholder contributions
    fing = data.get("financing",{})
    ex_sh = fing.get("existing_shareholders",{})
    nw_sh = fing.get("new_shareholders",{})
    dt_cn = fing.get("debt_cancellation",{})
    ALL_LC = {"year_minus_2":"O","year_minus_1":"P",**LC}
    for yk,col in ALL_LC.items(): sw(fin,788,col,ex_sh.get(yk,0))
    for yk,col in LC.items(): sw(fin,789,col,nw_sh.get(yk,0))
    PLC = {"year_minus_1":"P",**LC}
    for yk,col in PLC.items(): sw(fin,790,col,dt_cn.get(yk,0))

    # Grace periods
    og = ovo.get("grace_by_year",{})
    fg = fam.get("grace_by_year",{})
    bg = bnk.get("grace_by_year",{})
    for yk,col in LC.items():
        sw(fin,793,col,og.get(yk,0))
        sw(fin,802,col,fg.get(yk,0))
        sw(fin,811,col,bg.get(yk,0))

    # Interest H1
    sw(fin,798,"S",ovo.get("interest_h1_current_year",0))
    sw(fin,807,"S",fam.get("interest_h1_current_year",0))
    sw(fin,816,"S",bnk.get("interest_h1_current_year",0))

    return wb


def register_excel_endpoints(app):
    @app.post("/generate-ovo-excel")
    async def generate_ovo_excel(request: Request, authorization: Optional[str] = Header(None)):
        import os
        PK = os.getenv("PARSER_API_KEY","esono-parser-dev-key")
        if authorization != f"Bearer {PK}":
            raise HTTPException(status_code=401, detail="Invalid API key")

        body = await request.json()
        data = body.get("data",{})
        t64 = body.get("template_base64")
        if not t64:
            raise HTTPException(status_code=400, detail="template_base64 required")

        try:
            tb = base64.b64decode(t64)
            wb = openpyxl.load_workbook(io.BytesIO(tb), keep_vba=True, data_only=False)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid template: {e}")

        np = len(data.get("products",[]))
        ns = len(data.get("services",[]))
        print(f"[ovo-excel] {np} products, {ns} services")

        try:
            wb = fill_ovo(wb, data)
        except Exception as e:
            import traceback
            print(f"[ovo-excel] Error: {traceback.format_exc()}")
            raise HTTPException(status_code=500, detail=f"Fill error: {e}")

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        print(f"[ovo-excel] Done: {out.getbuffer().nbytes} bytes")

        return Response(content=out.read(),
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            headers={"Content-Disposition":"attachment; filename=PlanFinancierOVO.xlsm"})
