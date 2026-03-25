"""
plan_financier_adapter.py — Adapte le nouveau format unifié (plan_financier)
au format attendu par fill_ovo() existant.

Le fill_ovo() a un mapping cellule-par-cellule qui fonctionne.
Plutôt que le réécrire, on transforme le nouveau JSON en ancien JSON.

BUGS FIXES v2:
- B1: Years J25-J33 → fill_ovo only writes J24, rest are formulas ✅
- B2: Product names "-" → use real nom from produits[] ✅
- B3: Staff names empty → category/department from staff[] ✅
- B4: Staff data empty → headcount/salary from par_annee ✅
- B5: RevenueData cogs_r3/mix/volumes empty → fill all fields ✅
- B6: Prices aberrant → use real prix_r1 values ✅
- B7: CAPEX empty → group by category, map keys ✅
- B8: Loans amounts 0 → build amount_by_year from amount/term ✅
- B9: Loan durations empty → pass duration_years ✅
- B10: Scenarios only Worst → build worst/typical/best ✅
"""

import math

# New per_year label → old year key
YEAR_LABEL_MAP = {
    "YEAR-2": "year_minus_2",
    "YEAR-1": "year_minus_1",
    "CURRENT YEAR": "current_year",
    "YEAR2": "year2",
    "YEAR3": "year3",
    "YEAR4": "year4",
    "YEAR5": "year5",
    "YEAR6": "year6",
}

# Array index → year key (for 10-element OPEX arrays)
IDX_TO_YEAR = [
    "year_minus_2", "year_minus_1", "h1", "h2", "current_year",
    "year2", "year3", "year4", "year5", "year6",
]


def _array_to_dict(arr):
    """Convert a 10-element array to year-keyed dict."""
    if not isinstance(arr, list):
        return arr if isinstance(arr, dict) else {}
    result = {}
    for i, v in enumerate(arr[:10]):
        if i < len(IDX_TO_YEAR):
            result[IDX_TO_YEAR[i]] = v or 0
    return result


def _build_loan_amount_by_year(amount, term_years):
    """Build yearly repayment dict from total amount and duration."""
    if not amount or amount <= 0:
        return {}
    term = max(1, int(term_years or 5))
    yearly = round(amount / term)
    year_keys = ["current_year", "year2", "year3", "year4", "year5", "year6"]
    return {year_keys[i]: yearly for i in range(min(term, len(year_keys)))}


def _aggregate_products_into_slot1(products):
    """
    Aggregate up to 3 products into Product 1 using R1/R2/R3 ranges.
    - R1 = product 0, R2 = product 1, R3 = product 2
    - Total volume = sum of all product volumes
    - Mix R1/R2/R3 = share of each product's CA in total CA
    - Price R1/R2/R3 = each product's unit price
    - COGS R1/R2/R3 = each product's unit COGS
    """
    prods = products[:3]  # Max 3 ranges
    n = len(prods)

    # Build per_year by aggregating across years
    all_years = set()
    for p in prods:
        for pa in p.get("par_annee", []):
            label = pa.get("annee", "")
            yk = YEAR_LABEL_MAP.get(label)
            if yk:
                all_years.add(yk)

    per_year = {}
    for yk in all_years:
        # Collect data for each product at this year
        prod_data = []
        for p in prods:
            pa = next((a for a in p.get("par_annee", []) if YEAR_LABEL_MAP.get(a.get("annee")) == yk), None)
            if pa:
                prod_data.append(pa)
            else:
                prod_data.append({})

        # Total volume = sum
        total_vol = sum(pd.get("volume", 0) or 0 for pd in prod_data)
        total_ca = sum((pd.get("prix_r1", 0) or 0) * (pd.get("volume", 0) or 0) for pd in prod_data)

        # Mix = CA share of each product
        mix = [0.0] * 3
        for i, pd in enumerate(prod_data):
            ca_i = (pd.get("prix_r1", 0) or 0) * (pd.get("volume", 0) or 0)
            mix[i] = ca_i / total_ca if total_ca > 0 else (1.0 if i == 0 else 0.0)

        # Quarters — aggregate
        q1 = sum(pd.get("volume_q1", 0) or 0 for pd in prod_data)
        q2 = sum(pd.get("volume_q2", 0) or 0 for pd in prod_data)
        q3 = sum(pd.get("volume_q3", 0) or 0 for pd in prod_data)
        q4 = sum(pd.get("volume_q4", 0) or 0 for pd in prod_data)
        if q1 + q2 + q3 + q4 == 0 and total_vol > 0:
            q1 = round(total_vol * 0.22)
            q2 = round(total_vol * 0.25)
            q3 = round(total_vol * 0.27)
            q4 = total_vol - q1 - q2 - q3

        year_entry = {
            "unit_price_r1": prod_data[0].get("prix_r1", 0) if n > 0 else 0,
            "unit_price_r2": prod_data[1].get("prix_r1", 0) if n > 1 else 0,
            "unit_price_r3": prod_data[2].get("prix_r1", 0) if n > 2 else 0,
            "mix_r1": round(mix[0], 4),
            "mix_r2": round(mix[1], 4) if n > 1 else 0,
            "mix_r3": round(mix[2], 4) if n > 2 else 0,
            "cogs_r1": prod_data[0].get("cogs_r1", 0) if n > 0 else 0,
            "cogs_r2": prod_data[1].get("cogs_r1", 0) if n > 1 else 0,
            "cogs_r3": prod_data[2].get("cogs_r1", 0) if n > 2 else 0,
            "mix_r1_ch1": prod_data[0].get("mix_ch1", 0.5) if n > 0 else 1.0,
            "mix_r2_ch1": prod_data[1].get("mix_ch1", 0.5) if n > 1 else 1.0,
            "mix_r3_ch1": prod_data[2].get("mix_ch1", 0.5) if n > 2 else 1.0,
            "q1": q1, "volume_q1": q1,
            "q2": q2, "volume_q2": q2,
            "q3": q3, "volume_q3": q3,
            "q4": q4, "volume_q4": q4,
            "volume_h1": q1 + q2,
            "volume_h2": q3 + q4,
            "volume_total": total_vol,
            "total": total_vol,
        }
        per_year[yk] = year_entry

        # H1/H2 for current year
        if yk == "current_year":
            h1_vol = q1 + q2
            h2_vol = q3 + q4
            per_year["h1"] = {**year_entry, "q1": h1_vol, "volume_q1": h1_vol, "q2": 0, "volume_q2": 0,
                              "volume_h1": h1_vol, "volume_h2": 0, "volume_total": h1_vol, "total": h1_vol}
            per_year["h2"] = {**year_entry, "q1": h2_vol, "volume_q1": h2_vol, "q2": 0, "volume_q2": 0,
                              "volume_h1": h2_vol, "volume_h2": 0, "volume_total": h2_vol, "total": h2_vol}

    # Range flags: activate as many ranges as products
    range_flags = [1 if i < n else 0 for i in range(3)]

    return {
        "name": " + ".join(p.get("nom", "Produit") for p in prods),
        "active": True,
        "range_flags": range_flags,
        "channel_flags": [1, 1] if n > 0 else [0, 1],
        "per_year": per_year,
    }


def _convert_product(p):
    """Convert a single produit (v2) to product (v1)."""
    product = {
        "name": p.get("nom", "Produit"),
        "active": True,
        "range_flags": p.get("range_flags", [1, 0, 0]),
        "channel_flags": p.get("channel_flags", [0, 1]),
        "per_year": {},
    }

    for yr_data in p.get("par_annee", []):
        label = yr_data.get("annee", "")
        yr_key = YEAR_LABEL_MAP.get(label)
        if not yr_key:
            continue

        vol_total = yr_data.get("volume", 0) or 0
        q1 = yr_data.get("volume_q1", 0) or round(vol_total * 0.22)
        q2 = yr_data.get("volume_q2", 0) or round(vol_total * 0.25)
        q3 = yr_data.get("volume_q3", 0) or round(vol_total * 0.27)
        q4 = yr_data.get("volume_q4", 0) or (vol_total - q1 - q2 - q3)

        year_entry = {
            "unit_price_r1": yr_data.get("prix_r1", 0),
            "unit_price_r2": yr_data.get("prix_r2", 0),
            "unit_price_r3": yr_data.get("prix_r3", 0),
            "mix_r1": yr_data.get("mix_r1", 1.0),
            "mix_r2": yr_data.get("mix_r2", 0),
            "mix_r3": yr_data.get("mix_r3", 0),
            "cogs_r1": yr_data.get("cogs_r1", 0),
            "cogs_r2": yr_data.get("cogs_r2", 0),
            "cogs_r3": yr_data.get("cogs_r3", 0),
            "mix_r1_ch1": yr_data.get("mix_ch1", 1.0),
            "mix_r2_ch1": yr_data.get("mix_ch1", 1.0),
            "mix_r3_ch1": yr_data.get("mix_ch1", 1.0),
            "q1": q1, "volume_q1": q1,
            "q2": q2, "volume_q2": q2,
            "q3": q3, "volume_q3": q3,
            "q4": q4, "volume_q4": q4,
            "volume_h1": q1 + q2,
            "volume_h2": q3 + q4,
            "volume_total": vol_total,
            "total": vol_total,
        }
        product["per_year"][yr_key] = year_entry

        # H1/H2 split for current year
        if label == "CURRENT YEAR":
            h1_vol = q1 + q2
            h2_vol = q3 + q4
            product["per_year"]["h1"] = {
                **year_entry,
                "q1": h1_vol, "volume_q1": h1_vol,
                "q2": 0, "volume_q2": 0,
                "volume_h1": h1_vol, "volume_h2": 0,
                "volume_total": h1_vol, "total": h1_vol,
            }
            product["per_year"]["h2"] = {
                **year_entry,
                "q1": h2_vol, "volume_q1": h2_vol,
                "q2": 0, "volume_q2": 0,
                "volume_h1": h2_vol, "volume_h2": 0,
                "volume_total": h2_vol, "total": h2_vol,
            }

    return product


def _convert_staff(s):
    """Convert a single staff category (v2) to v1."""
    staff_item = {
        "category": s.get("categorie", "EMPLOYE(E)S"),
        "department": s.get("departement", "GENERAL"),
        "social_security_rate": s.get("taux_charges_sociales", 0.1645),
        "per_year": {},
    }

    for yr_data in s.get("par_annee", []):
        label = yr_data.get("annee", "")
        yr_key = YEAR_LABEL_MAP.get(label)
        if not yr_key:
            continue

        year_entry = {
            "headcount": yr_data.get("effectif", 0),
            "gross_salary": yr_data.get("salaire_mensuel_brut", 0),
            "gross_monthly_salary_per_person": yr_data.get("salaire_mensuel_brut", 0),
            "allowances": yr_data.get("primes_annuelles", 0),
            "annual_allowances_per_person": yr_data.get("primes_annuelles", 0),
            "social_security": yr_data.get("charges_sociales_mensuelles", 0),
        }
        staff_item["per_year"][yr_key] = year_entry

        # H1/H2 for current year
        if label == "CURRENT YEAR":
            staff_item["per_year"]["h1"] = year_entry.copy()
            staff_item["per_year"]["h2"] = year_entry.copy()

    return staff_item


def _convert_opex(opex_raw):
    """Convert OPEX: arrays → dicts, rename keys."""
    if not isinstance(opex_raw, dict):
        return {}

    opex = {}
    for category, subs in opex_raw.items():
        if not isinstance(subs, dict):
            continue

        # Rename category if needed
        cat_key = category
        if category == "taxes_on_staff":
            cat_key = "taxes_staff"

        opex[cat_key] = {}
        for sub_key, values in subs.items():
            # Rename travel sub-keys
            mapped_key = sub_key
            if category == "travel":
                if sub_key == "nb_travellers":
                    mapped_key = "headcount"
                elif sub_key == "avg_cost":
                    mapped_key = "cost_per_person"

            # Convert arrays to dicts
            if isinstance(values, list):
                opex[cat_key][mapped_key] = _array_to_dict(values)
            elif isinstance(values, dict):
                opex[cat_key][mapped_key] = values
            else:
                opex[cat_key][mapped_key] = values

    return opex


def _convert_capex(capex_raw):
    """Convert CAPEX: flat list → categorized dict."""
    if isinstance(capex_raw, dict):
        return capex_raw  # Already categorized

    if not isinstance(capex_raw, list):
        return {}

    categorized = {"office": [], "production": [], "other": [], "intangible": [], "startup": []}
    CAT_MAP = {
        "bureau": "office", "office": "office", "office_equipment": "office",
        "production": "production", "production_equipment": "production",
        "equipment": "production", "commercial": "other",
        "intangible": "intangible", "incorporel": "intangible", "logiciel": "intangible",
        "startup": "startup", "demarrage": "startup",
    }

    for item in capex_raw:
        if not isinstance(item, dict):
            continue
        raw_cat = (item.get("categorie") or item.get("category") or item.get("type") or "other").lower()
        cat = CAT_MAP.get(raw_cat, "other")

        categorized[cat].append({
            "label": item.get("label", item.get("nature", "")),
            "year": item.get("acquisition_year", item.get("annee", 0)),
            "value": item.get("acquisition_value", item.get("montant", item.get("value", 0))),
            "amort_rate": item.get("amortisation_rate", item.get("taux_amortissement", item.get("amort_rate", 0.20))),
        })

    return categorized


def _convert_scenarios(scen_raw, hyp_ia=None):
    """Convert scenarios to worst/typical/best format per category."""
    if not isinstance(scen_raw, dict):
        return {}

    # If already in correct format (has "revenue_products" with worst/typical/best)
    first_val = next(iter(scen_raw.values()), None)
    if isinstance(first_val, dict) and "worst" in first_val and "typical" in first_val:
        return scen_raw

    # Build from hypotheses IA
    hyp = hyp_ia or {}
    growth_rates = hyp.get("taux_croissance_ca", [0.15])
    avg_g = sum(growth_rates) / len(growth_rates) if growth_rates else 0.15

    return {
        "revenue_products": {"worst": round(avg_g * 0.5, 3), "typical": round(avg_g, 3), "best": round(avg_g * 1.4, 3)},
        "cogs_products": {"worst": 0.05, "typical": 0.03, "best": 0.02},
        "revenue_services": {"worst": round(avg_g * 0.5, 3), "typical": round(avg_g, 3), "best": round(avg_g * 1.4, 3)},
        "cogs_services": {"worst": 0.05, "typical": 0.03, "best": 0.02},
        "marketing": {"worst": 0.10, "typical": 0.08, "best": 0.05},
        "staff_salaries": {"worst": 0.08, "typical": 0.05, "best": 0.03},
        "taxes_staff": {"worst": 0.08, "typical": 0.05, "best": 0.03},
        "office": {"worst": 0.08, "typical": 0.05, "best": 0.03},
        "other_expenses": {"worst": 0.05, "typical": 0.03, "best": 0.02},
        "travel": {"worst": 0.08, "typical": 0.05, "best": 0.03},
        "insurance": {"worst": 0.05, "typical": 0.03, "best": 0.02},
        "maintenance": {"worst": 0.08, "typical": 0.05, "best": 0.03},
        "third_parties": {"worst": 0.10, "typical": 0.08, "best": 0.05},
    }


def adapt_plan_financier(data: dict) -> dict:
    """
    Convertit le format plan_financier unifié (v2)
    en format compatible avec fill_ovo() (v1).
    """
    # Si c'est déjà l'ancien format, retourner tel quel
    if "products" in data and "produits" not in data:
        return data
    if "produits" not in data:
        return data

    print("[adapter] Converting plan_financier v2 → v1 format")

    result = {}

    # ── Pass-through fields ──
    for key in ["company", "country", "currency", "exchange_rate_eur",
                "vat_rate", "inflation_rate", "tax_regime_1", "tax_regime_2",
                "years", "financing"]:
        if key in data:
            result[key] = data[key]

    # ── Products ──
    # Each product gets its OWN slot (Product 1, 2, 3, ...).
    # fill_ovo uses force-write (fw) for volumes of Products 2+
    # to override template formulas.
    raw_products = data.get("produits", [])
    raw_services = data.get("services", [])

    result["products"] = [_convert_product(p) for p in raw_products[:20]]

    # ── Services ──
    result["services"] = [_convert_product(s) for s in raw_services[:10]]

    # ── Staff ──
    result["staff"] = [_convert_staff(s) for s in data.get("staff", [])]

    # ── Ranges & Channels ──
    ranges = data.get("ranges", [])
    if isinstance(ranges, list):
        result["ranges"] = {
            "range1": ranges[0].get("name", "STANDARD") if len(ranges) > 0 else "STANDARD",
            "range2": ranges[1].get("name", "-") if len(ranges) > 1 else "-",
            "range3": ranges[2].get("name", "-") if len(ranges) > 2 else "-",
        }
    else:
        result["ranges"] = ranges or {}

    channels = data.get("channels", [])
    if isinstance(channels, list):
        result["channels"] = {
            "channel1": channels[0].get("name", "B2B") if len(channels) > 0 else "B2B",
            "channel2": channels[1].get("name", "B2C") if len(channels) > 1 else "B2C",
        }
    else:
        result["channels"] = channels or {}

    # ── OPEX ──
    result["opex"] = _convert_opex(data.get("opex", {}))

    # ── CAPEX ──
    result["capex"] = _convert_capex(data.get("capex", []))

    # ── Loans (B8/B9 fix: build amount_by_year + duration) ──
    loans_raw = data.get("loans", {})
    result["loans"] = {}
    for loan_type in ["ovo", "family", "bank"]:
        lr = loans_raw.get(loan_type, {})
        amount = lr.get("amount", 0) or 0
        rate = lr.get("rate", lr.get("interest_rate", 0.10))
        term = lr.get("term_years", lr.get("duration_years", 5)) or 5

        result["loans"][loan_type] = {
            "interest_rate": rate,
            "duration_years": term,
            "amount": amount,
            "amount_by_year": _build_loan_amount_by_year(amount, term),
            "grace_by_year": lr.get("grace_by_year", {}),
            "interest_h1_current_year": round(amount * rate / 2) if amount > 0 else 0,
        }

    # ── Working Capital: arrays → dicts ──
    wc = data.get("working_capital", {})
    result["working_capital"] = {}
    for wk in ["stock_days", "receivable_days", "payable_days"]:
        arr = wc.get(wk, [])
        if isinstance(arr, list):
            result["working_capital"][wk] = _array_to_dict(arr)
        elif isinstance(arr, dict):
            result["working_capital"][wk] = arr
        else:
            result["working_capital"][wk] = {}

    # ── Scenarios (B10 fix: build worst/typical/best) ──
    result["scenarios"] = _convert_scenarios(
        data.get("scenarios", {}),
        data.get("hypotheses_ia", data.get("analyse", {}).get("hypotheses_ia", {}))
    )

    # ── Cash Position ──
    projections = data.get("projections", [])
    ym2 = next((p for p in projections if p.get("annee") == "YEAR-2"), {})
    ym1 = next((p for p in projections if p.get("annee") == "YEAR-1"), {})
    result["cash_position"] = {
        "year_minus_2": data.get("kpis", {}).get("tresorerie", ym2.get("cashflow", 0)),
        "year_minus_1": ym1.get("cashflow", 0),
    }

    # ── Adjustments ──
    result["adjustments"] = data.get("adjustments", {
        "bank_charges": {"year_minus_2": 0, "year_minus_1": 0},
        "interest_ovo": {"year_minus_2": 0, "year_minus_1": 0},
        "interest_family": {"year_minus_2": 0, "year_minus_1": 0},
        "interest_bank": {"year_minus_2": 0, "year_minus_1": 0},
        "taxes": {"regime_1": {}, "regime_2": {}},
    })

    # Count items for logging
    n_prods = len(result.get("products", []))
    n_svcs = len(result.get("services", []))
    n_staff = len(result.get("staff", []))
    n_capex = sum(len(v) for v in result.get("capex", {}).values() if isinstance(v, list))
    n_loans = sum(1 for l in result.get("loans", {}).values() if isinstance(l, dict) and l.get("amount", 0) > 0)

    print(f"[adapter] Converted: {n_prods} products, {n_svcs} services, "
          f"{n_staff} staff, {n_capex} capex, {n_loans} loans")

    return result
