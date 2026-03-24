"""
plan_financier_adapter.py — Adapte le nouveau format unifié (plan_financier)
au format attendu par fill_ovo() existant.

Le fill_ovo() a un mapping cellule-par-cellule qui fonctionne.
Plutôt que le réécrire, on transforme le nouveau JSON en ancien JSON.
"""

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


def adapt_plan_financier(data: dict) -> dict:
    """
    Convertit le format plan_financier unifié (v2)
    en format compatible avec fill_ovo() (v1).

    Détection auto: si 'produits' existe → nouveau format, sinon ancien format.
    """
    # Si c'est déjà l'ancien format, retourner tel quel
    if "products" in data and "produits" not in data:
        return data

    # Si pas le nouveau format non plus, retourner tel quel
    if "produits" not in data:
        return data

    print("[adapter] Converting plan_financier v2 → v1 format")

    result = {}

    # ── Pass-through fields (identiques) ──
    for key in ["company", "country", "currency", "exchange_rate_eur",
                "vat_rate", "inflation_rate", "tax_regime_1", "tax_regime_2",
                "years", "financing", "working_capital"]:
        if key in data:
            result[key] = data[key]

    # ── Products: produits[] → products[] ──
    result["products"] = []
    for p in data.get("produits", []):
        product = {
            "name": p.get("nom", "Produit"),
            "active": True,
            "range_flags": [1, 0, 0],  # default
            "channel_flags": [0, 1],    # default
            "per_year": {},
        }

        # Convert par_annee array → per_year dict
        for yr_data in p.get("par_annee", []):
            label = yr_data.get("annee", "")
            yr_key = YEAR_LABEL_MAP.get(label)
            if not yr_key:
                continue

            vol_total = yr_data.get("volume", 0)
            q1 = yr_data.get("volume_q1", 0)
            q2 = yr_data.get("volume_q2", 0)
            q3 = yr_data.get("volume_q3", 0)
            q4 = yr_data.get("volume_q4", 0)

            # For projected years (year2+), use total volume in AI col
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
                # Quarters (for Y-2, Y-1, CY)
                "volume_q1": q1 or round(vol_total * 0.22),
                "volume_q2": q2 or round(vol_total * 0.25),
                "volume_q3": q3 or round(vol_total * 0.27),
                "volume_q4": q4 or (vol_total - round(vol_total * 0.22) - round(vol_total * 0.25) - round(vol_total * 0.27)),
                # Total (for projected years)
                "volume_total": vol_total,
                "total": vol_total,
            }
            product["per_year"][yr_key] = year_entry

            # Also build H1/H2 split for current year
            if label == "CURRENT YEAR":
                h1_vol = round(vol_total * 0.45)
                h2_vol = vol_total - h1_vol
                product["per_year"]["h1"] = {**year_entry, "volume_q1": h1_vol, "volume_total": h1_vol, "total": h1_vol}
                product["per_year"]["h2"] = {**year_entry, "volume_q1": h2_vol, "volume_total": h2_vol, "total": h2_vol}

        result["products"].append(product)

    # ── Services: same structure ──
    result["services"] = []
    for s in data.get("services", []):
        service = {
            "name": s.get("nom", "Service"),
            "active": True,
            "range_flags": [1, 0, 0],
            "channel_flags": [0, 1],
            "per_year": {},
        }
        for yr_data in s.get("par_annee", []):
            label = yr_data.get("annee", "")
            yr_key = YEAR_LABEL_MAP.get(label)
            if not yr_key:
                continue
            vol = yr_data.get("volume", 0)
            service["per_year"][yr_key] = {
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
                "volume_q1": yr_data.get("volume_q1", round(vol * 0.22)),
                "volume_q2": yr_data.get("volume_q2", round(vol * 0.25)),
                "volume_q3": yr_data.get("volume_q3", round(vol * 0.27)),
                "volume_q4": yr_data.get("volume_q4", vol - round(vol * 0.22) - round(vol * 0.25) - round(vol * 0.27)),
                "volume_total": vol,
                "total": vol,
            }
        result["services"].append(service)

    # ── Staff: convert par_annee to per_year ──
    result["staff"] = []
    for s in data.get("staff", []):
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
                "gross_monthly_salary_per_person": yr_data.get("salaire_mensuel_brut", 0),
                "gross_salary": yr_data.get("salaire_mensuel_brut", 0),
                "annual_allowances_per_person": yr_data.get("primes_annuelles", 0),
                "allowances": yr_data.get("primes_annuelles", 0),
                "social_security": yr_data.get("charges_sociales_mensuelles", 0),
            }
            staff_item["per_year"][yr_key] = year_entry

            # H1/H2 for current year
            if label == "CURRENT YEAR":
                staff_item["per_year"]["h1"] = year_entry.copy()
                staff_item["per_year"]["h2"] = year_entry.copy()

        result["staff"].append(staff_item)

    # ── Ranges & Channels ──
    ranges = data.get("ranges", [])
    if isinstance(ranges, list) and len(ranges) >= 1:
        result["ranges"] = {
            "range1": ranges[0].get("name", "STANDARD") if len(ranges) > 0 else "STANDARD",
            "range2": ranges[1].get("name", "-") if len(ranges) > 1 else "-",
            "range3": ranges[2].get("name", "-") if len(ranges) > 2 else "-",
        }
    else:
        result["ranges"] = data.get("ranges", {})

    channels = data.get("channels", [])
    if isinstance(channels, list) and len(channels) >= 1:
        result["channels"] = {
            "channel1": channels[0].get("name", "B2B") if len(channels) > 0 else "B2B",
            "channel2": channels[1].get("name", "B2C") if len(channels) > 1 else "B2C",
        }
    else:
        result["channels"] = data.get("channels", {})

    # ── OPEX: already in expanded format from financial-compute.ts ──
    # Format: { "marketing": { "research": [10 values], "advertising": [10 values] }, ... }
    # This is EXACTLY what fill_ovo expects after expandOpex()
    opex = data.get("opex", {})
    # Rename travel keys if needed
    if "travel" in opex:
        tr = opex["travel"]
        if "nb_travellers" in tr:
            opex["travel"] = {
                "headcount": tr["nb_travellers"],
                "cost_per_person": tr.get("avg_cost", tr.get("cost_per_person", {})),
            }
    # Rename taxes_on_staff → taxes_staff for fill_ovo
    if "taxes_on_staff" in opex and "taxes_staff" not in opex:
        opex["taxes_staff"] = opex.pop("taxes_on_staff")
    result["opex"] = opex

    # ── CAPEX: flat list → categorized dict ──
    capex_list = data.get("capex", [])
    if isinstance(capex_list, list):
        categorized = {"office": [], "production": [], "other": [], "intangible": [], "startup": []}
        for item in capex_list:
            if not isinstance(item, dict):
                continue
            cat = item.get("categorie", item.get("category", "other")).lower()
            # Map to valid categories
            if cat in ("bureau", "office", "office_equipment"):
                cat = "office"
            elif cat in ("production", "production_equipment", "equipment"):
                cat = "production"
            elif cat in ("intangible", "incorporel", "logiciel"):
                cat = "intangible"
            elif cat in ("startup", "demarrage"):
                cat = "startup"
            else:
                cat = "other"

            capex_item = {
                "label": item.get("label", ""),
                "year": item.get("acquisition_year", 0),
                "value": item.get("acquisition_value", item.get("montant", 0)),
                "amort_rate": item.get("amortisation_rate", item.get("taux_amortissement", 0.20)),
            }
            categorized[cat].append(capex_item)
        result["capex"] = categorized
    else:
        result["capex"] = capex_list

    # ── Loans ──
    loans = data.get("loans", {})
    result["loans"] = {
        "ovo": {
            "interest_rate": loans.get("ovo", {}).get("rate", 0.07),
            "duration_years": loans.get("ovo", {}).get("term_years", 5),
            "amount": loans.get("ovo", {}).get("amount", 0),
        },
        "family": {
            "interest_rate": loans.get("family", {}).get("rate", 0.10),
            "duration_years": loans.get("family", {}).get("term_years", 3),
            "amount": loans.get("family", {}).get("amount", 0),
        },
        "bank": {
            "interest_rate": loans.get("bank", {}).get("rate", 0.20),
            "duration_years": loans.get("bank", {}).get("term_years", 2),
            "amount": loans.get("bank", {}).get("amount", 0),
        },
    }

    # ── Working Capital: arrays → dicts ──
    wc = data.get("working_capital", {})
    result["working_capital"] = {}
    for wk in ["stock_days", "receivable_days", "payable_days"]:
        arr = wc.get(wk, [])
        if isinstance(arr, list) and len(arr) >= 10:
            idx_year = ["year_minus_2", "year_minus_1", "h1", "h2", "current_year",
                        "year2", "year3", "year4", "year5", "year6"]
            result["working_capital"][wk] = {idx_year[i]: arr[i] for i in range(min(len(arr), 10))}
        elif isinstance(arr, dict):
            result["working_capital"][wk] = arr
        else:
            result["working_capital"][wk] = {}

    # ── Scenarios ──
    scen = data.get("scenarios", {})
    if isinstance(scen, dict):
        # New format has { pessimiste: [vals], realiste: [vals], optimiste: [vals] }
        # Old format has { revenue_products: { worst, typical, best }, ... }
        # For now pass through — the old fill_ovo handles both
        result["scenarios"] = scen

    print(f"[adapter] Converted: {len(result.get('products',[]))} products, "
          f"{len(result.get('services',[]))} services, "
          f"{len(result.get('staff',[]))} staff categories, "
          f"{len(sum([v for v in result.get('capex',{}).values() if isinstance(v, list)], []))} capex items")

    return result
