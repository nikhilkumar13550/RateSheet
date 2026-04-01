"""
VHA Lives & Volumes Processor
Parses raw Lives & Volumes data → structured benefit groups for Proposed Rates report
"""
import pandas as pd
import numpy as np
from io import BytesIO


# ── Division Group Definitions ──────────────────────────────────────────────
DIV_GROUP_1_8   = [1, 8]    # Toronto Office + OPSEU
DIV_GROUP_2_4_9 = [2, 4, 9] # Toronto Field + Durham Field + Toronto Central OT
DIV_GROUP_1     = [1]
DIV_GROUP_8     = [8]
DIV_GROUP_2_4   = [2, 4]
DIV_GROUP_9     = [9]

# ── Default Current Experience Rates (from reference template) ───────────────
DEFAULT_RATES = {
    "LIFE": {
        "1,8":   {"exp_rate": 0.110, "dr_rate": 0.0},
        "2,4,9": {"exp_rate": 0.133, "dr_rate": 0.0},
    },
    "DEPL": {
        "1,8":   {"exp_rate": 2.307, "dr_rate": 0.0},
        "2,4,9": {"exp_rate": 1.189, "dr_rate": 0.0},
    },
    "ADD": {
        "All":   {"exp_rate": 0.030, "dr_rate": 0.0},
    },
    "LTD": {
        "1,8":   {"exp_rate": 2.811, "dr_rate": 0.0},
    },
    "STD": {
        "1":     {"exp_rate": 0.381, "dr_rate": 0.0},
        "2,4,9": {"exp_rate": 0.872, "dr_rate": 0.0},
        "8":     {"exp_rate": 0.460, "dr_rate": 0.0},
    },
}

# ── Default Proposed Rate Adjustments (%) ───────────────────────────────────
DEFAULT_ADJUSTMENTS = {
    "LIFE":  -0.20,
    "DEPL":  -0.20,
    "ADD":    0.00,
    "LTD":    0.20,
    "STD":   -0.10,
    "EHC":    0.15,
    "DENT":  -0.05,
}


def parse_and_clean(file_bytes: bytes) -> dict:
    """
    Parse raw L&V Excel (first sheet only) and return cleaned DataFrame + metadata.
    Returns dict with keys: df, client_name, period, contract_numbers
    """
    df_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=None)

    # ── Extract metadata from header rows ──────────────────────────────────
    client_name = "VHA Home HealthCare"
    period_str  = ""
    for idx in range(min(12, len(df_raw))):
        row_vals = df_raw.iloc[idx].dropna().astype(str).tolist()
        for v in row_vals:
            if "PERIOD REPORTED" in v.upper():
                period_str = v
            if "VHA" in v.upper() or "Home Health" in v.lower():
                client_name = v.strip()

    # ── Find header row ────────────────────────────────────────────────────
    header_row_idx = None
    _CONTRACT_KEYWORDS = ("contract number", "contract no", "contractnumber", "contract#")
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        row_str = [str(c).strip().lower() for c in row.values if str(c).strip() != "nan"]
        if any(any(kw in cell for kw in _CONTRACT_KEYWORDS) for cell in row_str):
            header_row_idx = idx
            break

    if header_row_idx is None:
        # Collect first non-empty row values to help diagnose
        sample_rows = []
        for idx in range(min(15, len(df_raw))):
            vals = [str(c).strip() for c in df_raw.iloc[idx].values if str(c).strip() not in ("", "nan")]
            if vals:
                sample_rows.append(f"  Row {idx}: {vals[:6]}")
        hint = "\n".join(sample_rows) if sample_rows else "  (file appears empty)"
        raise ValueError(
            f"Could not find a 'Contract Number' header row in the Excel file.\n"
            f"First rows found:\n{hint}"
        )

    # ── Re-read with proper headers ────────────────────────────────────────
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=header_row_idx)

    # ── Map columns by position (original has blank interleaved columns) ───
    cols = df.columns.tolist()
    rename_map = {
        cols[0]:  "Contract Number",
        cols[2]:  "Billing Division",
        cols[4]:  "Division Name",
        cols[5]:  "Class",
        cols[7]:  "Plan",
        cols[8]:  "Benefit",
        cols[9]:  "Option",
        cols[10]: "Billing Period",
        cols[12]: "Lives",
        cols[13]: "Volumes",
        cols[14]: "Rate Type",
        cols[15]: "Rate",
        cols[16]: "Bill Type",
    }
    df = df.rename(columns=rename_map)
    df = df[list(rename_map.values())]

    # ── Drop blank / non-data rows ─────────────────────────────────────────
    df = df.dropna(subset=["Contract Number"])
    df = df[df["Contract Number"].astype(str).str.strip().str.match(r"^0*\d{5,}$")]

    # ── Clean Contract Number (strip leading zeros) ────────────────────────
    df["Contract Number"] = (
        df["Contract Number"].astype(str).str.strip().str.lstrip("0").astype(int)
    )

    # ── Clean Billing Division ─────────────────────────────────────────────
    df["Billing Division"] = pd.to_numeric(
        df["Billing Division"].astype(str).str.strip().str.lstrip("0").replace("", "0"),
        errors="coerce"
    ).fillna(0).astype(int)

    # ── Clean numeric fields ───────────────────────────────────────────────
    df["Lives"]   = pd.to_numeric(df["Lives"],   errors="coerce").fillna(0).astype(int)
    df["Volumes"] = pd.to_numeric(df["Volumes"], errors="coerce").fillna(0).astype(int)
    df["Rate"]    = pd.to_numeric(df["Rate"],    errors="coerce").fillna(0).astype(float)

    # ── Normalise string fields ────────────────────────────────────────────
    for col in ["Plan", "Benefit", "Option", "Rate Type", "Bill Type"]:
        df[col] = df[col].astype(str).str.strip()

    # ── Parse Billing Period ───────────────────────────────────────────────
    df["Billing Period"] = pd.to_datetime(df["Billing Period"], errors="coerce").dt.date

    # ── Reset index ───────────────────────────────────────────────────────
    df = df.reset_index(drop=True)

    contract_numbers = sorted(df["Contract Number"].unique().tolist())

    return {
        "df": df,
        "client_name": client_name,
        "period_str":  period_str,
        "contract_numbers": contract_numbers,
    }


def _plans_str(df_subset):
    """Return comma-separated sorted unique plan letters."""
    plans = sorted(df_subset["Plan"].unique().tolist())
    return ",".join(plans)


def _div_str(divs):
    """Convert list of division codes to display string."""
    return ",".join(str(d) for d in sorted(divs))


def compute_benefit_groups(df: pd.DataFrame) -> dict:
    """
    Compute aggregated lives & volumes for each benefit group.
    Returns structured dict consumed by the Excel generator.
    """
    results = {}

    # ── Helper filters ─────────────────────────────────────────────────────
    def grp(benefit, divs, extra_filter=None):
        mask = (df["Benefit"] == benefit) & (df["Billing Division"].isin(divs))
        if extra_filter is not None:
            mask = mask & extra_filter
        sub = df[mask]
        return sub

    # ── 1. BASIC LIFE (Basic Employee Life) ────────────────────────────────
    results["LIFE"] = []
    for divs, div_label in [
        (DIV_GROUP_1_8,   "1,8"),
        (DIV_GROUP_2_4_9, "2,4,9"),
    ]:
        sub = grp("Basic Employee Life", divs)
        results["LIFE"].append({
            "division":  div_label,
            "plans":     _plans_str(sub),
            "lives":     int(sub["Lives"].sum()),
            "volumes":   int(sub["Volumes"].sum()),
            "basis":     "Per",
            "basis_unit": "1000",
        })

    # ── 2. DEPENDENT LIFE (Basic Dependent Life, Option != 00) ────────────
    results["DEPL"] = []
    for divs, div_label in [
        (DIV_GROUP_1_8,   "1,8"),
        (DIV_GROUP_2_4_9, "2,4,9"),
    ]:
        sub = grp("Basic Dependent Life", divs, df["Option"] != "00")
        # also list all plans from the division (including 0-lives)
        all_plans = grp("Basic Dependent Life", divs)  # include option 00 for plan list
        all_in_div = df[df["Billing Division"].isin(divs)]
        results["DEPL"].append({
            "division":  div_label,
            "plans":     _plans_str(all_in_div),
            "lives":     int(sub["Lives"].sum()),
            "volumes":   None,   # Per Member – no salary basis
            "basis":     "Per",
            "basis_unit": "Member",
        })

    # ── 3. AD&D (Basic AD&D) ── All divisions ─────────────────────────────
    sub_add = df[df["Benefit"] == "Basic AD&D"]
    results["ADD"] = [{
        "division":  "All",
        "plans":     "All",
        "lives":     int(sub_add["Lives"].sum()),
        "volumes":   int(sub_add["Volumes"].sum()),
        "basis":     "Per",
        "basis_unit": "1000",
    }]

    # ── 4. LTD (Long Term Disability) ──────────────────────────────────────
    sub_ltd = grp("Long Term Disability", DIV_GROUP_1_8)
    sub_ltd = sub_ltd[sub_ltd["Lives"] > 0]  # exclude zero-life rows
    if len(sub_ltd) > 0:
        results["LTD"] = [{
            "division":  "1,8",
            "plans":     _plans_str(sub_ltd),
            "lives":     int(sub_ltd["Lives"].sum()),
            "volumes":   int(sub_ltd["Volumes"].sum()),
            "basis":     "Per",
            "basis_unit": "100",
        }]
    else:
        results["LTD"] = []

    # ── 5. STD (Short Term Disability) ─────────────────────────────────────
    results["STD"] = []
    for divs, div_label in [
        (DIV_GROUP_1,     "1"),
        (DIV_GROUP_2_4_9, "2,4,9"),
        (DIV_GROUP_8,     "8"),
    ]:
        sub = grp("Short Term Disability", divs)
        if sub["Lives"].sum() > 0 or sub["Volumes"].sum() > 0:
            results["STD"].append({
                "division":  div_label,
                "plans":     _plans_str(sub),
                "lives":     int(sub["Lives"].sum()),
                "volumes":   int(sub["Volumes"].sum()),
                "basis":     "Per",
                "basis_unit": "10",
            })

    # ── 6. EHC (Extended Health Care) ── group by rate within division group
    results["EHC"] = []
    ehc_df = df[df["Benefit"] == "Extended Health Care"]

    # Group by division cluster and rate_type, using actual rate from data
    for divs, div_label in [
        ([1],     "1"),
        ([8],     "8"),
        ([2, 4],  "2,4"),
        ([9],     "9"),
    ]:
        sub_div = ehc_df[ehc_df["Billing Division"].isin(divs)]
        if sub_div.empty:
            continue
        for rate_type in ["Single", "Family"]:
            sub = sub_div[sub_div["Rate Type"] == rate_type]
            if sub.empty or sub["Lives"].sum() == 0:
                continue
            # Get the rate from the data (should be same for all rows in this group)
            current_rate = float(sub["Rate"].mode().iloc[0]) if not sub["Rate"].empty else 0
            results["EHC"].append({
                "division":   div_label,
                "plans":      _plans_str(sub),
                "lives":      int(sub["Lives"].sum()),
                "volumes":    None,
                "basis":      None,
                "basis_unit": rate_type,
                "current_rate": current_rate,
            })

    # ── 7. DENT (Dental Care) ── group by rate within division group
    results["DENT"] = []
    dent_df = df[df["Benefit"] == "Dental"]

    for divs, div_label in [
        (DIV_GROUP_1_8,   "1,8"),
        ([2, 4],          "2,4"),
        ([9],             "9"),
    ]:
        sub_div = dent_df[dent_df["Billing Division"].isin(divs)]
        if sub_div.empty:
            continue
        for rate_type in ["Single", "Family"]:
            sub = sub_div[sub_div["Rate Type"] == rate_type]
            if sub.empty or sub["Lives"].sum() == 0:
                continue
            current_rate = float(sub["Rate"].mode().iloc[0]) if not sub["Rate"].empty else 0
            results["DENT"].append({
                "division":   div_label,
                "plans":      _plans_str(sub),
                "lives":      int(sub["Lives"].sum()),
                "volumes":    None,
                "basis":      None,
                "basis_unit": rate_type,
                "current_rate": current_rate,
            })

    return results


def compute_report_data(parsed: dict, rates: dict = None, adjustments: dict = None) -> dict:
    """
    Full pipeline: parse → clean → group → compute rates & premiums.

    rates:       override DEFAULT_RATES  (same structure)
    adjustments: override DEFAULT_ADJUSTMENTS (dict of benefit_code → float)
    """
    df = parsed["df"]
    benefit_groups = compute_benefit_groups(df)

    # Merge supplied rates/adjustments with defaults
    final_rates = {**DEFAULT_RATES}
    if rates:
        for k, v in rates.items():
            if k in final_rates:
                final_rates[k].update(v)
            else:
                final_rates[k] = v

    final_adj = {**DEFAULT_ADJUSTMENTS}
    if adjustments:
        final_adj.update(adjustments)

    # ── Attach rate info to each row ────────────────────────────────────────
    def attach_rates(rows, benefit_code, basis_denominator):
        enriched = []
        for row in rows:
            div = row["division"]
            # For EHC/DENT, rate comes from data; for others from config
            if "current_rate" in row:
                curr_exp = row["current_rate"]
                curr_dr  = 0.0
            else:
                cfg = final_rates.get(benefit_code, {}).get(div, {"exp_rate": 0, "dr_rate": 0})
                curr_exp = cfg["exp_rate"]
                curr_dr  = cfg["dr_rate"]

            adj = final_adj.get(benefit_code, 0.0)
            # Proposed exp rate: =ROUND(curr_exp * (1 + adj), 3)
            prop_exp = round(curr_exp * (1 + adj), 3)
            prop_dr  = round(curr_dr, 3)

            curr_total = round(curr_exp + curr_dr, 3)
            prop_total = round(prop_exp + prop_dr, 3)

            lives   = row["lives"]
            volumes = row.get("volumes") or 0

            # Premium calculation: =ROUND(volumes * total_rate / basis, 2)
            basis = row.get("basis_unit", "")
            if basis in ("Single", "Family", "Member"):
                curr_premium = round(curr_total * lives, 2)
                prop_premium = round(prop_total * lives, 2)
            elif basis_denominator and basis_denominator > 0 and volumes:
                curr_premium = round(curr_total * volumes / basis_denominator, 2)
                prop_premium = round(prop_total * volumes / basis_denominator, 2)
            else:
                curr_premium = 0.0
                prop_premium = 0.0

            rate_chg  = round((prop_exp - curr_exp) / curr_exp, 4) if curr_exp else 0
            total_chg = round((prop_total - curr_total) / curr_total, 4) if curr_total else 0

            enriched.append({
                **row,
                "curr_exp_rate":   round(curr_exp, 3),
                "curr_dr_rate":    round(curr_dr,  3),
                "curr_total_rate": curr_total,
                "curr_premium":    curr_premium,
                "prop_exp_rate":   prop_exp,
                "prop_dr_rate":    prop_dr,
                "prop_total_rate": prop_total,
                "prop_premium":    prop_premium,
                "exp_rate_chg":    rate_chg,
                "dr_pct":          0.0,
                "total_rate_chg":  total_chg,
            })
        return enriched

    BASIS_DENOM = {
        "LIFE": 1000, "ADD": 1000, "DEPL": None,
        "LTD": 100,   "STD": 10,   "EHC": None, "DENT": None,
    }

    enriched_groups = {}
    for code, rows in benefit_groups.items():
        enriched_groups[code] = attach_rates(rows, code, BASIS_DENOM.get(code))

    return {
        "client_name":       parsed["client_name"],
        "period_str":        parsed["period_str"],
        "contract_numbers":  parsed["contract_numbers"],
        "groups":            enriched_groups,
        "adjustments":       final_adj,
        "rates_config":      final_rates,
    }
