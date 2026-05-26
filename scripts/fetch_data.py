"""
fetch_data.py — External Fleet Cost Dashboard
==============================================
Fetch 4 Google Sheets (published XLSX) → normalize → generate data/data.json

Sheet structure:
  Retail 2025     → tabs: 'Data LC AHI TGI FBI Smt 1' + 'Data LC HCI Smt 1'
  Retail 2026     → tabs: 'HCI Group - NDC' + 'AHI Group - NDC'
  Industrial 2025 → tab:  'Raw Data IND JBBK'
  Industrial 2026 → tab:  'Industrial Group'

GitHub Secrets needed:
  RETAIL_2025_URL, RETAIL_2026_URL
  INDUSTRIAL_2025_URL, INDUSTRIAL_2026_URL
"""

import json, os, io, re
from datetime import datetime, timezone
import pandas as pd
import requests

# ─── CONFIG ──────────────────────────────────────────────────────────────────

SHEET_CONFIG = {
    "retail": {
        "2025": {
            "env_key": "RETAIL_2025_URL",
            "tabs":    ["Data LC AHI TGI FBI Smt 1", "Data LC HCI Smt 1"],
            "year":    2025,
        },
        "2026": {
            "env_key": "RETAIL_2026_URL",
            "tabs":    ["HCI Group - NDC", "AHI Group - NDC"],
            "year":    2026,
        },
    },
    "industrial": {
        "2025": {
            "env_key": "INDUSTRIAL_2025_URL",
            "tabs":    ["Raw Data IND JBBK"],
            "year":    2025,
        },
        "2026": {
            "env_key": "INDUSTRIAL_2026_URL",
            "tabs":    ["Industrial Group"],
            "year":    2026,
        },
    },
}

MONTH_ABBR = {
    # Indonesian
    "januari":"Jan","februari":"Feb","maret":"Mar","april":"Apr",
    "mei":"May","juni":"Jun","juli":"Jul","agustus":"Aug",
    "september":"Sep","oktober":"Oct","november":"Nov","desember":"Dec",
    # English
    "january":"Jan","february":"Feb","march":"Mar","may":"May",
    "june":"Jun","july":"Jul","august":"Aug","october":"Oct","december":"Dec",
}

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def fetch_xlsx(url: str) -> pd.ExcelFile:
    r = requests.get(url, timeout=180)
    r.raise_for_status()
    return pd.ExcelFile(io.BytesIO(r.content))


def parse_month(val) -> str | None:
    """'Januari 2025' / 'January 26' / 'Jan-25' → 'Jan' or None."""
    s = str(val).strip()
    parts = re.split(r"[\s\-/]+", s)
    if not parts:
        return None
    key = parts[0].lower()
    return MONTH_ABBR.get(key)


ORIGIN_NORM = {
    "cikarang":         "Jababeka",
    "jababeka":         "Jababeka",
    "jbbk":             "Jababeka",
    "cikupa & jababeka":"Cikupa & Jababeka",
    "cikupa":           "Cikupa",
    "sidoarjo":         "Sidoarjo",
}

PULAU_NORM = {
    "papura": "Papua",
    "bali nusra": "Bali Nusra",
}

def normalize_origin(val: str) -> str:
    key = str(val).strip().lower()
    for k, v in ORIGIN_NORM.items():
        if k in key:
            return v
    return str(val).strip()

def normalize_pulau(val: str) -> str:
    key = str(val).strip().lower()
    return PULAU_NORM.get(key, str(val).strip())

def extract_origin(site_name: str) -> str:
    """Infer Origin from SITE NAME when no explicit Origin column."""
    s = str(site_name).upper()
    if "SIDOARJO" in s:
        return "Sidoarjo"
    if "CIKUPA" in s:
        return "Cikupa"
    if "JABABEKA" in s or "JBBK" in s or "CIKARANG" in s:
        return "Jababeka"
    return "Unknown"


def load_tabs(xl: pd.ExcelFile, tabs: list[str], year: int) -> pd.DataFrame:
    """Load & concat specific tabs, normalize columns."""
    frames = []
    for tab in tabs:
        if tab not in xl.sheet_names:
            print(f"    ⚠️  Tab '{tab}' tidak ada, skip.")
            continue
        df = xl.parse(tab)

        # Must have the core columns
        required = {"LC OWNER", "MONTH", "Area2", "Pulau", "CBM Aktual", "Cost"}
        if not required.issubset(df.columns):
            print(f"    ⚠️  Tab '{tab}' kolom kurang: {required - set(df.columns)}, skip.")
            continue

        df = df.copy()
        df["_year"] = year
        df["_tab"]  = tab

        # Parse month
        df["_month"] = df["MONTH"].apply(parse_month)

        # Origin: use existing column if present, else extract from SITE NAME
        if "Origin" in df.columns:
            site_col = df["SITE NAME"].astype(str) if "SITE NAME" in df.columns else pd.Series([""] * len(df))
            df["_origin"] = [
                normalize_origin(str(o).strip()) if str(o).strip().lower() not in ("nan","none","") else extract_origin(str(s))
                for o, s in zip(df["Origin"], site_col)
            ]
        else:
            site_col = df["SITE NAME"] if "SITE NAME" in df.columns else pd.Series([""] * len(df))
            df["_origin"] = site_col.apply(extract_origin)

        # Keep only numeric cost & CBM
        df["_cost"] = pd.to_numeric(df["Cost"], errors="coerce")
        df["_cbm"]  = pd.to_numeric(df["CBM Aktual"], errors="coerce")

        # Drop rows with no valid data
        df = df.dropna(subset=["_month", "_cost", "_cbm"])
        df = df[df["_cost"] > 0]
        df["Pulau"] = df["Pulau"].apply(lambda x: normalize_pulau(str(x)))

        frames.append(df[["LC OWNER", "Pulau", "Area2", "_month", "_year",
                           "_origin", "_cost", "_cbm"]])

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


# ─── AGGREGATION ─────────────────────────────────────────────────────────────

def agg_by_pulau(df: pd.DataFrame, year: int) -> dict:
    d = df[df["_year"] == year]
    if d.empty:
        return {}
    grp = d.groupby("Pulau").agg(cost=("_cost","sum"), cbm=("_cbm","sum")).reset_index()
    result = {}
    for _, row in grp.iterrows():
        pulau = str(row["Pulau"]).strip()
        if not pulau or pulau == "nan":
            continue
        cbm = float(row["cbm"])
        cost = float(row["cost"])
        result[pulau] = {
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost / cbm, 2) if cbm else 0,
        }
    return result


def agg_by_area(df: pd.DataFrame, year: int) -> dict:
    d = df[df["_year"] == year]
    if d.empty:
        return {}
    grp = d.groupby(["Pulau","Area2"]).agg(cost=("_cost","sum"), cbm=("_cbm","sum")).reset_index()
    result = {}
    for _, row in grp.iterrows():
        pulau = str(row["Pulau"]).strip()
        area  = str(row["Area2"]).strip()
        if not pulau or pulau == "nan" or not area or area == "nan":
            continue
        cbm  = float(row["cbm"])
        cost = float(row["cost"])
        if pulau not in result:
            result[pulau] = {}
        result[pulau][area] = {
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost / cbm, 2) if cbm else 0,
        }
    return result


def agg_by_month(df: pd.DataFrame, year: int) -> dict:
    d = df[df["_year"] == year]
    if d.empty:
        return {}
    order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    grp = d.groupby("_month").agg(cost=("_cost","sum"), cbm=("_cbm","sum")).reset_index()
    result = {}
    for _, row in grp.iterrows():
        m = str(row["_month"])
        cbm  = float(row["cbm"])
        cost = float(row["cost"])
        result[m] = {
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost / cbm, 2) if cbm else 0,
            "sort":  order.index(m) if m in order else 99,
        }
    return result


def agg_by_owner(df: pd.DataFrame, year: int) -> dict:
    d = df[df["_year"] == year]
    if d.empty:
        return {}
    grp = d.groupby("LC OWNER").agg(cost=("_cost","sum"), cbm=("_cbm","sum")).reset_index()
    result = {}
    for _, row in grp.iterrows():
        owner = str(row["LC OWNER"]).strip()
        if not owner or owner == "nan":
            continue
        cbm  = float(row["cbm"])
        cost = float(row["cost"])
        result[owner] = {
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost / cbm, 2) if cbm else 0,
        }
    return result


def agg_by_origin(df: pd.DataFrame, year: int) -> dict:
    d = df[df["_year"] == year]
    if d.empty:
        return {}
    grp = d.groupby("_origin").agg(cost=("_cost","sum"), cbm=("_cbm","sum")).reset_index()
    result = {}
    for _, row in grp.iterrows():
        origin = str(row["_origin"]).strip()
        if not origin or origin in ("nan","Unknown"):
            continue
        cbm  = float(row["cbm"])
        cost = float(row["cost"])
        result[origin] = {
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost / cbm, 2) if cbm else 0,
        }
    return result


def build_section(df_2025: pd.DataFrame, df_2026: pd.DataFrame) -> dict:
    """Build full data structure for one section (retail or industrial)."""
    combined = pd.concat([df_2025, df_2026], ignore_index=True)

    # Meta lists
    owners  = sorted(combined["LC OWNER"].dropna().unique().tolist())
    origins = sorted(combined["_origin"].dropna().unique().tolist())
    months  = sorted(
        combined["_month"].dropna().unique().tolist(),
        key=lambda m: ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].index(m)
        if m in ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"] else 99
    )

    def totals(df, year):
        d = df[df["_year"] == year]
        cbm  = float(d["_cbm"].sum())
        cost = float(d["_cost"].sum())
        return {
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost / cbm, 2) if cbm else 0,
            "rows":  len(d),
        }

    return {
        "meta": {
            "owners":  owners,
            "origins": origins,
            "months":  months,
        },
        "2025": {
            "total":   totals(combined, 2025),
            "byPulau": agg_by_pulau(combined, 2025),
            "byArea":  agg_by_area(combined, 2025),
            "byMonth": agg_by_month(combined, 2025),
            "byOwner": agg_by_owner(combined, 2025),
            "byOrigin":agg_by_origin(combined, 2025),
        },
        "2026": {
            "total":   totals(combined, 2026),
            "byPulau": agg_by_pulau(combined, 2026),
            "byArea":  agg_by_area(combined, 2026),
            "byMonth": agg_by_month(combined, 2026),
            "byOwner": agg_by_owner(combined, 2026),
            "byOrigin":agg_by_origin(combined, 2026),
        },
    }


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    output = {
        "lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }

    for section, years in SHEET_CONFIG.items():
        print(f"\n{'='*50}")
        print(f"📊 Processing: {section.upper()}")
        dfs = {}
        for yr_str, cfg in years.items():
            url = os.environ.get(cfg["env_key"], "").strip()
            if not url:
                print(f"  ⚠️  {cfg['env_key']} tidak ada di environment, skip.")
                dfs[int(yr_str)] = pd.DataFrame()
                continue
            print(f"  📥 Fetching {section} {yr_str}...")
            try:
                xl = fetch_xlsx(url)
                print(f"     Tabs tersedia: {xl.sheet_names}")
                df = load_tabs(xl, cfg["tabs"], cfg["year"])
                print(f"     Rows loaded: {len(df)}")
                dfs[int(yr_str)] = df
            except Exception as e:
                print(f"  ❌ Error: {e}")
                dfs[int(yr_str)] = pd.DataFrame()

        df25 = dfs.get(2025, pd.DataFrame())
        df26 = dfs.get(2026, pd.DataFrame())
        output[section] = build_section(df25, df26)
        print(f"  ✅ {section} done — 2025: {len(df25)} rows, 2026: {len(df26)} rows")

    os.makedirs("data", exist_ok=True)
    out_path = "data/data.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = os.path.getsize(out_path) / 1024
    print(f"\n✅ data.json saved → {size_kb:.1f} KB")


if __name__ == "__main__":
    main()
