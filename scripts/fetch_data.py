"""
fetch_data.py — External Fleet Cost Dashboard
Flat rows per (year, month, owner, origin, pulau, area) for client-side filtering.
"""

import json, os, io, re
from datetime import datetime, timezone
import pandas as pd
import requests

SHEET_CONFIG = {
    "retail": {
        "2025": {"env_key":"RETAIL_2025_URL",     "tabs":["Data LC AHI TGI FBI Smt 1","Data LC HCI Smt 1"], "year":2025},
        "2026": {"env_key":"RETAIL_2026_URL",     "tabs":["HCI Group - NDC","AHI Group - NDC"],             "year":2026},
    },
    "industrial": {
        "2025": {"env_key":"INDUSTRIAL_2025_URL", "tabs":["Raw Data IND JBBK"],  "year":2025},
        "2026": {"env_key":"INDUSTRIAL_2026_URL", "tabs":["Industrial Group"],   "year":2026},
    },
}

MONTH_ABBR = {
    "januari":"Jan","februari":"Feb","maret":"Mar","april":"Apr","mei":"May",
    "juni":"Jun","juli":"Jul","agustus":"Aug","september":"Sep","oktober":"Oct",
    "november":"Nov","desember":"Dec",
    "january":"Jan","february":"Feb","march":"Mar","may":"May","june":"Jun",
    "july":"Jul","august":"Aug","october":"Oct","december":"Dec",
}

ORIGIN_NORM = {
    "cikarang":         "Jababeka",
    "jababeka":         "Jababeka",
    "jbbk":             "Jababeka",
    "cikupa & jababeka":"Cikupa & Jababeka",
    "cikupa":           "Cikupa",
    "sidoarjo":         "Sidoarjo",
}

PULAU_NORM = {"papura": "Papua"}

def parse_month(val):
    parts = re.split(r"[\s\-/]+", str(val).strip())
    return MONTH_ABBR.get(parts[0].lower()) if parts else None

def normalize_origin(val):
    """Normalize origin string → clean category. Returns None if can't determine."""
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None
    key = s.lower()
    for k, v in ORIGIN_NORM.items():
        if k in key:
            return v
    return None  # unknown → will be dropped

def extract_origin_from_site(site_name):
    """Extract origin from SITE NAME column (case-insensitive)."""
    s = str(site_name).upper()
    if "SIDOARJO" in s: return "Sidoarjo"
    if "CIKUPA"   in s: return "Cikupa"
    if "JABABEKA" in s or "JBBK" in s or "CIKARANG" in s: return "Jababeka"
    return None

def get_col(df, *names):
    """Case-insensitive column lookup — returns first match."""
    cols_lower = {str(c).lower(): c for c in df.columns}
    for name in names:
        found = cols_lower.get(name.lower())
        if found:
            return df[found].astype(str)
    return pd.Series([""] * len(df))

def normalize_pulau(val):
    s = str(val).strip()
    return PULAU_NORM.get(s.lower(), s)

def fetch_xlsx(url):
    r = requests.get(url, timeout=180)
    r.raise_for_status()
    return pd.ExcelFile(io.BytesIO(r.content))

def load_tabs(xl, tabs, year):
    frames = []
    for tab in tabs:
        if tab not in xl.sheet_names:
            continue
        df = xl.parse(tab)
        required = {"LC OWNER", "MONTH", "Area2", "Pulau", "CBM Aktual", "Cost"}
        if not required.issubset(df.columns):
            print(f"    ⚠️  Tab '{tab}' kolom kurang, skip.")
            continue
        df = df.copy()
        df["_year"]  = year
        df["_month"] = df["MONTH"].apply(parse_month)

        # Site name — case-insensitive lookup (handles "SITE NAME" and "Site Name")
        site_col = get_col(df, "SITE NAME", "Site Name")

        # Origin: try explicit Origin col first, fallback to site name
        if "Origin" in df.columns:
            df["_origin"] = [
                normalize_origin(str(o)) or extract_origin_from_site(str(s))
                for o, s in zip(df["Origin"], site_col)
            ]
        else:
            df["_origin"] = site_col.apply(extract_origin_from_site)

        df["_cost"] = pd.to_numeric(df["Cost"],       errors="coerce")
        df["_cbm"]  = pd.to_numeric(df["CBM Aktual"], errors="coerce")
        df["Pulau"] = df["Pulau"].apply(normalize_pulau)

        # Drop rows with missing critical data
        df = df.dropna(subset=["_month", "_cost", "_cbm", "_origin"])
        df = df[df["_cost"] > 0]

        frames.append(df[["LC OWNER","Pulau","Area2","_month","_year","_origin","_cost","_cbm"]])
        print(f"    Tab '{tab}': {len(df)} rows")

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def build_section(df25, df26):
    combined = pd.concat([df25, df26], ignore_index=True) if len(df25) or len(df26) else pd.DataFrame()
    if combined.empty:
        return {"meta":{"owners":[],"origins":[],"months":[]},"rows":[]}

    meta_months = sorted(combined["_month"].dropna().unique().tolist(),
        key=lambda m: MONTH_ABBR and ["Jan","Feb","Mar","Apr","May","Jun",
                                       "Jul","Aug","Sep","Oct","Nov","Dec"].index(m)
        if m in ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"] else 99)

    grp = combined.groupby(["_year","_month","LC OWNER","_origin","Pulau","Area2"]).agg(
        cost=("_cost","sum"), cbm=("_cbm","sum")
    ).reset_index()

    rows = []
    for _, r in grp.iterrows():
        cbm  = float(r["cbm"])
        cost = float(r["cost"])
        rows.append({
            "yr":    int(r["_year"]),
            "mo":    r["_month"],
            "owner": str(r["LC OWNER"]).strip(),
            "origin":str(r["_origin"]).strip(),
            "pulau": str(r["Pulau"]).strip(),
            "area":  str(r["Area2"]).strip(),
            "cost":  round(cost),
            "cbm":   round(cbm, 2),
            "ratio": round(cost/cbm, 2) if cbm else 0,
        })

    return {
        "meta": {
            "owners":  sorted(combined["LC OWNER"].dropna().unique().tolist()),
            "origins": sorted(combined["_origin"].dropna().unique().tolist()),
            "months":  meta_months,
        },
        "rows": rows,
    }

def main():
    output = {"lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}

    for section, years in SHEET_CONFIG.items():
        print(f"\n{'='*50}\n📊 {section.upper()}")
        dfs = {}
        for yr_str, cfg in years.items():
            url = os.environ.get(cfg["env_key"], "").strip()
            if not url:
                print(f"  ⚠️  {cfg['env_key']} missing")
                dfs[int(yr_str)] = pd.DataFrame()
                continue
            print(f"  📥 Fetching {section} {yr_str}...")
            try:
                df = load_tabs(fetch_xlsx(url), cfg["tabs"], cfg["year"])
                print(f"  ✅ {len(df)} rows loaded")
                dfs[int(yr_str)] = df
            except Exception as e:
                print(f"  ❌ {e}")
                dfs[int(yr_str)] = pd.DataFrame()

        output[section] = build_section(dfs.get(2025, pd.DataFrame()), dfs.get(2026, pd.DataFrame()))
        print(f"  → {len(output[section]['rows'])} agg rows | origins: {output[section]['meta']['origins']}")

    data_dir = "data"
    if os.path.exists(data_dir) and not os.path.isdir(data_dir):
        os.remove(data_dir)
    os.makedirs(data_dir, exist_ok=True)

    out_path = "data/data.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))
    print(f"\n✅ data.json → {os.path.getsize(out_path)/1024:.1f} KB")

if __name__ == "__main__":
    main()
