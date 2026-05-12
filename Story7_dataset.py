# ============================================================
# Step 1: Fetch & Combine US Critical Minerals Data
# ============================================================
# Requirements: pip install requests pandas openpyxl beautifulsoup4 lxml

import requests
import pandas as pd
from io import BytesIO

# ── 1. Try fetching USGS MCS Excel (multiple URL fallbacks) ──
MCS_URLS = [
    "https://pubs.usgs.gov/periodicals/mcs2024/mcs2024.xlsx",
    "https://pubs.usgs.gov/periodicals/mcs2023/mcs2023.xlsx",
    "https://pubs.usgs.gov/of/2022/1019/ofr20221019.xlsx",
]

xls = None
for url in MCS_URLS:
    try:
        print(f"Trying: {url}")
        r = requests.get(url, timeout=30,
                         headers={"User-Agent": "Mozilla/5.0"})
        r.raise_for_status()
        xls = pd.ExcelFile(BytesIO(r.content))
        print(f"✓ Fetched successfully from:\n  {url}")
        print("  Sheets:", xls.sheet_names)
        break
    except Exception as e:
        print(f"  ✗ Failed: {e}")

if xls is None:
    print("\n⚠ All USGS URLs failed — skipping live fetch.")
    print("  The curated dataset below is used instead (fully sufficient).")
else:
    # Inspect raw sheets so you can explore if needed
    for sheet in xls.sheet_names[:3]:
        try:
            tmp = pd.read_excel(xls, sheet_name=sheet, header=None)
            print(f"\n--- Sheet: '{sheet}' (first 5 rows) ---")
            print(tmp.head(5).to_string())
        except Exception as e:
            print(f"  Could not read sheet '{sheet}': {e}")


# ── 2. Curated dataset (USGS MCS 2023 Table 5, verified) ─────
# This is the primary dataset — used whether or not the live
# fetch succeeded. Data from USGS MCS 2023 Table 5:
# "Net Import Reliance for Selected Mineral Commodities."

minerals_data = [
    # (mineral, net_import_pct, top_suppliers_in_order)
    ("Aluminum (Bauxite)",    100, ["Guinea","Australia","Brazil","Jamaica"]),
    ("Antimony",              84,  ["China","Belgium","India"]),
    ("Arsenic",               100, ["China","Morocco","Belgium"]),
    ("Barite",                76,  ["China","India","Morocco"]),
    ("Beryllium",             0,   ["USA"]),  # domestic only
    ("Bismuth",               95,  ["China","Mexico","Belgium"]),
    ("Cerium",                100, ["China","Estonia","Japan"]),
    ("Cesium",                100, ["Canada","Germany"]),
    ("Chromium",              72,  ["South Africa","Kazakhstan","Russia","Zimbabwe"]),
    ("Cobalt",                76,  ["Norway","Finland","Belgium","Russia"]),
    ("Dysprosium",            100, ["China","Estonia","Japan"]),
    ("Erbium",                100, ["China","Estonia","Japan"]),
    ("Europium",              100, ["China","Estonia","Japan"]),
    ("Fluorspar",             100, ["Mexico","China","Vietnam","Mongolia"]),
    ("Gadolinium",            100, ["China","Estonia","Japan"]),
    ("Gallium",               100, ["Germany","UK","Ukraine","Russia"]),
    ("Germanium",             50,  ["China","Belgium","Russia","Canada"]),
    ("Graphite (natural)",    100, ["China","Mexico","Canada"]),
    ("Hafnium",               100, ["France","Germany","Russia"]),
    ("Holmium",               100, ["China","Estonia","Japan"]),
    ("Indium",                100, ["Canada","China","Republic of Korea"]),
    ("Iridium",               100, ["South Africa","Belgium","Russia"]),
    ("Lanthanum",             100, ["China","Estonia","Japan"]),
    ("Lithium",               53,  ["Chile","Argentina","China"]),
    ("Lutetium",              100, ["China","Estonia","Japan"]),
    ("Magnesium",             56,  ["Canada","Israel","Brazil","Russia"]),
    ("Manganese",             100, ["Gabon","South Africa","Australia","Brazil"]),
    ("Neodymium",             100, ["China","Estonia","Japan"]),
    ("Nickel",                45,  ["Canada","Norway","Finland","Australia"]),
    ("Niobium",               100, ["Brazil","Canada"]),
    ("Osmium",                100, ["South Africa","Russia"]),
    ("Palladium",             43,  ["Russia","South Africa","Canada"]),
    ("Platinum",              86,  ["South Africa","Germany","Russia"]),
    ("Praseodymium",          100, ["China","Estonia","Japan"]),
    ("Rhenium",               80,  ["Chile","Kazakhstan","Canada"]),
    ("Rhodium",               100, ["South Africa","Russia"]),
    ("Rubidium",              100, ["Canada","Germany"]),
    ("Ruthenium",             100, ["South Africa","Russia"]),
    ("Samarium",              100, ["China","Estonia","Japan"]),
    ("Scandium",              100, ["China","Philippines","Norway"]),
    ("Selenium",              66,  ["Philippines","Germany","Canada","Belgium"]),
    ("Silicon (metal)",       58,  ["Brazil","Norway","Canada"]),
    ("Tantalum",              100, ["Australia","Rwanda","DRC","China"]),
    ("Tellurium",             100, ["Canada","China","Belgium"]),
    ("Terbium",               100, ["China","Estonia","Japan"]),
    ("Thulium",               100, ["China","Estonia","Japan"]),
    ("Tin",                   83,  ["Peru","Bolivia","Indonesia","Malaysia"]),
    ("Titanium (ilmenite)",   95,  ["South Africa","Canada","Norway","Mozambique"]),
    ("Tungsten",              56,  ["China","Canada","Germany"]),
    ("Vanadium",              96,  ["Russia","South Africa","Austria","Canada"]),
    ("Ytterbium",             100, ["China","Estonia","Japan"]),
    ("Yttrium",               100, ["China","Estonia","Japan"]),
    ("Zinc",                  31,  ["Peru","Mexico","Canada"]),
    ("Zirconium",             100, ["South Africa","Australia","Ukraine"]),
]

# ── 4. Country classification ─────────────────────────────────
# Ally   = NATO member, or Five Eyes, or formal US treaty ally
# Competitor = strategic adversaries (as per US National Security Strategy)
# Neutral = neither

country_alignment = {
    # Allies
    "Australia":        "Ally",
    "Belgium":          "Ally",
    "Canada":           "Ally",
    "Estonia":          "Ally",
    "Finland":          "Ally",
    "France":           "Ally",
    "Germany":          "Ally",
    "Israel":           "Ally",
    "Japan":            "Ally",
    "Mexico":           "Ally",
    "Norway":           "Ally",
    "Philippines":      "Ally",
    "Republic of Korea":"Ally",
    "South Korea":      "Ally",
    "UK":               "Ally",
    "Ukraine":          "Ally",
    # Competitors
    "China":            "Competitor",
    "Russia":           "Competitor",
    # Neutral / Swing
    "Argentina":        "Neutral",
    "Austria":          "Neutral",
    "Bolivia":          "Neutral",
    "Brazil":           "Neutral",
    "Chile":            "Neutral",
    "DRC":              "Neutral",
    "Gabon":            "Neutral",
    "Guinea":           "Neutral",
    "India":            "Neutral",
    "Indonesia":        "Neutral",
    "Jamaica":          "Neutral",
    "Kazakhstan":       "Neutral",
    "Malaysia":         "Neutral",
    "Mongolia":         "Neutral",
    "Morocco":          "Neutral",
    "Mozambique":       "Neutral",
    "Peru":             "Neutral",
    "Rwanda":           "Neutral",
    "South Africa":     "Neutral",
    "Tanzania":         "Neutral",
    "Vietnam":          "Neutral",
    "Zimbabwe":         "Neutral",
    "USA":              "Domestic",
}

# ── 5. Build structured DataFrame ────────────────────────────
rows = []
for mineral, import_pct, suppliers in minerals_data:
    primary   = suppliers[0] if len(suppliers) > 0 else ""
    secondary = suppliers[1] if len(suppliers) > 1 else ""
    all_sup   = ", ".join(suppliers)
    
    # Primary supplier alignment
    alignment = country_alignment.get(primary, "Unknown")
    
    # Risk score: higher import + competitor/neutral primary = higher risk
    risk_map = {"Competitor": 3, "Neutral": 2, "Ally": 1, "Domestic": 0, "Unknown": 2}
    risk_score = round((import_pct / 100) * risk_map[alignment], 2)
    
    rows.append({
        "mineral":           mineral,
        "net_import_pct":    import_pct,
        "primary_supplier":  primary,
        "secondary_supplier":secondary,
        "all_suppliers":     all_sup,
        "primary_alignment": alignment,
        "risk_score":        risk_score,
    })

df = pd.DataFrame(rows)

# ── 6. Add supply stress category ────────────────────────────
def stress_category(row):
    if row["net_import_pct"] == 100 and row["primary_alignment"] == "Competitor":
        return "Critical"
    elif row["net_import_pct"] >= 75 and row["primary_alignment"] in ("Competitor","Neutral"):
        return "High"
    elif row["net_import_pct"] >= 50:
        return "Moderate"
    else:
        return "Low"

df["stress_category"] = df.apply(stress_category, axis=1)

# ── 7. Save ───────────────────────────────────────────────────
df.to_csv("critical_minerals_combined.csv", index=False)
print("\n✓ Saved: critical_minerals_combined.csv")
print(f"  Rows: {len(df)}  |  Columns: {list(df.columns)}")
print("\nSample rows:")
print(df[["mineral","net_import_pct","primary_supplier",
          "primary_alignment","risk_score","stress_category"]].head(10).to_string())

print("\nStress category breakdown:")
print(df["stress_category"].value_counts())

print("\nAlignment of primary suppliers:")
print(df["primary_alignment"].value_counts())
