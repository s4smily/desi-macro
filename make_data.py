#!/usr/bin/env python3
"""
make_data.py — converts bd_macro_data.xlsx into data.json for the DCM Weekly dashboard.

Weekly workflow:
    1. Update bd_macro_data.xlsx (same sheets as before: Meta, KPIs, Series,
       Commentary, Calendar, Special).
    2. Run:  python make_data.py bd_macro_data.xlsx
    3. Upload the resulting data.json to your GitHub repository (replace the old one).
       The website updates itself — index.html never needs to change.
"""
import sys, json, math, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd

def nn(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    if isinstance(v, str) and not v.strip(): return None
    return v

def main(xlsx="bd_macro_data.xlsx", out="data.json"):
    xl = pd.read_excel(xlsx, sheet_name=None)

    meta_df = xl["Meta"].set_index("Field")["Value"]
    meta = {
        "weekEnding": nn(meta_df.get("WeekEnding")),
        "title":      nn(meta_df.get("Title")),
        "subtitle":   nn(meta_df.get("Subtitle")),
        "preparedBy": nn(meta_df.get("PreparedBy")),
        "forAudience":nn(meta_df.get("ForAudience")),
        "headline":   nn(meta_df.get("Headline")),
    }
    if isinstance(meta["weekEnding"], pd.Timestamp):
        meta["weekEnding"] = meta["weekEnding"].strftime("%B %d, %Y")

    kpis = []
    for _, r in xl["KPIs"].iterrows():
        kpis.append({
            "cat":  nn(r.get("Category")),
            "ind":  nn(r.get("Indicator")),
            "unit": nn(r.get("Unit")),
            "cur":  nn(r.get("Current")),
            "prev": nn(r.get("Previous")),
            "good": None if pd.isna(r.get("GoodWhenUp")) else int(r["GoodWhenUp"]),
            "dec":  2 if pd.isna(r.get("Decimals")) else int(r["Decimals"]),
            "chg":  nn(r.get("ChangeShown")) or "abs",
            "dash": bool(nn(r.get("OnDashboard"))),
        })

    s = xl["Series"].copy()
    series = {"dates": [d.strftime("%Y-%m") if isinstance(d, pd.Timestamp) else str(d)
                        for d in s["Date"]]}
    for col in s.columns:
        if col == "Date": continue
        series[col] = [nn(v) for v in s[col]]

    commentary = {}
    cdf = xl["Commentary"]
    note_col = cdf.columns[1]
    for _, r in cdf.iterrows():
        sec, txt = nn(r.get("Section")), nn(r.get(note_col))
        if sec: commentary[str(sec).strip()] = (str(txt).strip() if txt else "")

    calendar = [{"date": str(nn(r.get("Date")) or ""), "event": str(nn(r.get("Event")) or "")}
                for _, r in xl["Calendar"].iterrows() if nn(r.get("Event"))]

    special, points = {}, []
    for _, r in xl["Special"].iterrows():
        f, c = nn(r.get("Field")), nn(r.iloc[1])
        if f == "Point" and c: points.append(str(c))
        elif f and c: special[str(f).lower()] = str(c)
    special["points"] = points

    # ---- resolve auto/manual commentary + headline via shared module ----
    from commentary import resolve_section, resolve_headline
    K = {k["ind"]: k for k in kpis}
    sections = ["fx","global","rates","ycurve","inflation","external","reserves","commodity"]
    resolved = {}
    for sec in sections:
        resolved[sec] = resolve_section(sec, commentary.get(sec, ""), K)
    meta["headline"] = resolve_headline(meta.get("headline"), K)

    data = {"meta": meta, "kpis": kpis, "series": series,
            "commentary": commentary, "commentaryResolved": resolved,
            "calendar": calendar, "special": special}
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"Wrote {out}  ({len(kpis)} KPIs, {len(series['dates'])} months of history)")

if __name__ == "__main__":
    main(*(sys.argv[1:] or []))
