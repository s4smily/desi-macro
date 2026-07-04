#!/usr/bin/env python3
"""
commentary.py — the single source of truth for auto-generated text.

Rule (matches the user's original system):
  - If the Excel commentary cell for a section is BLANK -> auto-generate from data.
  - If the user WROTE something -> show only that (manual replaces auto entirely).
  - Same for the cover/intro paragraph (Meta -> Headline).

Both make_data.py (website + JSON) and make_deck.py (PPT/PDF) import from here,
so the website, the deck and the cover paragraph always read identically.
"""

def _f(v, d=2):
    return "—" if v is None else f"{v:,.{d}f}"

def _dir(cur, prev, unit_word="pp", eps=0.005):
    """Return a phrase like 'fell 0.25pp to', 'rose 0.10pp to', 'was unchanged at'."""
    if cur is None or prev is None:
        return "stood at"
    d = cur - prev
    if abs(d) < eps:
        return "was unchanged at"
    verb = "rose" if d > 0 else "fell"
    return f"{verb} {abs(d):,.2f}{unit_word} to"

def _pct_dir(cur, prev, eps=0.05):
    if cur is None or prev is None:
        return "stood at"
    if prev == 0:
        return "was at"
    d = (cur - prev) / abs(prev) * 100
    if abs(d) < eps:
        return "was little changed at"
    verb = "rose" if d > 0 else "fell"
    return f"{verb} {abs(d):,.1f}% to"


def auto_section(section, K):
    """K = dict of indicator-name -> {cur, prev, ...}. Returns list of auto sentences."""
    g   = lambda n, f="cur": (K.get(n, {}) or {}).get(f)
    out = []

    if section == "fx":
        ib, bc = g("Interbank (WAR)"), g("BC Selling (Prime)")
        out.append(f"The taka {_dir(ib, g('Interbank (WAR)','prev'),'')} {_f(ib)}/USD, "
                   f"with BC selling at {_f(bc)}.")

    elif section == "global":
        s3, s12 = g("SOFR 3M"), g("SOFR 12M")
        e3 = g("EURIBOR 3M")
        if None not in (s3, s12):
            out.append(f"Term SOFR spans {_f(s3)}-{_f(s12)}% across tenors, "
                       f"while EURIBOR sits lower at {_f(e3)}%+.")
        if None not in (s3, e3):
            out.append(f"The 3-month SOFR-EURIBOR gap is {_f(s3-e3)}pp, "
                       f"a live input for currency-of-borrowing decisions.")

    elif section == "rates":
        repo, bank = g("Policy (Repo) Rate"), g("Bank Rate")
        out.append(f"The repo rate is {_f(repo)}%, the Bank Rate {_f(bank)}%.")
        cm = g("Call Money")
        out.append(f"Call money {_dir(cm, g('Call Money','prev'))} {_f(cm)}%.")

    elif section == "ycurve":
        d91, d10 = g("T-Bill 91D"), g("T-Bond 10Y")
        if None not in (d91, d10):
            slope = "upward-sloping" if d10 > d91 else "inverted"
            out.append(f"The local-currency curve is {slope}: the 91-day ({_f(d91)}%) "
                       f"sits {'below' if d10>d91 else 'above'} the 10-year ({_f(d10)}%).")
        # avg w/w shift
        tenors = ["T-Bill 91D","T-Bill 182D","T-Bill 364D","T-Bond 2Y","T-Bond 5Y","T-Bond 10Y","T-Bond 20Y"]
        diffs = [g(t)-g(t,"prev") for t in tenors if g(t) is not None and g(t,"prev") is not None]
        if diffs:
            avg = sum(diffs)/len(diffs)
            way = "down" if avg < 0 else "up"
            out.append(f"Week-on-week the curve shifted {way} about {abs(avg):,.2f}pp on average.")

    elif section == "inflation":
        p2p, avg = g("P2P Inflation"), g("Inflation (12M avg)")
        out.append(f"Point-to-point inflation {_dir(p2p, g('P2P Inflation','prev'),'pp')} {_f(p2p)}%.")
        out.append(f"The 12-month average {_dir(avg, g('Inflation (12M avg)','prev'),'pp')} {_f(avg)}%.")

    elif section == "external":
        rem = g("Remittance (monthly)")
        out.append(f"Remittances {_pct_dir(rem, g('Remittance (monthly)','prev'))} ${_f(rem)}bn.")
        exp, imp = g("Export (monthly)"), g("Import (monthly)")
        if None not in (exp, imp):
            out.append(f"Exports were ${_f(exp)}bn against imports of ${_f(imp)}bn - "
                       f"a trade gap of ${_f(imp-exp)}bn.")

    elif section == "reserves":
        gr, b6 = g("Reserve Gross"), g("Reserve BPM6")
        out.append(f"Gross reserves {_pct_dir(gr, g('Reserve Gross','prev'))} ${_f(gr)}bn "
                   f"(BPM6 ${_f(b6)}bn).")

    elif section == "commodity":
        gold, oil = g("Gold"), g("Oil (Brent)")
        out.append(f"Gold {_pct_dir(gold, g('Gold','prev'))} ${_f(gold,0)}/oz.")
        out.append(f"Brent {_pct_dir(oil, g('Oil (Brent)','prev'))} ${_f(oil)}/bbl.")

    return out


def resolve_section(section, manual, K):
    """Return the list of lines to show for a section.
       Manual comment (if any) is shown FIRST, then the auto-generated lines
       are ALWAYS appended alongside it."""
    lines = []
    manual = (manual or "").strip()
    if manual:
        lines.append({"text": manual, "manual": True})
    for t in auto_section(section, K):
        lines.append({"text": t, "manual": False})
    return lines


def auto_headline(K):
    g = lambda n: (K.get(n, {}) or {}).get("cur")
    return (f"Reserves near ${_f(g('Reserve Gross'))}bn and ${_f(g('Remittance (monthly)'))}bn "
            f"of monthly remittances anchor the external sector, while inflation at "
            f"{_f(g('P2P Inflation'))}% keeps the repo rate on hold at {_f(g('Policy (Repo) Rate'))}%. "
            f"Globally, term SOFR holds around {_f(g('SOFR 3M'))}% while Brent trades near "
            f"${_f(g('Oil (Brent)'),0)}/bbl.")


def resolve_headline(manual, K):
    manual = (manual or "").strip()
    return manual if manual else auto_headline(K)
