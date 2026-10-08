"""Melrose Industries (MRO.L) discounted cash flow model.

A Python replica of model/Melrose_DCF.xlsx, used to cross-check the
spreadsheet. Money is in GBP millions and per-share values are in pence.
Forecasts are at US$1.37/GBP, the exchange rate behind Melrose's 2026
guidance, so they can be compared with that guidance directly.

Run:  python3 python/dcf.py
"""

YEARS = list(range(2026, 2036))  # 2026E to 2035E

# ---------------------------------------------------------------- inputs
# Market data
PRICE_P = 450.0      # closing price, 7 Oct 2026 (pence)
SHARES_M = 1246.0    # shares in issue excluding treasury, 30 Jun 2026
NET_DEBT = 1530.0    # 30 Jun 2026, excludes lease liabilities
PENSION = 11.0       # net pension deficit, 30 Jun 2026

# FY2025 actuals
ENG_REV_25, AF_REV_25 = 1632.0, 1957.0
ENG_OP_25, AF_OP_25, CORP_25 = 520.0, 156.0, -29.0
DA_25 = 138.0

# FY2026 guidance mid-points (excludes Garden Grove, at US$1.37)
ENG_REV_26, ENG_OP_26 = 1750.0, 580.0
AF_REV_26, AF_OP_26 = 2100.0, 180.0
GROUP_OP_26 = 725.0
VC_26 = 360.0        # variable consideration (non-cash RRSP profit)

# Garden Grove: lost 2026 revenue and profit (H1 actual + H2 estimate)
GG_REV_26, GG_OP_26 = 46.0, 39.0

# H1 2026 actuals, used to keep only the second half of 2026
H1_OP, H1_VC = 347.0, 206.0
H1_FCF_PRE_FACTORING, H1_INT_TAX, H1_GG_CASH = 28.0, 71.0, 5.0

# Assumptions shared by every scenario
DA_PCT = 0.038       # depreciation and amortisation, % of revenue
CAPEX_PCT = 0.030    # capital expenditure, % of revenue
LEASE_PCT = 0.009    # lease payments, % of revenue
WC_PCT = 0.10        # working capital, % of each year's revenue increase
CORP_GROWTH = 0.03   # growth in central costs
TAX = 0.204          # adjusted effective tax rate (FY2025)
GTF = {2026: 50.0, 2027: 30.0, 2028: 15.0}       # powder-metal cash cost
PENSION_CASH = 5.0
RESTR_OTHER = {2026: 25.0, 2027: 15.0, 2028: 5.0}
FADE_TO = 0.025      # growth that 2030-35 fades towards

# Discount rate
RF, ERP, BETA = 0.0548, 0.050, 1.00
KD_SPREAD, TAX_SHIELD = 0.015, 0.25
TERMINAL_G = 0.025

SCENARIOS = {
    #        Engines growth, 2029 margin | Airframes growth, 2029 margin |
    #        non-cash VC as % of Engines revenue in 2029 and 2035 |
    #        remaining Garden Grove cash cost | probability
    "Bear": dict(eng_g=0.050, eng_m29=0.320, af_g=0.030, af_m29=0.085,
                 vc29=0.160, vc35=0.080, gg=250.0, p=0.25),
    "Base": dict(eng_g=0.080, eng_m29=0.350, af_g=0.050, af_m29=0.110,
                 vc29=0.110, vc35=0.050, gg=120.0, p=0.50),
    "Bull": dict(eng_g=0.100, eng_m29=0.370, af_g=0.065, af_m29=0.125,
                 vc29=0.090, vc35=0.025, gg=60.0, p=0.25),
}


def wacc() -> float:
    equity = PRICE_P * SHARES_M / 100
    ke = RF + BETA * ERP
    kd = (RF + KD_SPREAD) * (1 - TAX_SHIELD)
    return (equity * ke + NET_DEBT * kd) / (equity + NET_DEBT)


def lerp(a: float, b: float, frac: float) -> float:
    return a + (b - a) * frac


def run(s: dict, rate: float | None = None, g: float = TERMINAL_G) -> dict:
    rate = wacc() if rate is None else rate
    eng_rev, af_rev = {2025: ENG_REV_25}, {2025: AF_REV_25}
    rev_total = {2025: ENG_REV_25 + AF_REV_25}
    out = {k: {} for k in (
        "eng_rev", "af_rev", "rev", "eng_m", "af_m", "op", "da", "ebitda",
        "lease", "capex", "wc", "vc_pct", "vc", "gtf", "pension", "restr",
        "pre_tax", "tax", "fcff", "cash_flow", "t", "df", "pv", "fcf_levered")}
    eng_m26, af_m26 = ENG_OP_26 / ENG_REV_26, AF_OP_26 / AF_REV_26
    vc26 = VC_26 / ENG_REV_26
    corp = -(ENG_OP_26 + AF_OP_26 - GROUP_OP_26)
    kd_after_tax = (RF + KD_SPREAD) * (1 - TAX_SHIELD)
    for y in YEARS:
        n = y - 2029  # years after 2029
        if y == 2026:
            eng_rev[y], af_rev[y] = ENG_REV_26, AF_REV_26
        else:
            ge = s["eng_g"] if n <= 0 else lerp(s["eng_g"], FADE_TO, n / 6)
            ga = s["af_g"] if n <= 0 else lerp(s["af_g"], FADE_TO, n / 6)
            eng_rev[y] = eng_rev[y - 1] * (1 + ge)
            af_rev[y] = af_rev[y - 1] * (1 + ga)
            corp *= 1 + CORP_GROWTH
        k = min(max(y - 2026, 0), 3) / 3
        eng_m = lerp(eng_m26, s["eng_m29"], k)
        af_m = lerp(af_m26, s["af_m29"], k)
        gg_rev = -GG_REV_26 if y == 2026 else 0.0
        gg_op = -GG_OP_26 if y == 2026 else 0.0
        rev = eng_rev[y] + af_rev[y] + gg_rev
        op = eng_rev[y] * eng_m + af_rev[y] * af_m + gg_op + corp
        vc_pct = (lerp(vc26, s["vc29"], k) if n <= 0
                  else lerp(s["vc29"], s["vc35"], n / 6))
        da = DA_PCT * rev
        lease = -LEASE_PCT * rev
        capex = -CAPEX_PCT * rev
        wc = -WC_PCT * (rev - rev_total[y - 1])
        vc = -vc_pct * eng_rev[y]
        gtf = -GTF.get(y, 0.0)
        pension = -PENSION_CASH
        restr = -RESTR_OTHER.get(y, 0.0)
        pre_tax = op + da + lease + capex + wc + vc + gtf + pension + restr
        tax = -TAX * max(0.0, op + vc)
        fcff = pre_tax + tax
        if y == 2026:  # first half already in the 30 June net debt
            h1 = (H1_FCF_PRE_FACTORING + H1_INT_TAX + H1_GG_CASH
                  - TAX * (H1_OP - H1_VC))
            cash_flow, t = fcff - h1, 0.25
        else:
            cash_flow, t = fcff, y - 2026.0
        df = 1 / (1 + rate) ** t
        rev_total[y] = rev
        for key, val in (
            ("eng_rev", eng_rev[y]), ("af_rev", af_rev[y]), ("rev", rev),
            ("eng_m", eng_m), ("af_m", af_m), ("op", op), ("da", da),
            ("ebitda", op + da), ("lease", lease), ("capex", capex),
            ("wc", wc), ("vc_pct", vc_pct), ("vc", vc), ("gtf", gtf),
            ("pension", pension), ("restr", restr), ("pre_tax", pre_tax),
            ("tax", tax), ("fcff", fcff), ("cash_flow", cash_flow),
            ("t", t), ("df", df), ("pv", cash_flow * df),
            ("fcf_levered", fcff - kd_after_tax * NET_DEBT)):
            out[key][y] = val
    tv = out["fcff"][2035] * (1 + g) / (rate - g)
    pv_tv = tv / (1 + rate) ** 9.5
    ev = sum(out["pv"].values()) + pv_tv
    equity = ev - NET_DEBT - PENSION - s["gg"]
    out.update(
        wacc=rate, tv=tv, pv_tv=pv_tv, ev=ev, equity=equity,
        value_p=equity / SHARES_M * 100,
        tv_share=pv_tv / ev, exit_multiple=tv / out["ebitda"][2035])
    return out


def main() -> None:
    w = wacc()
    print(f"WACC {w:.4%}")
    weighted = 0.0
    for name, s in SCENARIOS.items():
        r = run(s)
        weighted += s["p"] * r["value_p"]
        print(f"\n{name}: value {r['value_p']:.0f}p "
              f"({r['value_p'] / PRICE_P - 1:+.0%}), EV {r['ev']:,.0f}, "
              f"TV share {r['tv_share']:.0%}, exit {r['exit_multiple']:.1f}x")
        print("year      rev     op  margin    vc   fcff  fcf(lev)")
        for y in YEARS:
            print(f"{y}  {r['rev'][y]:7,.0f} {r['op'][y]:6,.0f} "
                  f"{r['op'][y] / r['rev'][y]:6.1%} {-r['vc'][y]:5.0f} "
                  f"{r['fcff'][y]:6.0f} {r['fcf_levered'][y]:7.0f}")
    print(f"\nProbability-weighted value {weighted:.0f}p "
          f"({weighted / PRICE_P - 1:+.0%})")
    base = SCENARIOS["Base"]
    print("\nBase value (p) by WACC (rows) and terminal growth (columns)")
    gs = [0.015, 0.020, 0.025, 0.030, 0.035]
    print("        " + "".join(f"{g:8.1%}" for g in gs))
    for rate in (0.0836, 0.0886, 0.0936, 0.0986, 0.1036):
        vals = "".join(f"{run(base, rate, g)['value_p']:8.0f}" for g in gs)
        print(f"{rate:7.2%} {vals}")


if __name__ == "__main__":
    main()
