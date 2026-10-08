# Melrose Industries (MRO.L): equity research

**Rating: BUY. Fair value 609p against a share price of 450p (7 Oct 2026), 35% upside.**

Melrose owns GKN Aerospace. About half of its profit is *variable consideration* (VC): engine aftermarket income from its risk and revenue sharing partnerships that is booked when engines are delivered but paid in cash over many years. At 450p the market is close to valuing that income at nothing. This project values the cash flows directly, giving the VC only partial credit, and still finds the shares undervalued.

- **[Research note (2 pages, PDF)](note/Melrose_Research_Note.pdf)**: thesis, valuation, catalysts and risks
- **[Valuation model (Excel)](model/Melrose_DCF.xlsx)**: 1,178 live formulas, so changing any input re-prices the model

## Results

| Scenario | Value per share | vs 450p | 2029 free cash flow | Probability |
|---|---:|---:|---:|---:|
| Bear | 363p | -19% | £277m (50% of target) | 25% |
| Base | 623p | +39% | £486m (88% of target) | 50% |
| Bull | 825p | +83% | £614m (112% of target) | 25% |
| **Probability-weighted** | **609p** | **+35%** | | |

Free cash flow percentages compare with management's £600m 2029 target restated at US$1.37 to the pound (about £550m), the rate behind its 2026 guidance.

In the base case, all 25 combinations of WACC (8.4-10.4%) and terminal growth (1.5-3.5%) give a value above the share price.

## Method

- **Ten-year DCF (2026-35)** of unlevered free cash flow, built division by division (Engines and Airframes) from the 2026 guidance mid-points and management's 2029 targets, with growth fading to 2.5% by 2035.
- **Non-cash VC is deducted in full** each year. The key judgement is how fast it falls as a share of Engines revenue (20.6% in 2026 guidance; 16%, 11% and 9% by 2029 in the bear, base and bull cases).
- **WACC 9.4%**: 10-year gilt 5.48%, equity risk premium 5%, beta 1.0, pre-tax cost of debt 7.0%, debt at 21% of capital.
- **Valued at 30 June 2026**, the date of the latest net debt (£1,530m), less the pension deficit and the Garden Grove costs still to come.
- **Cross-checks:** peer P/E multiples and the 15.2x EBITDA that Blackstone and Tinicum agreed to pay for Senior plc in April 2026.

## Workbook guide

| Sheet | What it holds |
|---|---|
| Summary | Rating, fair value, scenario results, football-field ranges |
| Inputs | Every figure from company filings and market data, with its source |
| Scenarios | The judgement calls (yellow) for the bear, base and bull cases |
| DCF_Base, DCF_Bear, DCF_Bull | Revenue, profit, cash flow and valuation, 2025A-2035E |
| Sensitivity | Base-case value across WACC and terminal growth |
| Comps | Peer P/E multiples and the Senior plc takeover cross-check |
| History | Reported figures for FY2024, FY2025, H1 2025 and H1 2026 |

Colour code: blue = typed-in data, black = formula, green = link to another sheet, yellow = key judgement.

## Reproducing it

```bash
pip install openpyxl matplotlib pytest
python3 python/dcf.py              # prints every scenario and the sensitivity grid
python3 python/build_workbook.py   # regenerates model/Melrose_DCF.xlsx (open in Excel to calculate)
python3 python/charts.py           # redraws the note's charts
python3 -m pytest python/test_model.py   # checks 376 workbook values against python/dcf.py
```

`python/dcf.py` is an independent implementation of the same model; the test confirms the spreadsheet and the code agree to within rounding.

## Sources

Melrose FY2025 results (27 Feb 2026) and H1 2026 results (31 Jul 2026); Senior plc Rule 2.7 announcement (7 Apr 2026); lse.co.uk (prices, 7 Oct 2026); S&P Global consensus via stockanalysis.com (7 Oct 2026); Trading Economics (UK 10-year gilt, 8 Oct 2026); Reuters (25 Aug 2026); Spectrum News (28 Sep 2026); interactive investor (27 Feb 2026).

*A personal learning project, not investment advice.*
