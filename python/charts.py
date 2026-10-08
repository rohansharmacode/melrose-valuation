"""Charts for the research note, drawn from the same numbers as the model.

Run:  python3 python/charts.py   (writes note/fig_*.pdf)
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyBboxPatch, Rectangle  # noqa: E402

import dcf  # noqa: E402

NOTE = Path(__file__).resolve().parents[1] / "note"

# Reference palette, light mode (validated: blue/orange pass all checks)
BLUE, ORANGE = "#2a78d6", "#eb6834"
CONTEXT = "#bdbbb3"          # de-emphasised marks
INK, INK2, MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, AXIS, SURFACE = "#e1e0d9", "#c3c2b7", "#ffffff"

plt.rcParams.update({
    "font.family": "TeX Gyre Heros",
    "font.size": 7.5,
    "pdf.fonttype": 42,
    "axes.edgecolor": AXIS,
    "axes.linewidth": 0.6,
    "xtick.color": MUTED, "ytick.color": MUTED,
    "xtick.labelcolor": INK2, "ytick.labelcolor": INK2,
    "xtick.major.size": 0, "ytick.major.size": 0,
})


def rounded_hbar(ax, x0, x1, y, h, color, aspect, r=6.0):
    """Floating range bar with both data ends rounded."""
    ax.add_patch(FancyBboxPatch(
        (x0, y - h / 2), x1 - x0, h,
        boxstyle=f"round,pad=0,rounding_size={r}",
        mutation_aspect=aspect, linewidth=0, facecolor=color))


def football_field(path):
    base = dcf.run(dcf.SCENARIOS["Base"])
    vals = {k: dcf.run(s)["value_p"] for k, s in dcf.SCENARIOS.items()}
    weighted = sum(dcf.SCENARIOS[k]["p"] * v for k, v in vals.items())
    w = base["wacc"]
    sens = [dcf.run(dcf.SCENARIOS["Base"], w + dw, g)["value_p"]
            for dw in (-0.005, 0, 0.005) for g in (0.02, 0.025, 0.03)]
    claims = dcf.NET_DEBT + 330 + dcf.PENSION + dcf.SCENARIOS["Base"]["gg"]
    ebitda25, vc25 = 647 + 138, 324
    senior = [(15.2 * e - claims) / dcf.SHARES_M * 100
              for e in (ebitda25 - vc25, ebitda25)]
    rows = [  # label, low, high, emphasised?
        ("DCF: bear to bull case", vals["Bear"], vals["Bull"], True),
        ("DCF: base case, WACC and growth ±0.5pt", min(sens), max(sens),
         True),
        ("Senior plc takeover multiple (15.2x EBITDA)", *senior, False),
        ("Analyst price targets (15 analysts)", 450.0, 825.0, False),
        ("52-week trading range", 409.9, 682.6, False),
    ]
    fig, ax = plt.subplots(figsize=(7.1, 1.85))
    fig.subplots_adjust(left=0.315, right=0.965, top=0.85, bottom=0.15)
    x_lo, x_hi = 300, 900
    ax.set_xlim(x_lo, x_hi)
    ax.set_ylim(-0.6, len(rows) - 0.4)
    bbox = ax.get_window_extent().transformed(fig.dpi_scale_trans.inverted())
    aspect = ((len(rows) - 0.4 + 0.6) / bbox.height) / ((x_hi - x_lo) /
                                                        bbox.width)
    for i, (label, lo, hi, emph) in enumerate(rows):
        y = len(rows) - 1 - i
        rounded_hbar(ax, lo, hi, y, 0.42, BLUE if emph else CONTEXT, aspect)
        for x, ha, dx in ((lo, "right", -5), (hi, "left", 5)):
            ax.text(x + dx, y, f"{x:,.0f}p", ha=ha, va="center", color=INK2,
                    fontsize=7)
        ax.text(x_lo - 12, y, label, ha="right", va="center", color=INK,
                fontsize=7.5)
    for x, text in ((dcf.PRICE_P, f"Share price {dcf.PRICE_P:.0f}p"),
                    (weighted, f"Fair value {weighted:.0f}p")):
        ax.axvline(x, color=INK, linewidth=0.9, zorder=3)
        ax.text(x, len(rows) - 0.32, text, ha="center", va="bottom",
                color=INK, fontsize=7.5, fontweight="bold")
    ax.set_xticks(range(x_lo, x_hi + 1, 100))
    ax.set_xticklabels([f"{t}p" for t in range(x_lo, x_hi + 1, 100)])
    ax.xaxis.grid(True, color=GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.set_yticks([])
    for side in ("left", "right", "top"):
        ax.spines[side].set_visible(False)
    fig.savefig(path)
    plt.close(fig)
    return dict(weighted=weighted, sens=(min(sens), max(sens)),
                senior=senior, **vals)


def profit_vs_cash(path):
    base = dcf.run(dcf.SCENARIOS["Base"])
    years = ["2025A", "2026E", "2027E", "2028E", "2029E"]
    op = [647.0] + [base["op"][y] for y in range(2026, 2030)]
    vc = [324.0] + [-base["vc"][y] for y in range(2026, 2030)]
    rest = [o - v for o, v in zip(op, vc)]
    fig, ax = plt.subplots(figsize=(3.35, 2.25))
    fig.subplots_adjust(left=0.15, right=0.98, top=0.80, bottom=0.11)
    ax.set_xlim(-0.6, len(years) - 0.4)
    ax.set_ylim(0, 1150)
    bbox = ax.get_window_extent().transformed(fig.dpi_scale_trans.inverted())
    aspect = (1150 / bbox.height) / ((len(years) - 0.4 + 0.6) / bbox.width)
    width, gap = 0.5, 6  # gap in £m between stacked segments
    for i in range(len(years)):
        x0 = i - width / 2
        ax.add_patch(Rectangle((x0, 0), width, rest[i] - gap / 2,
                               facecolor=BLUE, linewidth=0))
        top0, top1 = rest[i] + gap / 2, op[i]
        r = 0.06
        ax.add_patch(FancyBboxPatch(
            (x0, top0), width, top1 - top0,
            boxstyle=f"round,pad=0,rounding_size={r}",
            mutation_aspect=aspect, facecolor=ORANGE, linewidth=0))
        ax.add_patch(Rectangle((x0, top0), width, (top1 - top0) / 2,
                               facecolor=ORANGE, linewidth=0))
    for i in (0, len(years) - 1):
        ax.text(i, op[i] + 25, f"{vc[i] / op[i]:.0%} non-cash", ha="center",
                va="bottom", color=INK, fontsize=7, fontweight="bold")
    ax.set_xticks(range(len(years)))
    ax.set_xticklabels(years)
    ax.set_yticks([0, 250, 500, 750, 1000])
    ax.set_yticklabels(["0", "250", "500", "750", "1,000"])
    ax.yaxis.grid(True, color=GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for side in ("left", "right", "top"):
        ax.spines[side].set_visible(False)
    handles = [Rectangle((0, 0), 1, 1, facecolor=ORANGE),
               Rectangle((0, 0), 1, 1, facecolor=BLUE)]
    fig.legend(handles, ["Non-cash variable consideration",
                         "Rest of adjusted operating profit"],
               loc="upper left", bbox_to_anchor=(0.13, 1.0), ncol=1,
               frameon=False, fontsize=7, handlelength=1.0,
               handleheight=0.8, labelcolor=INK2)
    fig.savefig(path)
    plt.close(fig)
    return dict(zip(years, zip(op, vc)))


if __name__ == "__main__":
    NOTE.mkdir(exist_ok=True)
    print(football_field(NOTE / "fig_football.pdf"))
    print(profit_vs_cash(NOTE / "fig_profit_cash.pdf"))
