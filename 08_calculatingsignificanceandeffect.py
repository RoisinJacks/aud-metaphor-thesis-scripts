#!/usr/bin/env python3
"""
metaphor_stats_and_plots.py

One-stop script to:
1) compute corpus-comparison statistics for each metaphor vehicle grouping:
   - Log-likelihood (G^2; aka G-test / log-likelihood ratio) with chi-square p-value (df=1)
   - Log ratio (log2 effect size) using relative frequencies (per 1,000 words)
2) generate publication-ready visualisations that use those statistics:
   - Diverging bar chart of log ratio (SSC vs LEC) with significance markers
   - Bubble plot of log ratio vs vehicle group (bubble size ~ log-likelihood; outline = non-significant)

Inputs
------
A CSV file with columns:
    Vehicle_group, SSC_Raw, SSC_RF, LEC_Raw, LEC_RF
and (recommended) a row where Vehicle_group == "TOTALS" used to infer corpus sizes.

If the TOTALS row is missing, provide corpus sizes manually via:
    --ssc_words and --lec_words

Outputs
-------
- metaphor_statistics_results.csv (full table)
- metaphor_log_ratio_comparison.png (diverging bar chart)
- metaphor_significance_bubble_plot.png (bubble plot)
"""

from __future__ import annotations

import argparse
import math
from pathlib import Path
from typing import Optional, Tuple, List

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import chi2


# ----------------------------
# Statistics
# ----------------------------

def infer_corpus_sizes_from_totals(df: pd.DataFrame) -> Tuple[float, float]:
    """
    Infer corpus sizes from a TOTALS row, assuming RF is per 1,000 words.
    """
    totals = df.loc[df["Vehicle_group"].astype(str).str.upper() == "TOTALS"]
    if totals.empty:
        raise ValueError("No TOTALS row found.")
    ssc_total_rf = float(totals["SSC_RF"].values[0])
    lec_total_rf = float(totals["LEC_RF"].values[0])

    ssc_total_raw = float(totals["SSC_Raw"].values[0])
    lec_total_raw = float(totals["LEC_Raw"].values[0])

    if ssc_total_rf <= 0 or lec_total_rf <= 0:
        raise ValueError("TOTALS RF values must be > 0 to infer corpus size.")

    ssc_words = ssc_total_raw * (1000.0 / ssc_total_rf)
    lec_words = lec_total_raw * (1000.0 / lec_total_rf)
    return ssc_words, lec_words


def g2_log_likelihood(a: float, b: float, ssc_words: float, lec_words: float) -> float:
    """
    2x2 log-likelihood ratio statistic (G^2).

    Table:
        a = count in SSC
        b = count in LEC
        c = other words in SSC
        d = other words in LEC
    """
    c = ssc_words - a
    d = lec_words - b

    # Expected values for each cell
    total_words = ssc_words + lec_words
    if total_words <= 0:
        return 0.0

    e1 = (a + b) * ssc_words / total_words
    e2 = (a + b) * lec_words / total_words
    e3 = (c + d) * ssc_words / total_words
    e4 = (c + d) * lec_words / total_words

    g2 = 0.0
    if e1 > 0 and a > 0:
        g2 += a * math.log(a / e1)
    if e2 > 0 and b > 0:
        g2 += b * math.log(b / e2)
    if e3 > 0 and c > 0:
        g2 += c * math.log(c / e3)
    if e4 > 0 and d > 0:
        g2 += d * math.log(d / e4)

    return 2.0 * g2


def log_ratio_log2(ssc_rf: float, lec_rf: float, zero_floor: float = 1e-5) -> float:
    """
    Log ratio effect size (log2). Uses RF (per 1,000 words).
    Mirrors the Copilot approach by applying a small floor to avoid division by zero.
    """
    s = float(ssc_rf)
    l = float(lec_rf)

    if s > 0 and l > 0:
        return math.log2(s / l)
    if s > 0 and l <= 0:
        return math.log2(s / zero_floor)
    if s <= 0 and l > 0:
        return -math.log2(l / zero_floor)
    return 0.0


def p_to_marker(p: float) -> str:
    """
    Significance marker used in plots.
    Keeps the spirit of your plotting script, but makes the thresholds explicit.

    **** p < 0.0001
    ***  p < 0.001
    *    p < 0.05
    ns   otherwise
    """
    if p < 1e-4:
        return "****"
    if p < 1e-3:
        return "***"
    if p < 0.05:
        return "*"
    return "ns"


def calculate_statistics(
    df: pd.DataFrame,
    ssc_words: Optional[float] = None,
    lec_words: Optional[float] = None,
) -> pd.DataFrame:
    """
    Compute G^2, p-values, log ratios and interpretation strings for each Vehicle_group.
    """
    df = df.copy()

    # Normalise Vehicle_group to strings
    df["Vehicle_group"] = df["Vehicle_group"].astype(str)

    # Corpus sizes
    if ssc_words is None or lec_words is None:
        ssc_words, lec_words = infer_corpus_sizes_from_totals(df)

    # Safety
    if ssc_words <= 0 or lec_words <= 0:
        raise ValueError("Corpus sizes must be > 0.")

    results: List[dict] = []

    for _, row in df.iterrows():
        vg = str(row["Vehicle_group"])
        if vg.upper() == "TOTALS":
            continue

        a = float(row["SSC_Raw"])
        b = float(row["LEC_Raw"])
        ssc_rf = float(row["SSC_RF"])
        lec_rf = float(row["LEC_RF"])

        g2 = g2_log_likelihood(a, b, ssc_words, lec_words)
        p = float(chi2.sf(g2, 1))
        p = min(p, 0.9999)

        lr = log_ratio_log2(ssc_rf, lec_rf, zero_floor=1e-5)

        if lr > 0:
            interp = f"SSC uses {2 ** lr:.1f}× more"
        elif lr < 0:
            interp = f"LEC uses {2 ** abs(lr):.1f}× more"
        else:
            interp = "No difference"

        results.append(
            dict(
                Vehicle_group=vg,
                SSC_Raw=a,
                SSC_RF=ssc_rf,
                LEC_Raw=b,
                LEC_RF=lec_rf,
                Log_Likelihood=g2,
                p_value=p,
                Significant=p < 0.05,
                Sig_Marker=p_to_marker(p),
                Log_Ratio=lr,
                Interpretation=interp,
            )
        )

    # Add TOTALS back (if present) for convenience
    totals = df.loc[df["Vehicle_group"].str.upper() == "TOTALS"]
    if not totals.empty:
        t = totals.iloc[0]
        a_total = float(t["SSC_Raw"])
        b_total = float(t["LEC_Raw"])
        g2_total = g2_log_likelihood(a_total, b_total, ssc_words, lec_words)
        p_total = float(chi2.sf(g2_total, 1))
        p_total = min(p_total, 0.9999)

        lr_total = log_ratio_log2(float(t["SSC_RF"]), float(t["LEC_RF"]), zero_floor=1e-5)
        interp_total = f"SSC uses {2 ** lr_total:.1f}× more overall" if lr_total >= 0 else f"LEC uses {2 ** abs(lr_total):.1f}× more overall"

        results.append(
            dict(
                Vehicle_group="TOTALS",
                SSC_Raw=a_total,
                SSC_RF=float(t["SSC_RF"]),
                LEC_Raw=b_total,
                LEC_RF=float(t["LEC_RF"]),
                Log_Likelihood=g2_total,
                p_value=p_total,
                Significant=p_total < 0.05,
                Sig_Marker=p_to_marker(p_total),
                Log_Ratio=lr_total,
                Interpretation=interp_total,
            )
        )

    out = pd.DataFrame(results)

    # Formatting helpers
    out["p_value_formatted"] = out["p_value"].apply(lambda p: "<0.0001" if p < 1e-4 else f"{p:.4f}")
    out["Log_Likelihood"] = out["Log_Likelihood"].round(2)
    out["Log_Ratio"] = out["Log_Ratio"].round(2)

    return out


# ----------------------------
# Plotting
# ----------------------------

def plot_diverging_log_ratio(df_stats: pd.DataFrame, outpath: Path) -> None:
    """
    Diverging horizontal bar chart of log ratio by vehicle group.
    Positive = more in SSC; negative = more in LEC.
    """
    df = df_stats.copy()

    # Put TOTALS last if present
    df_no_totals = df[df["Vehicle_group"].str.upper() != "TOTALS"]
    df_totals = df[df["Vehicle_group"].str.upper() == "TOTALS"]
    df = pd.concat([df_no_totals, df_totals], ignore_index=True)

    # Sort by log ratio (ascending) for diverging bars
    df = df.sort_values("Log_Ratio", ascending=True).reset_index(drop=True)

    y = np.arange(len(df))
    x = df["Log_Ratio"].values

    # Colors: blue for SSC>LEC, orange for LEC>SSC, special for TOTALS
    colors = []
    for vg, lr in zip(df["Vehicle_group"], x):
        if vg.upper() == "TOTALS":
            colors.append("#E69F00")  # highlight totals
        else:
            colors.append("#0072B2" if lr >= 0 else "#D55E00")

    # Dynamic x-limits with padding
    max_abs = max(1.0, float(np.max(np.abs(x))) if len(x) else 1.0)
    pad = 0.6
    xlim = (-(max_abs + pad), (max_abs + pad))

    fig, ax = plt.subplots(figsize=(12, 10))
    ax.barh(y, x, color=colors)

    ax.axvline(0, color="black", linewidth=1)

    ax.set_yticks(y)
    ax.set_yticklabels(df["Vehicle_group"])
    ax.set_xlabel("Log Ratio (positive = more in SSC, negative = more in LEC)")
    ax.set_title("Metaphor Vehicle Group Usage: Log Ratio Comparison between SSC and LEC")
    ax.set_xlim(*xlim)
    ax.grid(axis="x", linestyle="--", alpha=0.4)

    # Bold TOTALS label if present
    for tick in ax.get_yticklabels():
        if tick.get_text().upper() == "TOTALS":
            tick.set_fontweight("bold")

    # Value labels + significance markers
    marker_x_right = xlim[1] - 0.15
    marker_x_left = xlim[0] + 0.15

    for i, (lr, marker) in enumerate(zip(df["Log_Ratio"].values, df["Sig_Marker"].values)):
        label = f"{lr:.2f}"
        if lr >= 0:
            ax.text(lr + 0.05, i, label, va="center", ha="left", fontsize=9)
            ax.text(marker_x_right, i, marker, va="center", ha="right", fontweight="bold", fontsize=10)
        else:
            ax.text(lr - 0.05, i, label, va="center", ha="right", fontsize=9)
            ax.text(marker_x_left, i, marker, va="center", ha="left", fontweight="bold", fontsize=10)

    # Footnote
    fig.text(0.5, 0.01, "**** p<0.0001, *** p<0.001, * p<0.05, ns = not significant",
             ha="center", fontsize=10)

    fig.tight_layout(rect=(0, 0.02, 1, 1))
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def plot_bubble_significance(df_stats: pd.DataFrame, outpath: Path) -> None:
    """
    Bubble plot: x = log ratio, y = vehicle group (categorical),
    bubble size ~ log-likelihood, outline indicates non-significance.
    """
    df = df_stats.copy()

    # Put TOTALS last if present
    df_no_totals = df[df["Vehicle_group"].str.upper() != "TOTALS"]
    df_totals = df[df["Vehicle_group"].str.upper() == "TOTALS"]
    df = pd.concat([df_no_totals, df_totals], ignore_index=True)

    # Sort so the y-axis matches the bar plot ordering (by log ratio)
    df = df.sort_values("Log_Ratio", ascending=True).reset_index(drop=True)

    # Bubble sizes (normalised)
    ll = df["Log_Likelihood"].astype(float).values
    ll_max = float(np.max(ll)) if len(ll) and np.max(ll) > 0 else 1.0
    min_size = 80.0
    max_size = 1500.0
    bubble_sizes = min_size + (ll / ll_max) * (max_size - min_size)

    # Colors similar to bar plot
    colors = []
    for vg, lr in zip(df["Vehicle_group"], df["Log_Ratio"].values):
        if vg.upper() == "TOTALS":
            colors.append("#E69F00")
        else:
            colors.append("#0072B2" if lr >= 0 else "#D55E00")

    # Map vehicle groups to numeric y positions
    y = np.arange(len(df))

    max_abs = max(1.0, float(np.max(np.abs(df["Log_Ratio"].values))) if len(df) else 1.0)
    pad = 0.6
    xlim = (-(max_abs + pad), (max_abs + pad))

    fig, ax = plt.subplots(figsize=(12, 10))

    # Significant points (filled)
    sig_mask = df["Significant"].astype(bool).values
    ax.scatter(df.loc[sig_mask, "Log_Ratio"], y[sig_mask],
               s=bubble_sizes[sig_mask], c=np.array(colors, dtype=object)[sig_mask],
               alpha=0.7, edgecolors="none")

    # Non-significant points (hollow with red outline)
    ns_mask = ~sig_mask
    if np.any(ns_mask):
        ax.scatter(df.loc[ns_mask, "Log_Ratio"], y[ns_mask],
                   s=bubble_sizes[ns_mask], facecolors="none",
                   edgecolors="red", linewidths=2, alpha=0.9)

    ax.axvline(0, color="black", linewidth=1, alpha=0.5)

    ax.set_yticks(y)
    ax.set_yticklabels(df["Vehicle_group"])
    ax.set_xlabel("Log Ratio (effect size)")
    ax.set_title("Metaphor Usage: Effect Size vs. Statistical Significance")
    ax.set_xlim(*xlim)
    ax.grid(alpha=0.3)

    # Bold TOTALS label if present
    for tick in ax.get_yticklabels():
        if tick.get_text().upper() == "TOTALS":
            tick.set_fontweight("bold")

    # Annotate log ratio values (small)
    for i, lr in enumerate(df["Log_Ratio"].values):
        dx = 0.08 if lr >= 0 else -0.08
        ha = "left" if lr >= 0 else "right"
        ax.text(lr + dx, i, f"{lr:.2f}", va="center", ha=ha, fontsize=8)

    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


# ----------------------------
# CLI
# ----------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Compute log-likelihood + log ratio and generate plots.")
    parser.add_argument("--input", required=True, help="Path to CSV with vehicle group counts and RFs.")
    parser.add_argument("--outdir", default="outputs", help="Directory to write outputs to.")
    parser.add_argument("--ssc_words", type=float, default=None, help="SSC corpus size (words). Use if no TOTALS row.")
    parser.add_argument("--lec_words", type=float, default=None, help="LEC corpus size (words). Use if no TOTALS row.")
    parser.add_argument("--prefix", default="metaphor", help="Prefix for output filenames.")
    args = parser.parse_args()

    in_path = Path(args.input)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(in_path)

    stats = calculate_statistics(df, ssc_words=args.ssc_words, lec_words=args.lec_words)

    # Save results CSV
    results_csv = outdir / f"{args.prefix}_statistics_results.csv"
    stats.to_csv(results_csv, index=False)

    # Plots
    bar_png = outdir / f"{args.prefix}_log_ratio_comparison.png"
    bubble_png = outdir / f"{args.prefix}_significance_bubble_plot.png"

    plot_diverging_log_ratio(stats, bar_png)
    plot_bubble_significance(stats, bubble_png)

    print("Done.")
    print(f"- Results: {results_csv}")
    print(f"- Plot 1:   {bar_png}")
    print(f"- Plot 2:   {bubble_png}")


if __name__ == "__main__":
    main()
