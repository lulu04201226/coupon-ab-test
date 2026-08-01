from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib import font_manager
from scipy import stats
from statsmodels.stats.multitest import multipletests
from statsmodels.stats.power import TTestPower


PROJECT = Path(__file__).resolve().parents[1]
DATA_PATH = PROJECT / "data" / "test.xlsx"
TABLE_DIR = PROJECT / "outputs" / "tables"
CHART_DIR = PROJECT / "outputs" / "charts"

METRICS = {
    "requests": "请求量",
    "trips": "完成订单量",
    "gmv": "GMV",
    "coupon_cost": "优惠券成本",
    "roi": "绝对ROI",
    "completion_rate": "完成率",
    "cancel_rate": "取消率",
    "aov": "客单价",
}


def setup() -> None:
    TABLE_DIR.mkdir(parents=True, exist_ok=True)
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid")
    font_path = Path(r"C:\Windows\Fonts\msyh.ttc")
    if font_path.exists():
        font_manager.fontManager.addfont(str(font_path))
        plt.rcParams["font.family"] = "Microsoft YaHei"
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 130
    plt.rcParams["savefig.dpi"] = 180


def load_and_prepare() -> pd.DataFrame:
    data = pd.read_excel(DATA_PATH)
    required = {
        "date",
        "group",
        "requests",
        "gmv",
        "coupon per trip",
        "trips",
        "canceled requests",
    }
    missing = required - set(data.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")
    data["date"] = pd.to_datetime(data["date"])
    data = data.sort_values(["date", "group"]).reset_index(drop=True)
    data["coupon_cost"] = data["coupon per trip"] * data["trips"]
    data["roi"] = data["gmv"] / data["coupon_cost"]
    data["completion_rate"] = data["trips"] / data["requests"]
    data["cancel_rate"] = data["canceled requests"] / data["requests"]
    data["aov"] = data["gmv"] / data["trips"]
    return data


def validate_structure(data: pd.DataFrame) -> dict:
    date_group_counts = data.groupby("date")["group"].nunique()
    pair_complete = bool((date_group_counts == 2).all())
    group_dates = data.groupby("group")["date"].agg(["min", "max", "nunique"])
    validation = {
        "rows": int(len(data)),
        "columns": int(data.shape[1]),
        "missing_cells": int(data.isna().sum().sum()),
        "duplicate_rows": int(data.duplicated().sum()),
        "unique_dates": int(data["date"].nunique()),
        "group_counts": {k: int(v) for k, v in data["group"].value_counts().items()},
        "complete_date_pairs": pair_complete,
        "date_start": data["date"].min().strftime("%Y-%m-%d"),
        "date_end": data["date"].max().strftime("%Y-%m-%d"),
        "group_date_summary": group_dates.astype(str).to_dict(),
    }
    (TABLE_DIR / "data_validation.json").write_text(
        json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return validation


def group_summary(data: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for group, frame in data.groupby("group"):
        requests = frame["requests"].sum()
        trips = frame["trips"].sum()
        canceled = frame["canceled requests"].sum()
        gmv = frame["gmv"].sum()
        cost = frame["coupon_cost"].sum()
        rows.append(
            {
                "group": group,
                "days": len(frame),
                "requests_total": requests,
                "trips_total": trips,
                "gmv_total": gmv,
                "coupon_cost_total": cost,
                "completion_rate_weighted": trips / requests,
                "cancel_rate_weighted": canceled / requests,
                "aov_weighted": gmv / trips,
                "roi_aggregate": gmv / cost,
                "coupon_per_trip_mean": frame["coupon per trip"].mean(),
            }
        )
    result = pd.DataFrame(rows).set_index("group").loc[["control", "experiment"]]
    result.to_csv(TABLE_DIR / "group_summary.csv", encoding="utf-8-sig")
    return result


def paired_tests(data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    wide = data.pivot(index="date", columns="group")
    rows = []
    for metric, label in METRICS.items():
        control = wide[metric]["control"].astype(float)
        experiment = wide[metric]["experiment"].astype(float)
        diff = experiment - control
        n = len(diff)
        mean_diff = diff.mean()
        sd_diff = diff.std(ddof=1)
        se = sd_diff / np.sqrt(n)
        t_critical = stats.t.ppf(0.975, df=n - 1)
        ci_low = mean_diff - t_critical * se
        ci_high = mean_diff + t_critical * se
        paired_t = stats.ttest_rel(experiment, control)
        try:
            wilcoxon = stats.wilcoxon(diff, alternative="two-sided", zero_method="wilcox")
            wilcoxon_p = wilcoxon.pvalue
        except ValueError:
            wilcoxon_p = np.nan
        shapiro_p = stats.shapiro(diff).pvalue
        dz = mean_diff / sd_diff if sd_diff else np.nan
        relative_diff = mean_diff / control.mean()
        rows.append(
            {
                "metric": metric,
                "metric_cn": label,
                "n_pairs": n,
                "control_mean": control.mean(),
                "experiment_mean": experiment.mean(),
                "mean_diff": mean_diff,
                "relative_diff": relative_diff,
                "ci_low": ci_low,
                "ci_high": ci_high,
                "t_stat": paired_t.statistic,
                "paired_t_p": paired_t.pvalue,
                "wilcoxon_p": wilcoxon_p,
                "shapiro_diff_p": shapiro_p,
                "cohen_dz": dz,
            }
        )
    results = pd.DataFrame(rows)
    results["paired_t_p_holm"] = multipletests(
        results["paired_t_p"], alpha=0.05, method="holm"
    )[1]
    results["paired_t_significant_holm"] = results["paired_t_p_holm"] < 0.05
    results.to_csv(TABLE_DIR / "paired_test_results.csv", index=False, encoding="utf-8-sig")

    # Deliberately compare the wrong and right analytical structures to document
    # why ignoring the day pairing loses precision.
    control_gmv = data.loc[data["group"] == "control", "gmv"].to_numpy()
    experiment_gmv = data.loc[data["group"] == "experiment", "gmv"].to_numpy()
    independent = stats.ttest_ind(experiment_gmv, control_gmv, equal_var=False)
    paired = stats.ttest_rel(experiment_gmv, control_gmv)
    reversal = pd.DataFrame(
        [
            {
                "method": "Welch independent-samples t-test",
                "t_stat": independent.statistic,
                "p_value": independent.pvalue,
                "conclusion_alpha_005": "not significant"
                if independent.pvalue >= 0.05
                else "significant",
            },
            {
                "method": "Paired t-test by date",
                "t_stat": paired.statistic,
                "p_value": paired.pvalue,
                "conclusion_alpha_005": "not significant"
                if paired.pvalue >= 0.05
                else "significant",
            },
        ]
    )
    reversal.to_csv(TABLE_DIR / "gmv_method_comparison.csv", index=False, encoding="utf-8-sig")
    return results, wide


def robustness_checks(data: pd.DataFrame, wide: pd.DataFrame) -> pd.DataFrame:
    diff = (wide["gmv"]["experiment"] - wide["gmv"]["control"]).astype(float)
    rng = np.random.default_rng(20260731)
    boot_means = np.array(
        [rng.choice(diff.to_numpy(), size=len(diff), replace=True).mean() for _ in range(20000)]
    )
    bootstrap_ci = np.quantile(boot_means, [0.025, 0.975])

    abs_z = np.abs(stats.zscore(diff, ddof=1))
    keep = abs_z < 2.5
    trimmed = diff.loc[keep]
    trimmed_t = stats.ttest_1samp(trimmed, popmean=0)

    half = len(diff) // 2
    early = diff.iloc[:half]
    late = diff.iloc[half:]
    result = pd.DataFrame(
        [
            {
                "check": "bootstrap_20000",
                "n": len(diff),
                "mean_diff": diff.mean(),
                "ci_low": bootstrap_ci[0],
                "ci_high": bootstrap_ci[1],
                "p_value": np.nan,
            },
            {
                "check": "exclude_abs_z_ge_2.5",
                "n": len(trimmed),
                "mean_diff": trimmed.mean(),
                "ci_low": np.nan,
                "ci_high": np.nan,
                "p_value": trimmed_t.pvalue,
            },
            {
                "check": "first_half",
                "n": len(early),
                "mean_diff": early.mean(),
                "ci_low": np.nan,
                "ci_high": np.nan,
                "p_value": stats.ttest_1samp(early, 0).pvalue,
            },
            {
                "check": "second_half",
                "n": len(late),
                "mean_diff": late.mean(),
                "ci_low": np.nan,
                "ci_high": np.nan,
                "p_value": stats.ttest_1samp(late, 0).pvalue,
            },
        ]
    )
    result.to_csv(TABLE_DIR / "gmv_robustness_checks.csv", index=False, encoding="utf-8-sig")
    return result


def power_analysis(data: pd.DataFrame, test_results: pd.DataFrame) -> pd.DataFrame:
    gmv = test_results.loc[test_results["metric"] == "gmv"].iloc[0]
    effect = abs(gmv["cohen_dz"])
    n = int(gmv["n_pairs"])
    model = TTestPower()
    observed_power = model.power(effect_size=effect, nobs=n, alpha=0.05, alternative="two-sided")
    n_for_80 = model.solve_power(effect_size=effect, nobs=None, alpha=0.05, power=0.80)
    mde_dz = model.solve_power(effect_size=None, nobs=n, alpha=0.05, power=0.80)
    wide = data.pivot(index="date", columns="group")
    diff_sd = (wide["gmv"]["experiment"] - wide["gmv"]["control"]).std(ddof=1)
    control_mean = wide["gmv"]["control"].mean()
    mde_abs = mde_dz * diff_sd
    mde_relative = mde_abs / control_mean
    result = pd.DataFrame(
        [
            {
                "n_pairs": n,
                "observed_cohen_dz_abs": effect,
                "posthoc_power": observed_power,
                "pairs_needed_80_power": np.ceil(n_for_80),
                "mde_cohen_dz_80_power": mde_dz,
                "mde_gmv_absolute": mde_abs,
                "mde_gmv_relative": mde_relative,
                "alpha": 0.05,
                "target_power": 0.80,
            }
        ]
    )
    result.to_csv(TABLE_DIR / "power_analysis.csv", index=False, encoding="utf-8-sig")
    return result


def economics_summary(data: pd.DataFrame, summary: pd.DataFrame) -> pd.DataFrame:
    control = summary.loc["control"]
    experiment = summary.loc["experiment"]
    delta_gmv = experiment["gmv_total"] - control["gmv_total"]
    delta_cost = experiment["coupon_cost_total"] - control["coupon_cost_total"]
    incremental_roi = delta_gmv / delta_cost if delta_cost != 0 else np.nan
    result = pd.DataFrame(
        [
            {
                "delta_gmv": delta_gmv,
                "delta_coupon_cost": delta_cost,
                "incremental_roi_gmv_per_coupon": incremental_roi,
                "net_incremental_value_gmv_minus_coupon": delta_gmv - delta_cost,
            }
        ]
    )
    result.to_csv(TABLE_DIR / "economics_summary.csv", index=False, encoding="utf-8-sig")
    return result


def coupon_trend(data: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for group, frame in data.groupby("group"):
        x = np.arange(len(frame))
        slope, intercept, r, p, _ = stats.linregress(x, frame["coupon per trip"])
        rows.append(
            {
                "group": group,
                "daily_slope": slope,
                "r_squared": r**2,
                "slope_p": p,
                "start_fitted": intercept,
                "end_fitted": intercept + slope * (len(frame) - 1),
            }
        )
    result = pd.DataFrame(rows)
    result.to_csv(TABLE_DIR / "coupon_trend.csv", index=False, encoding="utf-8-sig")
    return result


def save_charts(data: pd.DataFrame, tests: pd.DataFrame, wide: pd.DataFrame) -> None:
    palette = {"control": "#64748B", "experiment": "#2563EB"}

    fig, axes = plt.subplots(2, 2, figsize=(12, 8), sharex=True)
    chart_specs = [
        ("requests", "每日请求量"),
        ("trips", "每日完成订单量"),
        ("gmv", "每日GMV"),
        ("coupon per trip", "每单优惠券金额"),
    ]
    for ax, (metric, title) in zip(axes.flat, chart_specs):
        sns.lineplot(
            data=data,
            x="date",
            y=metric,
            hue="group",
            palette=palette,
            linewidth=2,
            ax=ax,
        )
        ax.set_title(title)
        ax.set_xlabel("")
        ax.set_ylabel("")
        ax.legend(title="", labels=["对照组", "实验组"])
        ax.tick_params(axis="x", rotation=30)
    fig.suptitle("两组核心指标的每日趋势（2019年1月1日—29日）", fontsize=15, fontweight="bold")
    fig.tight_layout()
    fig.savefig(CHART_DIR / "01_daily_trends.png", bbox_inches="tight")
    plt.close(fig)

    gmv_diff = (wide["gmv"]["experiment"] - wide["gmv"]["control"]).rename("gmv_diff")
    fig, ax = plt.subplots(figsize=(11, 4.8))
    colors = np.where(gmv_diff < 0, "#DC2626", "#16A34A")
    ax.bar(gmv_diff.index, gmv_diff.values, color=colors, width=0.8)
    ax.axhline(0, color="#334155", linewidth=1)
    ax.axhline(gmv_diff.mean(), color="#7C3AED", linestyle="--", linewidth=2, label=f"日均差值 {gmv_diff.mean():,.0f}")
    ax.set_title("每日GMV配对差值（实验组－对照组）")
    ax.set_ylabel("GMV差值")
    ax.set_xlabel("")
    ax.tick_params(axis="x", rotation=30)
    ax.legend()
    fig.tight_layout()
    fig.savefig(CHART_DIR / "02_gmv_paired_difference.png", bbox_inches="tight")
    plt.close(fig)

    display_metrics = ["requests", "trips", "gmv", "coupon_cost", "completion_rate", "cancel_rate", "aov", "roi"]
    plot_data = tests.set_index("metric").loc[display_metrics].copy()
    fig, ax = plt.subplots(figsize=(10, 5.5))
    values = plot_data["relative_diff"] * 100
    colors = []
    for metric, value in zip(plot_data.index, values):
        if metric == "cancel_rate":
            favorable = value < 0
        elif metric == "coupon_cost":
            favorable = value < 0
        else:
            favorable = value > 0
        colors.append("#16A34A" if favorable else "#DC2626")
    bars = ax.barh(plot_data["metric_cn"], values, color=colors)
    ax.axvline(0, color="#334155", linewidth=1)
    ax.set_title("实验组相对对照组的日均指标变化")
    ax.set_xlabel("相对变化（%）")
    for bar, value in zip(bars, values):
        ax.text(
            value + (0.08 if value >= 0 else -0.08),
            bar.get_y() + bar.get_height() / 2,
            f"{value:+.2f}%",
            va="center",
            ha="left" if value >= 0 else "right",
            fontsize=9,
        )
    fig.tight_layout()
    fig.savefig(CHART_DIR / "03_metric_lifts.png", bbox_inches="tight")
    plt.close(fig)

    method = pd.read_csv(TABLE_DIR / "gmv_method_comparison.csv")
    fig, ax = plt.subplots(figsize=(8.5, 4.6))
    labels = ["独立样本检验", "按日期配对检验"]
    pvals = method["p_value"].to_numpy()
    heights = -np.log10(pvals)
    bars = ax.bar(labels, heights, color=["#94A3B8", "#2563EB"], width=0.55)
    ax.axhline(-np.log10(0.05), color="#DC2626", linestyle="--", label="显著性阈值 p=0.05")
    ax.set_ylabel("-log10(p值)")
    ax.set_title("GMV检验方法对显著性判断的影响")
    ax.legend()
    for b, p in zip(bars, pvals):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.08, f"p={p:.4f}", ha="center")
    fig.tight_layout()
    fig.savefig(CHART_DIR / "04_method_reversal.png", bbox_inches="tight")
    plt.close(fig)

    gmv = tests.loc[tests["metric"] == "gmv"].iloc[0]
    model = TTestPower()
    sample_sizes = np.arange(5, 61)
    power = [model.power(abs(gmv["cohen_dz"]), int(n), 0.05) for n in sample_sizes]
    mde = [model.solve_power(None, int(n), 0.05, 0.80) for n in sample_sizes]
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.8))
    axes[0].plot(sample_sizes, power, color="#2563EB", linewidth=2)
    axes[0].axhline(0.8, color="#DC2626", linestyle="--")
    axes[0].axvline(29, color="#64748B", linestyle=":")
    axes[0].set_title("观测效应量下的检验功效")
    axes[0].set_xlabel("配对天数")
    axes[0].set_ylabel("Power")
    axes[0].set_ylim(0, 1.03)
    axes[1].plot(sample_sizes, mde, color="#7C3AED", linewidth=2)
    axes[1].axvline(29, color="#64748B", linestyle=":")
    axes[1].set_title("80%功效下的最小可检测效应")
    axes[1].set_xlabel("配对天数")
    axes[1].set_ylabel("Cohen's dz")
    fig.tight_layout()
    fig.savefig(CHART_DIR / "05_power_and_mde.png", bbox_inches="tight")
    plt.close(fig)

    coupon = data.pivot(index="date", columns="group", values="coupon per trip")
    fig, ax = plt.subplots(figsize=(10.5, 4.8))
    ax.plot(coupon.index, coupon["control"], label="对照组", color=palette["control"], linewidth=2)
    ax.plot(coupon.index, coupon["experiment"], label="实验组", color=palette["experiment"], linewidth=2)
    ax.fill_between(
        coupon.index,
        coupon["control"],
        coupon["experiment"],
        color="#93C5FD",
        alpha=0.3,
        label="组间券额差",
    )
    ax.set_title("每单优惠券金额随实验时间递减")
    ax.set_ylabel("每单优惠券金额")
    ax.set_xlabel("")
    ax.tick_params(axis="x", rotation=30)
    ax.legend()
    fig.tight_layout()
    fig.savefig(CHART_DIR / "06_coupon_intensity_trend.png", bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    setup()
    data = load_and_prepare()
    validation = validate_structure(data)
    summary = group_summary(data)
    tests, wide = paired_tests(data)
    robustness = robustness_checks(data, wide)
    power = power_analysis(data, tests)
    economics = economics_summary(data, summary)
    coupon = coupon_trend(data)
    save_charts(data, tests, wide)

    headline = {
        "validation": validation,
        "gmv": tests.loc[tests["metric"] == "gmv"].iloc[0].to_dict(),
        "power": power.iloc[0].to_dict(),
        "economics": economics.iloc[0].to_dict(),
        "coupon_trend": coupon.to_dict(orient="records"),
        "robustness": robustness.to_dict(orient="records"),
    }
    (TABLE_DIR / "headline_results.json").write_text(
        json.dumps(headline, ensure_ascii=False, indent=2, default=float), encoding="utf-8"
    )
    print(json.dumps(headline, ensure_ascii=False, indent=2, default=float))


if __name__ == "__main__":
    main()
