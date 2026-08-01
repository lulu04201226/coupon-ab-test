"""网约车优惠券 A/B 测试交互式 Dashboard。"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy import stats


st.set_page_config(
    page_title="网约车优惠券 A/B 测试",
    page_icon="🚕",
    layout="wide",
)

DATA_PATH = Path(__file__).resolve().parents[1] / "data" / "test.xlsx"

GROUP_LABELS = {"control": "对照组", "experiment": "实验组"}
GROUP_COLORS = {"对照组": "#64748B", "实验组": "#2563EB"}
POSITIVE_COLOR = "#16A34A"
NEGATIVE_COLOR = "#DC2626"
NEUTRAL_COLOR = "#64748B"

BASE_METRICS = {
    "requests": "请求数",
    "trips": "完成订单数",
    "gmv": "GMV",
    "coupon per trip": "每单优惠券金额",
}

ANALYSIS_METRICS = {
    "requests": {"label": "请求数", "kind": "number", "better": "up"},
    "trips": {"label": "完成订单数", "kind": "number", "better": "up"},
    "gmv": {"label": "GMV", "kind": "currency", "better": "up"},
    "coupon per trip": {
        "label": "每单优惠券金额",
        "kind": "currency_small",
        "better": "down",
    },
    "completion_rate": {"label": "完成率", "kind": "rate", "better": "up"},
    "cancel_rate": {"label": "取消率", "kind": "rate", "better": "down"},
    "asp": {"label": "客单价", "kind": "currency", "better": "up"},
    "coupon_cost": {"label": "优惠券成本", "kind": "currency", "better": "down"},
}

KPI_METRICS = ["gmv", "completion_rate", "cancel_rate", "asp", "coupon_cost"]
DIFF_METRICS = {
    "gmv": "GMV",
    "completion_rate": "完成率",
    "cancel_rate": "取消率",
    "asp": "客单价",
    "coupon_cost": "优惠券成本",
}

# 全周期报告参考值，仅用于页面说明；所有可视化与检验均由筛选数据实时计算。
FULL_PERIOD_REFERENCE = {
    "GMV": "日均差 -6,111，配对 t 检验 p=0.0002",
    "客单价": "下降 1.63%，p<0.001",
    "完成率": "提高 0.20 个百分点，p=0.0010",
    "取消率": "下降 0.21 个百分点，p<0.001",
    "优惠券成本": "增加 0.77%，p=0.0103",
}


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    """读取数据、校验字段并计算衍生指标。"""
    required = {
        "date",
        "group",
        "requests",
        "gmv",
        "coupon per trip",
        "trips",
        "canceled requests",
    }
    data = pd.read_excel(path)
    missing = required - set(data.columns)
    if missing:
        raise ValueError(f"数据缺少必要字段：{', '.join(sorted(missing))}")

    data = data.copy()
    data["date"] = pd.to_datetime(data["date"])
    data = data.sort_values(["date", "group"]).reset_index(drop=True)
    data["completion_rate"] = data["trips"] / data["requests"]
    data["cancel_rate"] = data["canceled requests"] / data["requests"]
    data["asp"] = data["gmv"] / data["trips"]
    data["coupon_cost"] = data["coupon per trip"] * data["trips"]
    return data


def paired_frame(data: pd.DataFrame, metric: str) -> pd.DataFrame:
    """生成同日期实验组与对照组的一一配对数据。"""
    pair = data.pivot(index="date", columns="group", values=metric)
    pair = pair.reindex(columns=["control", "experiment"]).dropna()
    pair["difference"] = pair["experiment"] - pair["control"]
    return pair.reset_index()


def paired_test(data: pd.DataFrame, metric: str) -> dict[str, float]:
    """返回配对均值差、t 检验、置信区间和相对变化。"""
    pair = paired_frame(data, metric)
    n = len(pair)
    control_mean = pair["control"].mean() if n else np.nan
    experiment_mean = pair["experiment"].mean() if n else np.nan
    mean_diff = pair["difference"].mean() if n else np.nan
    relative_change = mean_diff / control_mean if n and control_mean != 0 else np.nan

    if n < 2:
        return {
            "n_pairs": n,
            "control_mean": control_mean,
            "experiment_mean": experiment_mean,
            "mean_diff": mean_diff,
            "relative_change": relative_change,
            "t_stat": np.nan,
            "p_value": np.nan,
            "ci_low": np.nan,
            "ci_high": np.nan,
        }

    result = stats.ttest_rel(pair["experiment"], pair["control"])
    diff_std = pair["difference"].std(ddof=1)
    standard_error = diff_std / np.sqrt(n)
    critical_value = stats.t.ppf(0.975, df=n - 1)
    return {
        "n_pairs": n,
        "control_mean": control_mean,
        "experiment_mean": experiment_mean,
        "mean_diff": mean_diff,
        "relative_change": relative_change,
        "t_stat": float(result.statistic),
        "p_value": float(result.pvalue),
        "ci_low": mean_diff - critical_value * standard_error,
        "ci_high": mean_diff + critical_value * standard_error,
    }


def format_p_value(value: float) -> str:
    if pd.isna(value):
        return "样本不足"
    if value < 0.001:
        return "p<0.001"
    return f"p={value:.4f}"


def format_value(value: float, kind: str) -> str:
    if pd.isna(value):
        return "—"
    if kind == "rate":
        return f"{value:.2%}"
    if kind == "currency_small":
        return f"{value:,.2f}"
    if kind == "currency":
        return f"{value:,.0f}"
    return f"{value:,.0f}"


def format_difference(value: float, kind: str) -> str:
    if pd.isna(value):
        return "—"
    if kind == "rate":
        return f"{value * 100:+.2f}pp"
    if kind == "currency_small":
        return f"{value:+,.3f}"
    return f"{value:+,.0f}"


def significance_color(result: dict[str, float], metric: str) -> str:
    """显著且方向有利用绿色，显著且方向不利用红色。"""
    p_value = result["p_value"]
    if pd.isna(p_value) or p_value >= 0.05:
        return NEUTRAL_COLOR
    better = ANALYSIS_METRICS[metric]["better"]
    favorable = result["mean_diff"] > 0 if better == "up" else result["mean_diff"] < 0
    return POSITIVE_COLOR if favorable else NEGATIVE_COLOR


def render_kpi_card(metric: str, result: dict[str, float]) -> None:
    spec = ANALYSIS_METRICS[metric]
    accent = significance_color(result, metric)
    relative = result["relative_change"]
    relative_text = "—" if pd.isna(relative) else f"{relative:+.2%}"
    p_text = format_p_value(result["p_value"])
    significance = (
        "显著" if not pd.isna(result["p_value"]) and result["p_value"] < 0.05 else "不显著"
    )
    st.markdown(
        f"""
        <div class="kpi-card" style="border-top: 4px solid {accent};">
          <div class="kpi-title">{spec['label']}</div>
          <div class="kpi-row"><span>对照组</span><b>{format_value(result['control_mean'], spec['kind'])}</b></div>
          <div class="kpi-row"><span>实验组</span><b>{format_value(result['experiment_mean'], spec['kind'])}</b></div>
          <div class="kpi-change">{format_difference(result['mean_diff'], spec['kind'])} · {relative_text}</div>
          <div class="kpi-p" style="color:{accent};">{p_text} · {significance}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def trend_chart(data: pd.DataFrame, metric: str) -> go.Figure:
    chart_data = data.copy()
    chart_data["组别"] = chart_data["group"].map(GROUP_LABELS)
    fig = px.line(
        chart_data,
        x="date",
        y=metric,
        color="组别",
        markers=True,
        color_discrete_map=GROUP_COLORS,
        labels={"date": "日期", metric: BASE_METRICS[metric]},
    )
    fig.update_traces(line={"width": 2.4}, marker={"size": 6})
    fig.update_layout(
        height=430,
        hovermode="x unified",
        legend_title_text="",
        margin={"l": 10, "r": 10, "t": 20, "b": 10},
    )
    return fig


def difference_chart(data: pd.DataFrame, metric: str) -> go.Figure:
    pair = paired_frame(data, metric)
    spec = ANALYSIS_METRICS[metric]
    is_rate = spec["kind"] == "rate"
    plot_values = pair["difference"] * 100 if is_rate else pair["difference"]
    control_values = pair["control"] * 100 if is_rate else pair["control"]
    experiment_values = pair["experiment"] * 100 if is_rate else pair["experiment"]
    mean_value = plot_values.mean()
    colors = np.where(plot_values >= 0, POSITIVE_COLOR, NEGATIVE_COLOR)

    fig = go.Figure()
    fig.add_bar(
        x=pair["date"],
        y=plot_values,
        marker_color=colors,
        name="每日配对差值",
        customdata=np.column_stack([control_values, experiment_values]),
        hovertemplate=(
            "日期：%{x|%Y-%m-%d}<br>差值：%{y:,.2f}<br>"
            "对照组：%{customdata[0]:,.2f}<br>实验组：%{customdata[1]:,.2f}<extra></extra>"
        ),
    )
    fig.add_hline(
        y=0,
        line_color="#334155",
        line_width=1,
    )
    fig.add_hline(
        y=mean_value,
        line_color="#7C3AED",
        line_dash="dash",
        line_width=2,
        annotation_text=f"日均差值 {mean_value:,.2f}",
        annotation_position="top left",
    )
    y_label = "差值（百分点）" if is_rate else "差值"
    fig.update_layout(
        height=430,
        xaxis_title="日期",
        yaxis_title=y_label,
        showlegend=False,
        margin={"l": 10, "r": 10, "t": 35, "b": 10},
    )
    return fig


def method_comparison(data: pd.DataFrame) -> tuple[pd.DataFrame, go.Figure]:
    pair = paired_frame(data, "gmv")
    if len(pair) < 2:
        comparison = pd.DataFrame(
            [
                {"检验方法": "Welch 独立样本 t 检验", "t统计量": np.nan, "p值": np.nan},
                {"检验方法": "按日期配对 t 检验", "t统计量": np.nan, "p值": np.nan},
            ]
        )
    else:
        welch = stats.ttest_ind(pair["experiment"], pair["control"], equal_var=False)
        paired = stats.ttest_rel(pair["experiment"], pair["control"])
        comparison = pd.DataFrame(
            [
                {
                    "检验方法": "Welch 独立样本 t 检验",
                    "t统计量": float(welch.statistic),
                    "p值": float(welch.pvalue),
                },
                {
                    "检验方法": "按日期配对 t 检验",
                    "t统计量": float(paired.statistic),
                    "p值": float(paired.pvalue),
                },
            ]
        )
    comparison["显著性判断"] = np.where(
        comparison["p值"] < 0.05, "显著", "不显著"
    )

    safe_p = comparison["p值"].clip(lower=1e-12)
    comparison["-log10(p)"] = -np.log10(safe_p)
    fig = px.bar(
        comparison,
        x="检验方法",
        y="-log10(p)",
        color="检验方法",
        text=comparison["p值"].map(format_p_value),
        color_discrete_sequence=["#94A3B8", "#2563EB"],
    )
    fig.add_hline(
        y=-np.log10(0.05),
        line_color=NEGATIVE_COLOR,
        line_dash="dash",
        annotation_text="显著性阈值 p=0.05",
        annotation_position="top left",
    )
    fig.update_traces(textposition="outside", hovertemplate="%{x}<br>-log10(p)：%{y:.2f}<extra></extra>")
    fig.update_layout(
        height=390,
        showlegend=False,
        xaxis_title="",
        yaxis_title="-log10(p值)",
        margin={"l": 10, "r": 10, "t": 30, "b": 10},
    )
    return comparison, fig


def statistical_results(data: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for metric, spec in ANALYSIS_METRICS.items():
        result = paired_test(data, metric)
        scale = 100 if spec["kind"] == "rate" else 1
        rows.append(
            {
                "指标": spec["label"],
                "单位": "百分点" if spec["kind"] == "rate" else "原始单位",
                "配对天数": result["n_pairs"],
                "对照组均值": result["control_mean"] * scale,
                "实验组均值": result["experiment_mean"] * scale,
                "均值差(实验-对照)": result["mean_diff"] * scale,
                "相对变化%": result["relative_change"] * 100,
                "t统计量": result["t_stat"],
                "p值": result["p_value"],
                "95%CI下限": result["ci_low"] * scale,
                "95%CI上限": result["ci_high"] * scale,
                "显著性": (
                    "显著"
                    if not pd.isna(result["p_value"]) and result["p_value"] < 0.05
                    else "不显著"
                ),
            }
        )
    return pd.DataFrame(rows)


st.markdown(
    """
    <style>
      .block-container {padding-top: 2rem; padding-bottom: 3rem; max-width: 1450px;}
      [data-testid="stMetric"] {background: white; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px;}
      .kpi-card {background: white; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px 15px; min-height: 190px; box-shadow: 0 1px 2px rgba(15,23,42,.04);}
      .kpi-title {font-size: 1.05rem; font-weight: 700; color: #0F172A; margin-bottom: 12px;}
      .kpi-row {display: flex; justify-content: space-between; color: #475569; font-size: .90rem; margin: 7px 0;}
      .kpi-row b {color: #0F172A;}
      .kpi-change {font-size: 1rem; font-weight: 700; color: #0F172A; margin-top: 13px;}
      .kpi-p {font-size: .82rem; font-weight: 700; margin-top: 7px;}
      .section-note {color: #64748B; font-size: .92rem; margin-top: -8px; margin-bottom: 10px;}
      div[data-testid="stDataFrame"] {border: 1px solid #E2E8F0; border-radius: 8px;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🚕 网约车优惠券 A/B 测试分析")
st.markdown(
    "基于 **2019-01-01 至 2019-01-29** 的日期级配对数据，比较提高优惠券力度后，"
    "实验组与对照组在增长、交易体验和补贴效率上的差异。两组同日期一一配对，"
    "因此配对检验是本项目的主要统计方法。"
)

try:
    full_data = load_data(DATA_PATH)
except (FileNotFoundError, ValueError) as exc:
    st.error(f"数据加载失败：{exc}")
    st.stop()

min_date = full_data["date"].min().date()
max_date = full_data["date"].max().date()

with st.sidebar:
    st.header("筛选条件")
    selected_dates = st.slider(
        "日期范围",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),
        format="YYYY-MM-DD",
    )
    st.caption("日期筛选将联动刷新全部 KPI、图表和统计检验。")
    st.divider()
    st.markdown("**数据口径**")
    st.caption("粒度：日期 × 实验组别")
    st.caption("来源：data/test.xlsx")
    st.caption("完整日期配对：29 对")

start_date, end_date = pd.Timestamp(selected_dates[0]), pd.Timestamp(selected_dates[1])
filtered = full_data.loc[full_data["date"].between(start_date, end_date)].copy()
pair_count = filtered["date"].nunique()

if pair_count < 2:
    st.warning("当前日期范围不足 2 个完整配对日：可以查看描述性数据，但无法计算 t 检验。")

st.subheader("核心指标概览")
st.markdown(
    f'<div class="section-note">当前范围：{start_date:%Y-%m-%d} 至 {end_date:%Y-%m-%d}，共 {pair_count} 个配对日。卡片变化均为“实验组 − 对照组”。</div>',
    unsafe_allow_html=True,
)

kpi_results = {metric: paired_test(filtered, metric) for metric in KPI_METRICS}
kpi_columns = st.columns(5)
for column, metric in zip(kpi_columns, KPI_METRICS):
    with column:
        render_kpi_card(metric, kpi_results[metric])

with st.expander("查看全周期报告参考结论", expanded=False):
    for label, text in FULL_PERIOD_REFERENCE.items():
        st.markdown(f"- **{label}：**{text}")
    st.markdown(
        "- **稳健性：**20,000 次 Bootstrap 的 GMV 差值 95% CI 为 [-8,909, -3,394]。"
    )
    st.markdown(
        "- **时间异质性：**前 14 天 GMV 均值差 +723（p=0.194）；后 15 天 -12,490（p<0.001）。"
    )

st.divider()
st.subheader("核心指标每日趋势")
trend_metric = st.selectbox(
    "选择趋势指标",
    options=list(BASE_METRICS),
    format_func=BASE_METRICS.get,
    index=2,
)
st.plotly_chart(trend_chart(filtered, trend_metric))

st.divider()
st.subheader("日期级配对差值")
st.markdown(
    '<div class="section-note">柱高表示同一天“实验组 − 对照组”；绿色为正、红色为负，虚线表示筛选范围内的日均差值。</div>',
    unsafe_allow_html=True,
)
diff_metric = st.selectbox(
    "选择配对差值指标",
    options=list(DIFF_METRICS),
    format_func=DIFF_METRICS.get,
    index=0,
)
st.plotly_chart(difference_chart(filtered, diff_metric))

st.divider()
st.subheader("为什么必须按日期配对？")
st.markdown(
    "独立样本检验把日期间的大盘波动当作组内噪声；配对检验先消除同一天共有的波动，"
    "因此能够识别被掩盖的组间差异。柱越高代表 p 值越小。"
)
comparison, comparison_fig = method_comparison(filtered)
chart_col, table_col = st.columns([1.55, 1])
with chart_col:
    st.plotly_chart(comparison_fig)
with table_col:
    display_comparison = comparison[["检验方法", "t统计量", "p值", "显著性判断"]].copy()
    st.dataframe(
        display_comparison,
        hide_index=True,
        width="stretch",
        column_config={
            "t统计量": st.column_config.NumberColumn(format="%.3f"),
            "p值": st.column_config.NumberColumn(format="%.4f"),
        },
    )
    if pair_count == full_data["date"].nunique():
        st.error(
            "全周期核心发现：Welch 检验 p=0.9387，错误地指向“无差异”；"
            "按日期配对后 p=0.0002，GMV 显著下降。"
        )

st.divider()
st.subheader("统计检验结果表")
st.markdown(
    '<div class="section-note">所有结果均基于当前筛选范围实时计算。点击列名可排序；率类指标的均值差与置信区间使用“百分点”。</div>',
    unsafe_allow_html=True,
)
results_table = statistical_results(filtered)
st.dataframe(
    results_table,
    hide_index=True,
    width="stretch",
    height=360,
    column_config={
        "对照组均值": st.column_config.NumberColumn(format="%.2f"),
        "实验组均值": st.column_config.NumberColumn(format="%.2f"),
        "均值差(实验-对照)": st.column_config.NumberColumn(format="%.2f"),
        "相对变化%": st.column_config.NumberColumn(format="%.2f%%"),
        "t统计量": st.column_config.NumberColumn(format="%.3f"),
        "p值": st.column_config.NumberColumn(format="%.4f"),
        "95%CI下限": st.column_config.NumberColumn(format="%.2f"),
        "95%CI上限": st.column_config.NumberColumn(format="%.2f"),
    },
)

st.caption(
    "说明：本项目使用日期级汇总数据，无法替代用户级随机化质量检查；筛选后的短周期结果仅用于探索，不应脱离样本量和业务周期单独决策。"
)
