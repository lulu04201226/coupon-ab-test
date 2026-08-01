from pathlib import Path

import nbformat as nbf


PROJECT = Path(__file__).resolve().parents[1]
NOTEBOOK_DIR = PROJECT / "notebooks"
NOTEBOOK_DIR.mkdir(parents=True, exist_ok=True)
OUT = NOTEBOOK_DIR / "01_didi_ab_test_analysis.ipynb"

nb = nbf.v4.new_notebook()
nb["metadata"] = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3"},
}

cells = [
    nbf.v4.new_markdown_cell(
        """# 网约车优惠券 A/B 测试分析

## tl;dr

- 实验组日均 GMV 比对照组低 **6,111（-1.26%）**。按日期配对后差异显著：配对 t 检验 **p=0.0002**，Wilcoxon **p=0.0018**。
- 若错误忽略日期配对，Welch 独立样本检验得到 **p=0.9387**。方法选择改变了结论，因为配对分析消除了两组共同的日期波动。
- 实验组优惠券成本增加约 **0.77%**，绝对 ROI 下降约 **2.41%**；补贴没有转化为 GMV 增长。
- 完成率提高约 **0.20 个百分点**、取消率降低约 **0.21 个百分点**，但不足以抵消 GMV 和客单价下降。
- GMV负向效果主要集中在实验后半程，且两组券额均随时间下降，因此处理强度并非恒定，推广前应控制券额路径并进行二次验证。
"""
    ),
    nbf.v4.new_markdown_cell(
        """## Context & Methods

### Decision

判断提高优惠券力度是否值得推广，并区分统计显著性与业务价值。

### Key Assumptions

- 数据为 29 天、对照组与实验组同日期对齐的汇总结果。
- 日期对齐与同步波动支持日期级配对分析，但缺少用户级分流日志，不能完全验证随机化机制。
- 原始数据未提供实验方案、灰度记录与样本量设计；相关流程在报告中标记为“方案补充”。
- ROI 仅表示 GMV 与优惠券成本的关系，不代表利润 ROI。

### Methods

- 描述性指标与补贴经济性
- 日期级配对 t 检验
- Wilcoxon 符号秩检验
- Bootstrap 置信区间与分阶段稳健性检查
- Holm 多重检验校正
- 功效、样本量和 MDE 敏感性分析
"""
    ),
    nbf.v4.new_markdown_cell("## Data\n\n### 1. 运行可复现分析"),
    nbf.v4.new_code_cell(
        """from pathlib import Path
import sys
import pandas as pd
from IPython.display import display, Image

PROJECT = Path.cwd().resolve().parent if Path.cwd().name == "notebooks" else Path.cwd().resolve()
if PROJECT.name != "didi_ab_test_project":
    PROJECT = Path.cwd().resolve() / "didi_ab_test_project"
sys.path.insert(0, str(PROJECT / "src"))

import run_analysis
run_analysis.main()
TABLE_DIR = PROJECT / "outputs" / "tables"
CHART_DIR = PROJECT / "outputs" / "charts"
"""
    ),
    nbf.v4.new_markdown_cell("### 2. 数据结构与质量"),
    nbf.v4.new_code_cell(
        """data = run_analysis.load_and_prepare()
validation = run_analysis.validate_structure(data)
validation"""
    ),
    nbf.v4.new_markdown_cell(
        """数据包含 58 行、29 个日期，两组各 29 行；无缺失和重复，所有日期均形成完整配对。"""
    ),
    nbf.v4.new_markdown_cell("## Results\n\n### 3. 两组总体表现"),
    nbf.v4.new_code_cell(
        """summary = pd.read_csv(TABLE_DIR / "group_summary.csv", index_col=0)
display(summary.round(4))"""
    ),
    nbf.v4.new_markdown_cell("### 4. 每日趋势与处理强度"),
    nbf.v4.new_code_cell(
        """display(Image(filename=str(CHART_DIR / "01_daily_trends.png")))"""
    ),
    nbf.v4.new_code_cell(
        """display(Image(filename=str(CHART_DIR / "06_coupon_intensity_trend.png")))"""
    ),
    nbf.v4.new_markdown_cell(
        """两组请求量、订单量和 GMV 随日期同步波动，配对结构明显。每单券额并非固定干预，而是随时间持续下降。"""
    ),
    nbf.v4.new_markdown_cell("### 5. 配对检验与效应量"),
    nbf.v4.new_code_cell(
        """tests = pd.read_csv(TABLE_DIR / "paired_test_results.csv")
cols = ["metric_cn", "control_mean", "experiment_mean", "relative_diff",
        "mean_diff", "ci_low", "ci_high", "paired_t_p", "wilcoxon_p",
        "paired_t_p_holm", "cohen_dz"]
display(tests[cols].round(6))"""
    ),
    nbf.v4.new_code_cell(
        """display(Image(filename=str(CHART_DIR / "03_metric_lifts.png")))"""
    ),
    nbf.v4.new_markdown_cell("### 6. 独立样本与配对检验的结论反转"),
    nbf.v4.new_code_cell(
        """method_compare = pd.read_csv(TABLE_DIR / "gmv_method_comparison.csv")
display(method_compare.round(6))
display(Image(filename=str(CHART_DIR / "04_method_reversal.png")))"""
    ),
    nbf.v4.new_markdown_cell(
        """忽略日期配对时，跨日期波动掩盖了组间差异；配对检验利用同一天的组间差值，显著降低噪声。"""
    ),
    nbf.v4.new_markdown_cell("### 7. GMV差值与稳健性"),
    nbf.v4.new_code_cell(
        """display(Image(filename=str(CHART_DIR / "02_gmv_paired_difference.png")))
robustness = pd.read_csv(TABLE_DIR / "gmv_robustness_checks.csv")
display(robustness.round(6))"""
    ),
    nbf.v4.new_markdown_cell(
        """Bootstrap区间仍完全低于0；但前半程与后半程效果差异明显，提示策略效果随时间或券额路径变化。"""
    ),
    nbf.v4.new_markdown_cell("### 8. 补贴经济性"),
    nbf.v4.new_code_cell(
        """economics = pd.read_csv(TABLE_DIR / "economics_summary.csv")
display(economics.round(4))"""
    ),
    nbf.v4.new_markdown_cell(
        """实验期内，实验组总GMV比对照组低约177,224，而优惠券成本高约287。增量ROI分母很小且收益为负，因此不应把该比值作为稳定的外推参数；决策重点是“额外补贴没有形成正向GMV增量”。"""
    ),
    nbf.v4.new_markdown_cell("### 9. 功效与MDE"),
    nbf.v4.new_code_cell(
        """power = pd.read_csv(TABLE_DIR / "power_analysis.csv")
display(power.round(6))
display(Image(filename=str(CHART_DIR / "05_power_and_mde.png")))"""
    ),
    nbf.v4.new_markdown_cell(
        """当前29对样本下，80%功效对应的GMV MDE约为日均0.86%。基于已观测效应量计算的事后功效约98.4%，只作为敏感性描述，不作为独立证据。"""
    ),
    nbf.v4.new_markdown_cell(
        """## Takeaways

1. **不建议直接推广当前优惠券路径。** 实验组GMV和客单价显著下降，优惠券成本反而上升。
2. **配对检验是本项目的关键方法选择。** 独立样本检验会被共同日期波动掩盖，错误得出“没有差异”。
3. **体验护栏有所改善，但不能单独支持推广。** 完成率小幅提高、取消率下降，需要结合收益指标综合判断。
4. **建议进行第二轮受控实验。** 固定或分层控制券额，预先定义GMV或增量价值为主要指标，并保留完成率与取消率护栏。
5. **结论需保留边界。** 日期级汇总数据无法验证用户级随机化、SRM、重复曝光和长期效果。
"""
    ),
]

nb["cells"] = cells
nbf.write(nb, OUT)
print(OUT)
