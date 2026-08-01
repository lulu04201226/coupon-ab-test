# 网约车优惠券 A/B 测试分析

本项目基于 `test.xlsx` 中 29 天、对照组与实验组按日期对齐的汇总数据，评估提高优惠券力度是否值得推广。

核心分析包括：

- 数据质量与实验结构判断
- 核心业务指标与补贴经济性
- 日期级配对 t 检验
- Wilcoxon 与 Bootstrap 稳健性验证
- 独立样本检验与配对检验的结论差异
- 功效、样本量与最小可检测效应（MDE）
- 咨询式项目报告

## 核心结论

- 实验组日均 GMV 比对照组低约 6,111（-1.26%）。
- Welch 独立样本检验 p=0.9387；按日期配对后 p=0.0002，Wilcoxon p=0.0018。
- 实验组优惠券成本增加约 0.77%，绝对 ROI 下降约 2.41%。
- 完成率提高约 0.20 个百分点，取消率降低约 0.21 个百分点。
- 当前 29 对样本下，80% 功效对应的 GMV MDE 约为日均 0.86%。
- GMV 负向效果主要出现在后半程，且券额随时间变化，因此建议停止直接放量并进行固定处理强度的二次实验。

## 项目结构

```text
coupon-ab-test/
├─ data/
│  └─ test.xlsx
├─ notebooks/
│  └─ 01_didi_ab_test_analysis.ipynb
├─ src/
│  └─ run_analysis.py
├─ outputs/
│  ├─ charts/
│  └─ tables/
├─ report/
│  └─ 网约车优惠券AB测试分析报告.docx
├─ README.md
└─ requirements.txt
```

## 运行方式

项目已包含分析使用的 `data/test.xlsx`。安装依赖后执行：

```bash
pip install -r requirements.txt
python src/run_analysis.py
jupyter nbconvert --execute --to notebook --inplace notebooks/01_didi_ab_test_analysis.ipynb
```

## 项目报告

完整项目报告发布在飞书知识库：

- [网约车优惠券A/B测试分析｜项目报告](https://my.feishu.cn/wiki/Wl2IwREdhijsppkqENzcAsYXnXb)

## 分析边界

原始数据为日期级汇总数据，并非用户级实验日志。项目可以验证日期配对后的指标差异，但无法直接核验用户级随机化、重复曝光和样本比例异常。
