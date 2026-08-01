# 网约车优惠券 A/B 测试 Dashboard

这是主项目的轻量级交互展示层。Dashboard 使用 `data/test.xlsx`，按日期实时重算配对指标与统计检验，适合部署到 Streamlit Community Cloud 作为求职作品集。

## 功能

- GMV、完成率、取消率、客单价和优惠券成本 KPI 卡片
- 日期范围全局联动筛选
- 请求数、订单数、GMV、每单优惠券金额趋势图
- GMV 等五项指标的日期级配对差值图
- Welch 独立样本检验与日期配对检验对比
- 可排序的完整统计检验结果表

## 本地运行

在仓库根目录执行：

```bash
pip install -r dashboard/requirements.txt
streamlit run dashboard/app.py
```

应用会自动读取仓库中的 `data/test.xlsx`。

## 部署到 Streamlit Community Cloud

1. 登录 [Streamlit Community Cloud](https://share.streamlit.io/)，并连接保存本项目的 GitHub 账号。
2. 点击 **Create app**，选择 `coupon-ab-test` 仓库和 `main` 分支。
3. **Main file path** 填写 `dashboard/app.py`。
4. 点击 **Deploy**。平台会读取 `dashboard/requirements.txt` 并安装依赖。
5. 部署完成后，将公开应用链接补充到主项目 README 和飞书项目报告中。

本项目使用仓库内的静态 Excel 文件，不需要配置 Secrets、数据库或外部 API。
