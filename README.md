# Raw Data Dashboard Generator

这是一个可直接运行的轻量网站（基于 Streamlit），支持：

1. 上传多个 Excel / CSV raw data 文件
2. 自动读取 Excel 中的全部 Sheet
3. 将不同媒体的 raw data 统一映射到以下目标字段：
   - Media, Position, Market, Date, Landing page, Cost, IMP, CLICK, CPM, CPC, CTR, ENG, Like, Forward, Comment, Revenue, Orders
4. 在线生成 Dashboard
5. 下载 Excel 结果：
   - Sheet1_Dashboard：汇总指标、By Media、By Market、By Date
   - 后续 Sheet：按 Media 拆分后的明细数据（保留 source_file / source_sheet）

## 运行方式

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 使用方法

1. 打开浏览器进入本地 Streamlit 地址（通常是 http://localhost:8501 ）
2. 上传多个 `.xlsx` / `.xls` / `.csv` 文件
3. 网站会自动读取全部 sheet、标准化字段并计算 KPI
4. 可在页面查看 Dashboard，并点击按钮下载输出 Excel

## 说明

- 如果原始文件缺少 `CTR / CPM / CPC`，程序会自动根据 `Cost / IMP / CLICK` 补算。
- 如果不同媒体字段命名不同，可以继续在 `app.py` 中扩充 `COLUMN_MAP`。
