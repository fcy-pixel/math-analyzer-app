# 香港小學數學分析系統 📊

Streamlit app that uses Qwen vision AI to:

- **批量批改全班試卷** — 上傳全班掃描 PDF，AI 自動逐份批改並生成全班報告
- **AQP + 答案版分析** — 分析年度評估計劃題目及答案，診斷全年級弱點

## 功能

- 📝 逐題正確率統計、全班熱圖、學生排名
- 🎯 弱題及弱項範疇診斷
- 💡 AI 生成教學建議
- 📥 PDF 報告匯出（含圖表）

## 安裝與運行

```bash
pip install -r requirements.txt
streamlit run app.py
```

運行後在側欄輸入 **Qwen International API Key**（格式：`sk-...`）。

> API 申請：https://www.alibabacloud.com/help/en/dashscope/

## 技術棧

- [Streamlit](https://streamlit.io/) — UI
- [Qwen VL Max](https://qwenlm.github.io/) — 視覺模型（批改手寫試卷）
- [PyMuPDF](https://pymupdf.readthedocs.io/) — PDF 處理
- [Plotly](https://plotly.com/) — 互動圖表
- [ReportLab](https://www.reportlab.com/) — PDF 匯出
