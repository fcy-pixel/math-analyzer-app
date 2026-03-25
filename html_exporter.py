"""
HTML Report Exporter — 完整還原 Streamlit 分析介面
Generates a self-contained HTML file with embedded Plotly charts + styled tables.
"""

import html
from datetime import datetime
from typing import Dict, List, Optional

import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio


# ---------------------------------------------------------------------------
# CSS — mirrors Streamlit light-theme styling
# ---------------------------------------------------------------------------
_CSS = """
<style>
  :root{--bg:#ffffff;--fg:#0e1117;--fg2:#555;--accent:#667eea;--green:#43a047;
    --blue:#1e88e5;--yellow:#f9a825;--red:#e53935;--border:#e0e0e0;--card-bg:#f8f9fa;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Noto Sans CJK TC",
    "PingFang TC","Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--fg);
    line-height:1.6;padding:24px 40px;max-width:1100px;margin:0 auto;}
  h1{font-size:1.8em;margin:16px 0 8px;border-bottom:2px solid var(--accent);padding-bottom:6px;}
  h2{font-size:1.4em;margin:20px 0 8px;color:var(--accent);}
  h3{font-size:1.15em;margin:16px 0 6px;}
  hr{border:none;border-top:1px solid var(--border);margin:18px 0;}
  .metrics{display:flex;gap:12px;margin:12px 0;}
  .metric{flex:1;background:var(--card-bg);border:1px solid var(--border);border-radius:8px;
    padding:14px 16px;text-align:center;}
  .metric .label{font-size:0.85em;color:var(--fg2);}
  .metric .value{font-size:1.5em;font-weight:700;color:var(--accent);}
  .info-box{background:#e8f4fd;border-left:4px solid var(--blue);padding:10px 14px;
    border-radius:4px;margin:10px 0;font-size:0.95em;}
  .warn-box{background:#fff8e1;border-left:4px solid var(--yellow);padding:10px 14px;
    border-radius:4px;margin:10px 0;font-size:0.95em;}
  .success-box{background:#e8f5e9;border-left:4px solid var(--green);padding:10px 14px;
    border-radius:4px;margin:10px 0;font-size:0.95em;}
  .error-box{background:#ffebee;border-left:4px solid var(--red);padding:10px 14px;
    border-radius:4px;margin:10px 0;font-size:0.95em;}
  table{width:100%;border-collapse:collapse;margin:10px 0;font-size:0.88em;}
  th{background:var(--accent);color:#fff;padding:8px 10px;text-align:left;font-weight:600;}
  td{padding:7px 10px;border-bottom:1px solid var(--border);}
  tr:nth-child(even) td{background:#f8f9fa;}
  .card{padding:10px 14px;border-radius:6px;margin:6px 0;border:1px solid var(--border);}
  .card-red{background:#ffebee;border-color:#ef9a9a;}
  .card-yellow{background:#fff8e1;border-color:#ffe082;}
  .card-green{background:#e8f5e9;border-color:#a5d6a7;}
  .chart{margin:14px 0;}
  .two-col{display:flex;gap:16px;margin:12px 0;}
  .two-col > div{flex:1;}
  .section{margin-bottom:24px;}
  details{margin:6px 0;border:1px solid var(--border);border-radius:6px;padding:0;}
  summary{padding:10px 14px;cursor:pointer;font-weight:600;background:var(--card-bg);
    border-radius:6px;}
  details[open] summary{border-bottom:1px solid var(--border);border-radius:6px 6px 0 0;}
  details .inner{padding:10px 14px;}
  .footer{text-align:center;color:var(--fg2);font-size:0.8em;margin-top:30px;
    border-top:1px solid var(--border);padding-top:10px;}
  .badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:0.8em;font-weight:600;}
  .badge-red{background:#ffcdd2;color:#c62828;}
  .badge-yellow{background:#fff9c4;color:#f57f17;}
  .badge-green{background:#c8e6c9;color:#2e7d32;}
  .badge-blue{background:#bbdefb;color:#1565c0;}
  @media print{body{padding:10px;max-width:100%;} .no-print{display:none;}}
</style>
"""


def _e(text) -> str:
    """HTML-escape helper."""
    if text is None:
        return ""
    return html.escape(str(text))


def _plotly_html(fig, height: int = 400) -> str:
    """Render a Plotly figure to embedded HTML div (no external JS needed)."""
    return pio.to_html(fig, full_html=False, include_plotlyjs="cdn",
                       config={"displayModeBar": False},
                       default_height=f"{height}px")


def _table_html(headers: List[str], rows: List[List[str]]) -> str:
    """Render a styled HTML table."""
    parts = ["<table><thead><tr>"]
    for h in headers:
        parts.append(f"<th>{_e(h)}</th>")
    parts.append("</tr></thead><tbody>")
    for row in rows:
        parts.append("<tr>")
        for cell in row:
            parts.append(f"<td>{_e(cell)}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table>")
    return "".join(parts)


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------
def build_student_html_report(
    agg: Dict,
    insights: Optional[Dict],
    grade: str,
    label: str = "",
) -> str:
    """Build a self-contained HTML report that mirrors the Streamlit analysis UI."""

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    label_str = f"（{_e(label)}）" if label else ""
    total_s = agg.get("total_students", 0)
    avg = agg.get("class_average", 0)
    dist = agg.get("class_distribution", {})
    weak_q = agg.get("weak_questions", [])
    strand_stats = agg.get("strand_stats", [])
    ranking = agg.get("student_ranking", [])
    q_stats = agg.get("question_stats", [])
    student_results = agg.get("student_results", [])
    valid_s = agg.get("valid_students", total_s)
    need_help = dist.get("需要改善(<55%)", 0)

    parts = [
        "<!DOCTYPE html><html lang='zh-Hant'><head>",
        "<meta charset='utf-8'>",
        f"<title>{_e(grade)}{label_str} 全班數學表現分析報告</title>",
        _CSS,
        "</head><body>",
    ]

    # ── Title ──
    parts.append(f"<h1>📊 {_e(grade)}{label_str} 全班數學表現分析報告</h1>")
    parts.append(f"<p style='color:var(--fg2);font-size:0.9em;'>報告生成日期：{now_str}</p>")

    # ══════════════════════════════════════════════════════════════════
    # Section 1: 整體概覽
    # ══════════════════════════════════════════════════════════════════
    parts.append('<div class="section">')
    parts.append("<h2>📋 整體概覽</h2>")

    # Metrics
    parts.append('<div class="metrics">')
    parts.append(f'<div class="metric"><div class="label">分析學生數</div><div class="value">{total_s} 人</div></div>')
    parts.append(f'<div class="metric"><div class="label">全班平均分</div><div class="value">{avg:.1f}%</div></div>')
    parts.append(f'<div class="metric"><div class="label">弱題數目（&lt;60%）</div><div class="value">{len(weak_q)}</div></div>')
    parts.append(f'<div class="metric"><div class="label">需要關注學生</div><div class="value">{need_help} 人</div></div>')
    parts.append("</div>")

    # Diagnosis summary
    if insights and not insights.get("parse_error"):
        diag = insights.get("overall_diagnosis", "")
        if diag:
            parts.append(f'<div class="info-box">🔬 <strong>診斷摘要：</strong> {_e(diag)}</div>')

    # Pie chart + histogram side by side
    parts.append('<div class="two-col">')

    # Pie chart
    parts.append("<div>")
    if dist:
        df_data = [{"表現等級": k, "人數": v} for k, v in dist.items() if v > 0]
        if df_data:
            import pandas as pd
            df_pie = pd.DataFrame(df_data)
            color_map = {
                "優秀(≥85%)": "#43a047", "良好(70-84%)": "#1e88e5",
                "一般(55-69%)": "#f9a825", "需要改善(<55%)": "#e53935",
            }
            fig = px.pie(df_pie, names="表現等級", values="人數",
                         color="表現等級", color_discrete_map=color_map,
                         title="全班成績等級分佈")
            fig.update_traces(textinfo="label+value+percent")
            parts.append(f'<div class="chart">{_plotly_html(fig, 360)}</div>')
    parts.append("</div>")

    # Histogram
    parts.append("<div>")
    pcts = [r.get("percentage", 0) for r in student_results if not r.get("parse_error")]
    if pcts:
        fig = px.histogram(x=pcts, nbins=10,
                           labels={"x": "得分率 (%)", "y": "學生人數"},
                           title="全班得分率分佈直方圖",
                           color_discrete_sequence=["#667eea"])
        fig.add_vline(x=avg, line_dash="dash", line_color="red",
                      annotation_text=f"平均 {avg:.1f}%")
        fig.update_xaxes(range=[0, 100])
        parts.append(f'<div class="chart">{_plotly_html(fig, 360)}</div>')
    parts.append("</div>")
    parts.append("</div>")  # end two-col

    # Strand radar / bar
    if strand_stats:
        parts.append("<h3>📈 各課程範疇全班正確率</h3>")
        cats = [s["strand"] for s in strand_stats]
        vals = [s["class_average_rate"] for s in strand_stats]
        if len(cats) >= 3:
            fig = go.Figure(go.Scatterpolar(
                r=vals + [vals[0]], theta=cats + [cats[0]], fill="toself",
                fillcolor="rgba(102,126,234,0.25)",
                line=dict(color="rgb(102,126,234)", width=2),
            ))
            fig.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                showlegend=False, height=380,
                margin=dict(l=60, r=60, t=40, b=40),
            )
        else:
            import pandas as pd
            df_s = pd.DataFrame(strand_stats)
            fig = px.bar(df_s, x="strand", y="class_average_rate",
                         color="class_average_rate", color_continuous_scale="RdYlGn",
                         range_y=[0, 100], title="各課程範疇全班正確率")
            fig.add_hline(y=60, line_dash="dash", line_color="red", annotation_text="60% 基準線")
        parts.append(f'<div class="chart">{_plotly_html(fig, 380)}</div>')

    parts.append("</div>")  # end section

    # ══════════════════════════════════════════════════════════════════
    # Section 2: 學生成績排名
    # ══════════════════════════════════════════════════════════════════
    parts.append('<hr><div class="section">')
    parts.append(f"<h2>🏅 學生成績排名（共 {len(ranking)} 位學生）</h2>")

    if ranking:
        sicon = {"優秀(≥85%)": "🌟", "良好(70-84%)": "👍", "一般(55-69%)": "△", "需要改善(<55%)": "⚠️", "分析失敗": "❌"}
        headers = ["排名", "學生", "得分率", "得分", "表現等級"]
        rows = []
        for s in ranking:
            level = s.get("performance_level", "")
            icon = sicon.get(level, "")
            rows.append([
                str(s.get("rank", "")),
                s.get("student_name", ""),
                f"{s.get('percentage', 0):.1f}%",
                f"{s.get('total_marks_awarded', '—')} / {s.get('total_marks_possible', '—')}",
                f"{icon} {level}",
            ])
        parts.append(_table_html(headers, rows))

        # Bar chart
        import pandas as pd
        df_rank = pd.DataFrame([
            {"學生": s["student_name"], "得分率": s["percentage"]}
            for s in sorted(ranking, key=lambda x: x.get("percentage", 0))
        ])
        fig = px.bar(df_rank, x="得分率", y="學生", orientation="h",
                     color="得分率", color_continuous_scale="RdYlGn", range_x=[0, 100],
                     title="全班學生得分率排行",
                     labels={"得分率": "得分率 (%)", "學生": ""})
        fig.add_vline(x=avg, line_dash="dash", line_color="blue",
                      annotation_text=f"平均 {avg:.1f}%")
        fig.add_vline(x=60, line_dash="dot", line_color="red", annotation_text="60% 基準")
        fig.update_layout(height=max(300, 22 * len(df_rank)), margin=dict(l=120))
        parts.append(f'<div class="chart">{_plotly_html(fig, max(300, 22 * len(df_rank)))}</div>')

    parts.append("</div>")

    # ══════════════════════════════════════════════════════════════════
    # Section 3: 自動批改
    # ══════════════════════════════════════════════════════════════════
    parts.append('<hr><div class="section">')
    parts.append("<h2>✏️ 自動批改 — 各學生答錯題目</h2>")
    parts.append('<p style="color:var(--fg2);font-size:0.9em;">只列出每位學生答錯的題目，方便老師用紅筆在紙本工作紙上批改。</p>')

    for student in student_results:
        name = student.get("student_name", "未知")
        if student.get("parse_error"):
            parts.append(f'<div class="warn-box"><strong>{_e(name)}</strong> — 分析失敗，無法取得答題結果</div>')
            continue

        q_results = student.get("question_results", [])
        wrong = [q for q in q_results if q.get("is_correct") is False]
        total_q = len(q_results)
        wrong_count = len(wrong)
        pct = student.get("percentage", 0)

        if not wrong:
            parts.append(f'<div class="success-box"><strong>{_e(name)}</strong> — ✅ 全部答對（{total_q}/{total_q}）</div>')
            continue

        parts.append(f'<details{"" if wrong_count < 3 else " open"}>')
        parts.append(f'<summary>❌ {_e(name)}　—　答錯 {wrong_count} 題 / 共 {total_q} 題　（得分率 {pct:.0f}%）</summary>')
        parts.append('<div class="inner">')

        headers = ["題目", "考核主題", "學生答案", "正確答案", "得分", "錯誤類型", "錯誤說明"]
        rows = []
        for q in wrong:
            rows.append([
                q.get("question_ref", ""),
                q.get("topic", ""),
                q.get("student_answer", "—"),
                q.get("correct_answer", "—"),
                f"{q.get('marks_awarded', 0)} / {q.get('marks_possible', '')}",
                q.get("error_type", "") or "",
                q.get("error_description", "") or "",
            ])
        parts.append(_table_html(headers, rows))
        parts.append("</div></details>")

    # Summary cross-table
    if q_stats and student_results:
        parts.append("<h3>📋 全班答錯題目一覽表</h3>")
        parts.append('<p style="color:var(--fg2);font-size:0.85em;">❌ = 答錯，空白 = 答對或未作答</p>')
        all_refs = [q["question_ref"] for q in q_stats]
        headers = ["學生"] + all_refs
        rows = []
        for student in student_results:
            if student.get("parse_error"):
                continue
            name = student.get("student_name", "未知")
            q_map = {str(q.get("question_ref", "")): q.get("is_correct")
                     for q in student.get("question_results", [])}
            row = [name]
            for ref in all_refs:
                val = q_map.get(str(ref))
                row.append("❌" if val is False else "")
            rows.append(row)
        if rows:
            parts.append(_table_html(headers, rows))

    parts.append("</div>")

    # ══════════════════════════════════════════════════════════════════
    # Section 4: 逐題分析
    # ══════════════════════════════════════════════════════════════════
    parts.append('<hr><div class="section">')
    parts.append(f"<h2>📝 逐題全班正確率（共 {len(q_stats)} 題）</h2>")

    if q_stats:
        headers = ["題目", "考核主題", "範疇", "全班正確率", "正確人數", "常見錯誤"]
        rows = []
        for q in q_stats:
            rate = q.get("class_correct_rate", 0)
            rows.append([
                q["question_ref"],
                q.get("topic", ""),
                q.get("strand", ""),
                f"{rate}%",
                f"{q['class_correct_count']} / {valid_s}",
                "；".join(q.get("common_errors", [])[:2]) or "—",
            ])
        parts.append(_table_html(headers, rows))

        # Bar chart
        import pandas as pd
        df_q = pd.DataFrame([
            {"題目": q["question_ref"], "正確率": q["class_correct_rate"]}
            for q in q_stats
        ])
        fig = px.bar(df_q, x="題目", y="正確率",
                     color="正確率", color_continuous_scale="RdYlGn", range_y=[0, 100],
                     title="各題全班正確率", labels={"正確率": "正確率 (%)"})
        fig.add_hline(y=60, line_dash="dash", line_color="red", annotation_text="60% 基準線")
        parts.append(f'<div class="chart">{_plotly_html(fig, 380)}</div>')

    parts.append("</div>")

    # ══════════════════════════════════════════════════════════════════
    # Section 5: 弱點熱圖
    # ══════════════════════════════════════════════════════════════════
    if q_stats and student_results:
        parts.append('<hr><div class="section">')
        parts.append("<h2>🔥 學生 × 題目 答對熱圖</h2>")
        parts.append('<p style="color:var(--fg2);font-size:0.85em;">🟢 答對　🔴 答錯　⬜ 未作答</p>')

        all_refs = [q["question_ref"] for q in q_stats]
        names = [s.get("student_name", f"學生{i+1}") for i, s in enumerate(student_results)]
        z_matrix, text_matrix = [], []
        for student in student_results:
            q_map = {str(q.get("question_ref", "")): q.get("is_correct")
                     for q in student.get("question_results", [])}
            row_z, row_t = [], []
            for ref in all_refs:
                val = q_map.get(str(ref))
                if val is True:
                    row_z.append(1); row_t.append("✓")
                elif val is False:
                    row_z.append(0); row_t.append("✗")
                else:
                    row_z.append(0.5); row_t.append("—")
            z_matrix.append(row_z); text_matrix.append(row_t)

        fig = go.Figure(data=go.Heatmap(
            z=z_matrix, x=all_refs, y=names,
            colorscale=[[0, "#e53935"], [0.45, "#ffb300"], [0.55, "#ffb300"], [1, "#43a047"]],
            showscale=False, text=text_matrix, texttemplate="%{text}", xgap=2, ygap=2,
        ))
        fig.update_layout(xaxis_title="題目", yaxis_title="學生",
                          height=max(350, 26 * len(names)),
                          margin=dict(l=100, r=20, t=40, b=60), font_size=11)
        parts.append(f'<div class="chart">{_plotly_html(fig, max(350, 26 * len(names)))}</div>')
        parts.append("</div>")

    # ══════════════════════════════════════════════════════════════════
    # Section 6: 弱點診斷
    # ══════════════════════════════════════════════════════════════════
    parts.append('<hr><div class="section">')
    parts.append(f"<h2>🎯 弱點診斷</h2>")

    if weak_q:
        parts.append(f"<h3>🔴 弱題排行榜（正確率 &lt; 60%，共 {len(weak_q)} 題）</h3>")
        headers = ["排名", "", "題目", "全班正確率", "正確人數", "考核主題", "範疇", "常見錯誤"]
        rows = []
        for q in weak_q:
            rate = q.get("class_correct_rate", 0)
            icon = "🔴" if rate < 40 else "🟡"
            rows.append([
                str(q["rank"]), icon, q["question_ref"], f"{rate}%",
                f"{q['class_correct_count']} / {valid_s}",
                q.get("topic", ""), q.get("strand", ""),
                "；".join(q.get("common_errors", [])[:2]) or "—",
            ])
        parts.append(_table_html(headers, rows))

        # Bar chart
        import pandas as pd
        df_wq = pd.DataFrame([
            {"題目": q["question_ref"], "正確率": q["class_correct_rate"]}
            for q in weak_q
        ])
        fig = px.bar(df_wq, x="題目", y="正確率", color="正確率",
                     color_continuous_scale="RdYlGn", range_y=[0, 100],
                     title="弱題正確率", labels={"正確率": "正確率 (%)"})
        fig.add_hline(y=60, line_dash="dash", line_color="red")
        parts.append(f'<div class="chart">{_plotly_html(fig, 350)}</div>')

    # Strand cards
    if strand_stats:
        parts.append("<h3>📊 各課程範疇表現</h3>")
        for s in strand_stats:
            rate = s.get("class_average_rate", 0)
            status = s.get("status", "")
            icon = "🔴" if status == "弱項" else "🟡" if status == "一般" else "✅"
            css_cls = "card-red" if status == "弱項" else "card-yellow" if status == "一般" else "card-green"
            filled = int(rate / 10)
            bar_str = "█" * filled + "░" * (10 - filled)
            qs = ", ".join(s.get("questions", [])[:6])
            if len(s.get("questions", [])) > 6:
                qs += "…"
            parts.append(
                f'<div class="card {css_cls}">'
                f'{icon} <strong>{_e(s["strand"])}</strong>　{rate}%　'
                f'<code style="font-size:0.8em">{bar_str}</code>　'
                f'<em>（涉及題目：{_e(qs)}）</em></div>'
            )

    # AI weak strand analysis
    if insights and not insights.get("parse_error"):
        ws_list = insights.get("weak_strands_analysis", [])
        if ws_list:
            parts.append("<h3>🔍 弱項範疇深入分析</h3>")
            for ws in ws_list:
                parts.append(f'<details><summary>{_e(ws.get("strand",""))}　（全班正確率：{ws.get("class_average_rate","")}%）</summary>')
                parts.append('<div class="inner">')
                for issue in ws.get("key_issues", []):
                    parts.append(f"<p>- {_e(issue)}</p>")
                if ws.get("misconception"):
                    parts.append(f'<div class="warn-box">🧩 可能的概念誤解：{_e(ws["misconception"])}</div>')
                if ws.get("curriculum_link"):
                    parts.append(f'<div class="info-box">📚 課程連結：{_e(ws["curriculum_link"])}</div>')
                parts.append("</div></details>")

        # Error type analysis
        et = insights.get("error_type_analysis", {})
        if et:
            parts.append("<h3>🔎 錯誤類型分析</h3>")
            parts.append('<div class="two-col">')
            parts.append(f'<div><strong>🧩 概念性誤解</strong><p>{_e(et.get("conceptual", "—"))}</p></div>')
            parts.append(f'<div><strong>🔢 程序性錯誤</strong><p>{_e(et.get("procedural", "—"))}</p></div>')
            parts.append("</div>")

    parts.append("</div>")

    # ══════════════════════════════════════════════════════════════════
    # Section 7: 教學建議
    # ══════════════════════════════════════════════════════════════════
    if insights and not insights.get("parse_error"):
        recs = insights.get("teaching_recommendations", [])
        if recs:
            parts.append('<hr><div class="section">')
            parts.append("<h2>💡 教學建議</h2>")
            parts.append("<h3>📅 補救教學建議</h3>")
            pri_order = {"高": 0, "中": 1, "低": 2}
            recs_sorted = sorted(recs, key=lambda r: pri_order.get(r.get("priority", ""), 9))
            for rec in recs_sorted:
                pri = rec.get("priority", "")
                icon = "🔴" if pri == "高" else "🟡" if pri == "中" else "🟢"
                parts.append(f'<details><summary>{icon} {_e(rec.get("strand",""))} — {_e(rec.get("strategy",""))}</summary>')
                parts.append('<div class="inner">')
                parts.append(f"<p><strong>優先級：</strong>{_e(pri)}　<strong>建議時間：</strong>{_e(rec.get('timeline',''))}</p>")
                acts = rec.get("activities", [])
                if acts:
                    parts.append("<p><strong>教學活動：</strong></p><ul>")
                    for a in acts:
                        parts.append(f"<li>{_e(a)}</li>")
                    parts.append("</ul>")
                parts.append("</div></details>")

            # Attention students
            attn = insights.get("attention_students_note", "")
            if attn:
                parts.append("<h3>👀 需要個別關注的學生</h3>")
                parts.append(f'<div class="warn-box">{_e(attn)}</div>')

            # Positive findings
            pos = insights.get("positive_findings", "")
            if pos:
                parts.append("<h3>💪 全班亮點</h3>")
                parts.append(f'<div class="success-box">{_e(pos)}</div>')

            parts.append("</div>")

    # ── Footer ──
    parts.append(f'<div class="footer">報告生成日期：{now_str} | 由 Qwen AI 驅動 | 小學數學學生表現分析系統<br>'
                 '課程參考：香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版</div>')

    parts.append("</body></html>")
    return "".join(parts)
