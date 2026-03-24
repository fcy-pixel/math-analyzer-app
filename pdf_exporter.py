"""
PDF Report Generator — 小學數學全年級表現分析系統
Uses ReportLab with STHeiti Chinese font + Plotly/Kaleido for chart images.
"""

import io
import os
from datetime import datetime
from typing import Dict, Optional, List

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ---------------------------------------------------------------------------
# Font setup — STHeiti supports Traditional + Simplified Chinese on macOS
# ---------------------------------------------------------------------------
_FONT = "Helvetica"
_FONT_REGISTERED = False

_CJK_CANDIDATES = [
    # macOS
    ("/System/Library/Fonts/STHeiti Light.ttc", 0),
    ("/System/Library/Fonts/STHeiti Medium.ttc", 0),
    ("/System/Library/Fonts/Hiragino Sans GB.ttc", 0),
    # Linux / Streamlit Cloud (apt: fonts-wqy-microhei)
    ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 0),
    # Linux / Streamlit Cloud (apt: fonts-noto-cjk)
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
    ("/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc", 0),
    # Windows fallback
    ("C:/Windows/Fonts/msyh.ttc", 0),
    ("C:/Windows/Fonts/simsun.ttc", 0),
]


def _ensure_font() -> str:
    global _FONT, _FONT_REGISTERED
    if _FONT_REGISTERED:
        return _FONT
    for path, idx in _CJK_CANDIDATES:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("CJK", path, subfontIndex=idx))
                _FONT = "CJK"
                _FONT_REGISTERED = True
                return _FONT
            except Exception:
                continue
    _FONT_REGISTERED = True
    return _FONT


# ---------------------------------------------------------------------------
# Colour scheme
# ---------------------------------------------------------------------------
C_PRIMARY = colors.HexColor("#667eea")
C_DANGER = colors.HexColor("#e53935")
C_WARNING = colors.HexColor("#f9a825")
C_SUCCESS = colors.HexColor("#43a047")
C_LIGHT = colors.HexColor("#f5f5f5")
C_DARK = colors.HexColor("#212529")
C_GREY = colors.HexColor("#6c757d")


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def _styles(font: str) -> Dict[str, ParagraphStyle]:
    return {
        "title": ParagraphStyle("title", fontName=font, fontSize=20,
                                textColor=colors.white, leading=26,
                                spaceAfter=4, alignment=TA_CENTER),
        "subtitle": ParagraphStyle("subtitle", fontName=font, fontSize=11,
                                   textColor=colors.white, leading=15,
                                   alignment=TA_CENTER),
        "h1": ParagraphStyle("h1", fontName=font, fontSize=14,
                              textColor=C_PRIMARY, leading=20,
                              spaceBefore=12, spaceAfter=5),
        "h2": ParagraphStyle("h2", fontName=font, fontSize=11,
                              textColor=C_DARK, leading=16,
                              spaceBefore=8, spaceAfter=3),
        "body": ParagraphStyle("body", fontName=font, fontSize=9.5,
                               textColor=C_DARK, leading=15, spaceAfter=3),
        "warn": ParagraphStyle("warn", fontName=font, fontSize=9.5,
                               textColor=C_DANGER, leading=15, spaceAfter=3),
        "success": ParagraphStyle("success", fontName=font, fontSize=9.5,
                                  textColor=C_SUCCESS, leading=15, spaceAfter=3),
        "caption": ParagraphStyle("caption", fontName=font, fontSize=8,
                                  textColor=C_GREY, leading=12,
                                  alignment=TA_CENTER),
        "small": ParagraphStyle("small", fontName=font, fontSize=8.5,
                                textColor=C_GREY, leading=12, spaceAfter=2),
        "th": ParagraphStyle("th", fontName=font, fontSize=8.5,
                             textColor=colors.white, leading=12),
        "td": ParagraphStyle("td", fontName=font, fontSize=8,
                             textColor=C_DARK, leading=11),
    }


def _p(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(str(text) if text is not None else "", style)


def _build_table(data: List[List], col_widths: List, s: Dict,
                 header_bg=C_PRIMARY) -> Table:
    """Build a styled Table; all string cells become Paragraphs for word-wrap."""
    processed = []
    for ri, row in enumerate(data):
        new_row = []
        for cell in row:
            if isinstance(cell, str):
                style = s["th"] if ri == 0 else s["td"]
                new_row.append(Paragraph(cell, style))
            elif cell is None:
                new_row.append(Paragraph("", s["td"]))
            else:
                new_row.append(Paragraph(str(cell), s["td"] if ri > 0 else s["th"]))
        processed.append(new_row)

    tbl_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_LIGHT, colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#dee2e6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ])
    tbl = Table(processed, colWidths=col_widths, repeatRows=1,
                hAlign="LEFT", spaceBefore=4, spaceAfter=6)
    tbl.setStyle(tbl_style)
    return tbl


def _chart(fig, w_cm: float = 14.0, h_cm: float = 7.0) -> Optional[Image]:
    """Render a Plotly figure to a ReportLab Image via kaleido."""
    try:
        png = pio.to_image(fig, format="png", width=int(w_cm * 60),
                           height=int(h_cm * 60), scale=2)
        return Image(io.BytesIO(png), width=w_cm * cm, height=h_cm * cm)
    except Exception:
        return None


def _hr(w) -> HRFlowable:
    return HRFlowable(width=w, color=C_PRIMARY, thickness=1.5,
                      spaceAfter=6, spaceBefore=2)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_pdf(results: Dict, grade: str, class_label: str = "") -> bytes:
    """Render the full analysis report as PDF bytes."""
    font = _ensure_font()
    s = _styles(font)

    buf = io.BytesIO()
    W_PT = A4[0] - 3.6 * cm          # usable width in points
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=2.0 * cm, bottomMargin=2.0 * cm,
        title=f"{grade} 年級數學全年級分析報告",
        author="Qwen AI 小學數學分析系統",
    )

    label_str = f"（{class_label}）" if class_label else ""
    now_str = datetime.now().strftime("%Y 年 %m 月 %d 日")
    story = []

    # ── Title block ───────────────────────────────────────────────────────────
    title_tbl = Table(
        [[_p("小學數學全年級表現分析報告", s["title"])],
         [_p(f"{grade} 年級{label_str}  ·  {now_str}", s["subtitle"])]],
        colWidths=[W_PT],
    )
    title_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_PRIMARY),
        ("TOPPADDING", (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING", (0, 0), (-1, -1), 18),
        ("RIGHTPADDING", (0, 0), (-1, -1), 18),
    ]))
    story.append(title_tbl)
    story.append(Spacer(1, 0.4 * cm))

    meta = _build_table(
        [["項目", "內容"],
         ["年級", f"{grade} 年級{label_str}"],
         ["報告日期", now_str],
         ["課程依據", "香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版"],
         ["AQP 來源", "香港教育局全港基本能力評估計劃（AQP）"],
         ["分析系統", "小學數學全年級表現分析系統 · 由 Qwen AI 驅動"]],
        [4.5 * cm, W_PT - 4.5 * cm], s,
        header_bg=colors.HexColor("#5a6fd6"),
    )
    story.append(meta)
    story.append(PageBreak())

    # ── Section 1: AQP ───────────────────────────────────────────────────────
    if "aqp" in results and not results["aqp"].get("parse_error"):
        aqp = results["aqp"]
        story.append(_p("一、AQP 全年級成績分析", s["h1"]))
        story.append(_hr(W_PT))
        story.append(_p(
            "數據來源：香港教育局全港基本能力評估計劃（AQP）" + "  ·  " +
            "以下分析反映整個年級所有班別學生的共同表現",
            s["small"],
        ))
        story.append(Spacer(1, 0.25 * cm))

        # Overall performance
        overall = aqp.get("overall_performance", {})
        if overall:
            level = overall.get("performance_level", "")
            avg = overall.get("class_average_percentage")
            perf_line = f"整體表現水平：{level}"
            if avg is not None:
                perf_line += f"   |   全年級平均得分率：{avg}%"
            story.append(_p(perf_line, s["h2"]))
            story.append(_p(overall.get("summary", ""), s["body"]))
            story.append(Spacer(1, 0.2 * cm))

        # Strand table
        strands = aqp.get("strand_analysis", [])
        if strands:
            story.append(_p("◆ 各課程範疇表現", s["h2"]))
            data = [["課程範疇", "全年級得分率", "表現", "狀態", "普遍困難主題"]]
            for strand in strands:
                struggled = "、".join(strand.get("specific_topics_struggled", []))
                data.append([
                    strand.get("strand", ""),
                    str(strand.get("class_score", "—")),
                    strand.get("performance", ""),
                    strand.get("status", ""),
                    struggled,
                ])
            story.append(_build_table(
                data, [3.2 * cm, 2.2 * cm, 4.5 * cm, 1.8 * cm, W_PT - 11.7 * cm], s
            ))

            # Radar chart
            cats = [st.get("strand", "") for st in strands]
            vals = []
            for st in strands:
                raw = st.get("class_score") or st.get("score")
                status = st.get("status", "")
                if raw is not None:
                    try:
                        vals.append(min(100.0, float(str(raw).replace("%", ""))))
                        continue
                    except (ValueError, TypeError):
                        pass
                vals.append({"強項": 85, "一般": 58, "弱項": 30}.get(status, 50))
            if cats and vals:
                fig_radar = go.Figure(go.Scatterpolar(
                    r=vals + [vals[0]], theta=cats + [cats[0]],
                    fill="toself",
                    fillcolor="rgba(102,126,234,0.25)",
                    line=dict(color="rgb(102,126,234)", width=2),
                ))
                fig_radar.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                    showlegend=False, height=380, width=480,
                    margin=dict(l=60, r=60, t=30, b=30),
                )
                img = _chart(fig_radar, w_cm=11, h_cm=7.5)
                if img:
                    story.append(img)
                    story.append(_p("各課程範疇全年級表現雷達圖", s["caption"]))
                    story.append(Spacer(1, 0.2 * cm))

        # Weak question ranking
        weak_qs = aqp.get("weak_questions", [])
        if weak_qs:
            story.append(_p("◆ 弱題排行榜（全年級正確率低於 60%，按正確率由低至高排列）", s["h2"]))
            try:
                sorted_qs = sorted(
                    weak_qs,
                    key=lambda x: float(str(x.get("correct_rate", 100)).replace("%", ""))
                )
            except (ValueError, TypeError):
                sorted_qs = weak_qs
            data = [["排名", "題目", "全年級正確率", "考核主題", "課程範疇", "常見錯誤"]]
            for i, q in enumerate(sorted_qs, 1):
                rate = q.get("correct_rate", "—")
                rate_str = f"{rate}%" if isinstance(rate, (int, float)) else str(rate)
                data.append([
                    str(i),
                    q.get("question_ref", ""),
                    rate_str,
                    q.get("topic", ""),
                    q.get("strand", ""),
                    q.get("common_error", ""),
                ])
            story.append(_build_table(
                data, [1 * cm, 1.5 * cm, 2.2 * cm, 3.5 * cm, 2.5 * cm, W_PT - 10.7 * cm], s,
                header_bg=C_DANGER,
            ))

            # Bar chart of weak questions
            numeric_wq = []
            for q in sorted_qs:
                try:
                    numeric_wq.append({
                        "題目": q.get("question_ref", ""),
                        "正確率": float(str(q.get("correct_rate", "")).replace("%", "")),
                    })
                except (ValueError, TypeError):
                    pass
            if len(numeric_wq) >= 2:
                fig_wq = px.bar(
                    pd.DataFrame(numeric_wq), x="題目", y="正確率",
                    color="正確率", color_continuous_scale="RdYlGn", range_y=[0, 100],
                    title="弱題全年級正確率（由低至高）",
                )
                fig_wq.update_layout(height=380, coloraxis_showscale=False,
                                     margin=dict(l=50, r=20, t=40, b=60))
                img = _chart(fig_wq, w_cm=W_PT / cm, h_cm=7)
                if img:
                    story.append(img)
                    story.append(Spacer(1, 0.2 * cm))

        # Full question performance table
        qperf = aqp.get("question_performance", [])
        if qperf:
            story.append(_p("◆ 各題全年級正確率（完整列表）", s["h2"]))
            data = [["題目", "考核主題", "課程範疇", "全年級正確率", "難度", "常見錯誤"]]
            for q in qperf:
                rate = q.get("class_correct_rate")
                data.append([
                    q.get("question_ref", ""),
                    q.get("topic", ""),
                    q.get("strand", ""),
                    f"{rate}%" if rate is not None else "—",
                    q.get("difficulty", ""),
                    q.get("common_errors", ""),
                ])
            story.append(_build_table(
                data, [1.5 * cm, 3 * cm, 2.5 * cm, 2.5 * cm, 1.5 * cm, W_PT - 11 * cm], s
            ))

        # Weak areas
        class_weak = aqp.get("class_weak_areas", aqp.get("weak_areas", []))
        if class_weak:
            story.append(_p("◆ 全年級主要弱點及概念誤解", s["h2"]))
            for area in class_weak:
                sev = area.get("severity", "")
                icon = {"嚴重": "🔴", "中等": "🟡", "輕微": "🟢"}.get(sev, "◆")
                story.append(_p(
                    f"{icon} {area.get('topic', '')}（{area.get('strand', '')}）— {sev}",
                    s["h2"],
                ))
                story.append(_p(area.get("description", ""), s["body"]))
                if area.get("likely_misconception"):
                    story.append(_p(f"❗ 概念誤解：{area['likely_misconception']}", s["warn"]))
                if area.get("data_evidence"):
                    story.append(_p(f"📊 數據佐證：{area['data_evidence']}", s["small"]))
                story.append(Spacer(1, 0.1 * cm))

        # Teaching implications
        implications = aqp.get("teaching_implications", [])
        if implications:
            story.append(_p("◆ 教學啟示", s["h2"]))
            data = [["教學問題", "AQP 數據佐證", "建議教學策略"]]
            for impl in implications:
                data.append([
                    impl.get("issue", ""),
                    impl.get("evidence", ""),
                    impl.get("suggested_teaching_strategy", ""),
                ])
            story.append(_build_table(
                data, [3.5 * cm, 5 * cm, W_PT - 8.5 * cm], s
            ))

        story.append(PageBreak())

    # ── Section 2: Exam (answer key) ─────────────────────────────────────────
    if "exam" in results and not results["exam"].get("parse_error"):
        exam = results["exam"]
        story.append(_p("二、試卷結構分析（答案版）", s["h1"]))
        story.append(_hr(W_PT))
        story.append(_p(
            "說明：以下分析基於試卷答案版（參考答案卷）。依據香港課程發展議會"
            "《數學課程指引》（2017）逐題分析考核概念、難度及預測學生出錯原因。",
            s["small"],
        ))
        story.append(Spacer(1, 0.25 * cm))

        overview = exam.get("exam_overview", {})
        if overview:
            meta_rows = [
                ["題目總數", str(overview.get("total_questions", "—"))],
                ["試卷頁數", str(overview.get("total_pages", "—"))],
                ["整體難度", overview.get("estimated_difficulty", "—")],
                ["課程範疇", "、".join(overview.get("strands_tested", []))],
                ["涵蓋主題", "、".join(overview.get("topics_covered", []))],
            ]
            story.append(_build_table(
                [["項目", "內容"]] + meta_rows, [4 * cm, W_PT - 4 * cm], s,
                header_bg=colors.HexColor("#5a6fd6"),
            ))
            story.append(Spacer(1, 0.2 * cm))

        # Per-question table
        questions = exam.get("question_analysis", [])
        if questions:
            story.append(_p(f"◆ 逐題考核分析（共 {len(questions)} 題）", s["h2"]))
            story.append(_p(
                "每題根據香港課程發展議會《數學課程指引》（2017）標示考核範疇及學習目標",
                s["small"],
            ))
            data = [["題號", "頁", "題目內容", "考核主題", "課程範疇",
                     "分值", "難度", "正確答案", "預測學生錯誤"]]
            for q in questions:
                data.append([
                    q.get("question_ref", ""),
                    str(q.get("page", "")),
                    q.get("question_description", ""),
                    q.get("topic", ""),
                    q.get("strand", ""),
                    str(q.get("marks", "")) if q.get("marks") is not None else "—",
                    q.get("difficulty", ""),
                    q.get("correct_answer", ""),
                    q.get("predicted_errors", ""),
                ])
            story.append(_build_table(
                data,
                [1.2 * cm, 0.8 * cm, 3.2 * cm, 2.5 * cm,
                 2.0 * cm, 0.9 * cm, 1.2 * cm, 2.5 * cm, W_PT - 14.3 * cm],
                s,
            ))

            # Difficulty pie chart
            diff_counts = {}
            for q in questions:
                k = q.get("difficulty", "未知")
                diff_counts[k] = diff_counts.get(k, 0) + 1
            if len(diff_counts) >= 2:
                fig_diff = px.pie(
                    pd.DataFrame([{"難度": k, "題數": v} for k, v in diff_counts.items()]),
                    names="難度", values="題數",
                    color="難度",
                    color_discrete_map={"容易": "#43a047", "中等": "#f9a825", "困難": "#e53935"},
                    title="試題難度分佈",
                )
                fig_diff.update_layout(height=350, width=450, margin=dict(l=20, r=20, t=40, b=20))
                img = _chart(fig_diff, w_cm=9, h_cm=6)
                if img:
                    story.append(img)
                    story.append(_p("試題難度分佈圖", s["caption"]))
                    story.append(Spacer(1, 0.2 * cm))

        # Predicted error patterns
        patterns = exam.get("predicted_error_patterns", exam.get("error_patterns", []))
        if patterns:
            story.append(_p("◆ 預測學生錯誤模式（根據題目設計及課程要求推斷）", s["h2"]))
            data = [["預測錯誤模式", "相關數學概念", "涉及題目", "原因分析"]]
            for p in patterns:
                data.append([
                    p.get("pattern", ""),
                    p.get("related_concept", ""),
                    "、".join(p.get("affected_questions", [])),
                    p.get("reason", ""),
                ])
            story.append(_build_table(
                data, [3.5 * cm, 3 * cm, 2.5 * cm, W_PT - 9 * cm], s
            ))

        # Challenging areas
        challenging = exam.get("challenging_areas", exam.get("weak_areas", []))
        if challenging:
            story.append(Spacer(1, 0.2 * cm))
            story.append(_p("◆ 較具挑戰性的課程範疇", s["h2"]))
            data = [["主題", "課程範疇", "相關題目", "挑戰原因"]]
            for a in challenging:
                data.append([
                    a.get("topic", ""),
                    a.get("strand", ""),
                    "、".join(a.get("questions", [])),
                    a.get("reason", "") or a.get("evidence", ""),
                ])
            story.append(_build_table(
                data, [3 * cm, 2.5 * cm, 2.5 * cm, W_PT - 8 * cm], s
            ))

        story.append(PageBreak())

    # ── Section 3: Combined ───────────────────────────────────────────────────
    if "combined" in results and not results["combined"].get("parse_error"):
        combined = results["combined"]
        story.append(_p("三、綜合診斷分析", s["h1"]))
        story.append(_hr(W_PT))
        story.append(Spacer(1, 0.25 * cm))

        diag = combined.get("diagnostic_summary", {})
        if diag:
            story.append(_p("◆ 核心學習診斷", s["h2"]))
            story.append(_p(diag.get("overall_diagnosis", ""), s["body"]))
            story.append(Spacer(1, 0.15 * cm))
            story.append(_p("AQP 弱點與試卷關聯分析：", s["h2"]))
            story.append(_p(diag.get("aqp_exam_correlation", ""), s["body"]))

            mc = diag.get("misconception_vs_procedural", {})
            if mc:
                c_issues = mc.get("conceptual_issues", [])
                p_issues = mc.get("procedural_issues", [])
                max_len = max(len(c_issues), len(p_issues), 1)
                mc_data = [["概念性誤解", "程序性錯誤"]]
                for i in range(max_len):
                    mc_data.append([
                        c_issues[i] if i < len(c_issues) else "",
                        p_issues[i] if i < len(p_issues) else "",
                    ])
                story.append(_build_table(mc_data, [W_PT / 2, W_PT / 2], s))

            # Key weak question cross-reference
            kw = diag.get("key_weak_questions", [])
            if kw:
                story.append(Spacer(1, 0.2 * cm))
                story.append(_p("◆ AQP 弱題 × 試卷題目 對照表", s["h2"]))
                data = [["AQP 弱題", "全年級正確率", "對應試卷題號", "試卷題目內容", "關聯說明"]]
                for q in kw:
                    rate = q.get("aqp_correct_rate")
                    data.append([
                        q.get("aqp_question", ""),
                        f"{rate}%" if rate is not None else "—",
                        q.get("exam_question", ""),
                        q.get("exam_question_content", ""),
                        q.get("connection", ""),
                    ])
                story.append(_build_table(
                    data, [1.5 * cm, 2.2 * cm, 2.2 * cm, 4.5 * cm, W_PT - 10.4 * cm], s,
                    header_bg=C_WARNING,
                ))

        # Root cause analysis
        rca = combined.get("question_root_cause_analysis", [])
        rca_errors = [q for q in rca if q.get("correctness") != "正確"]
        if rca_errors:
            story.append(Spacer(1, 0.25 * cm))
            story.append(_p(f"◆ 逐題出錯根因分析（{len(rca_errors)} 題）", s["h2"]))
            data = [["題號", "試卷題目", "考核主題", "AQP 數據佐證", "根本原因", "錯誤類型"]]
            for q in rca_errors:
                data.append([
                    q.get("question_ref", ""),
                    q.get("question_content", "") or q.get("error_observed", ""),
                    q.get("topic", ""),
                    q.get("aqp_evidence", ""),
                    q.get("root_cause", ""),
                    q.get("error_type", ""),
                ])
            story.append(_build_table(
                data,
                [1.2 * cm, 3 * cm, 2.5 * cm, 2.8 * cm, 4 * cm, W_PT - 13.5 * cm],
                s,
            ))

        # Consolidated weak areas
        weak_areas = combined.get("consolidated_weak_areas", [])
        if weak_areas:
            story.append(Spacer(1, 0.25 * cm))
            story.append(_p("◆ 重點需改善範疇", s["h2"]))
            data = [["優先級", "課程範疇", "相關主題", "AQP 佐證", "試卷佐證", "深層原因"]]
            for a in weak_areas:
                data.append([
                    a.get("priority_level", ""),
                    a.get("strand", ""),
                    "、".join(a.get("topics", [])),
                    a.get("aqp_evidence", ""),
                    a.get("exam_evidence", ""),
                    a.get("root_cause_analysis", ""),
                ])
            story.append(_build_table(
                data,
                [1.5 * cm, 2 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, W_PT - 11 * cm],
                s,
            ))

        # Remediation plan
        plan = combined.get("remediation_plan", [])
        if plan:
            story.append(Spacer(1, 0.25 * cm))
            story.append(_p("◆ 補救教學計劃", s["h2"]))
            for phase in plan:
                story.append(_p(
                    f"📅 {phase.get('phase', '')} — {phase.get('target_weakness', '')}",
                    s["h2"],
                ))
                story.append(_p(
                    f"教學方法：{phase.get('teaching_approach', '')}", s["body"]
                ))
                for act in phase.get("practice_activities", []):
                    story.append(_p(f"• {act}", s["body"]))
                if phase.get("success_criteria"):
                    story.append(_p(f"達成標準：{phase['success_criteria']}", s["small"]))
                if phase.get("assessment"):
                    story.append(_p(f"評估方法：{phase['assessment']}", s["small"]))
                story.append(Spacer(1, 0.1 * cm))

        story.append(PageBreak())

    # ── Section 4: Recommendations ────────────────────────────────────────────
    all_recs = []
    for key in ("aqp", "exam", "combined"):
        if key in results and not results[key].get("parse_error"):
            for r in results[key].get("recommendations", []):
                r["_source"] = key
                all_recs.append(r)

    if all_recs:
        story.append(_p("四、優先改善建議", s["h1"]))
        story.append(_hr(W_PT))
        p_order = {"高": 0, "緊急": 0, "中": 1, "重要": 1, "低": 2, "一般": 2}
        all_recs.sort(key=lambda x: p_order.get(x.get("priority", "低"), 2))
        data = [["優先級", "課程範疇", "建議行動", "建議練習資源"]]
        for rec in all_recs:
            action = rec.get("action") or rec.get("specific_action", "")
            res_text = rec.get("resources") or rec.get("suggested_exercises", "")
            data.append([
                rec.get("priority", ""),
                rec.get("area", ""),
                action,
                res_text,
            ])
        story.append(_build_table(
            data, [1.5 * cm, 3 * cm, 6 * cm, W_PT - 10.5 * cm], s
        ))

    # Priority interventions
    if "combined" in results and not results["combined"].get("parse_error"):
        priority = results["combined"].get("priority_interventions", [])
        if priority:
            story.append(Spacer(1, 0.4 * cm))
            story.append(_p("◆ 最優先介入行動（立即執行）", s["h2"]))
            data = [["優先", "弱點（附數據）", "優先原因", "即時行動（下星期）", "教材建議"]]
            for item in sorted(priority, key=lambda x: x.get("rank", 99)):
                data.append([
                    f"#{item.get('rank', '')}",
                    item.get("weakness", ""),
                    item.get("reason", ""),
                    item.get("immediate_action", ""),
                    item.get("resources", ""),
                ])
            story.append(_build_table(
                data, [1.2 * cm, 2.8 * cm, 3.5 * cm, 4 * cm, W_PT - 11.5 * cm], s,
                header_bg=C_DANGER,
            ))

    # ── Footer ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 0.8 * cm))
    story.append(HRFlowable(width=W_PT, color=C_GREY, thickness=0.5, spaceAfter=4))
    story.append(_p(
        f"本報告由香港小學數學全年級表現分析系統生成  ·  由 Qwen AI 驅動  ·  生成日期：{now_str}",
        s["caption"],
    ))
    story.append(_p(
        "課程參考：香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版",
        s["caption"],
    ))

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Student batch report — builds a teacher-facing class report from per-student
# analysis aggregated results + AI insights.
# ---------------------------------------------------------------------------

def build_student_report(
    class_results: Dict,
    insights: Dict,
    grade: str,
    notes: str = "",
) -> bytes:
    """
    Build a full PDF class report from aggregated student analysis.

    Parameters
    ----------
    class_results : Dict  — output of aggregate_student_results()
    insights      : Dict  — output of MathAnalyzer.generate_class_insights()
    grade         : str   — e.g. "P4"
    notes         : str   — free-text label, e.g. "2024-25 上學期"
    """
    font = _ensure_font()
    s = _styles(font)
    buf = io.BytesIO()
    W_PT = A4[0] - 3.6 * cm  # usable width in points

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
    )

    def _p(text, style):
        return Paragraph(str(text), style)

    story = []

    # ── Title block ──────────────────────────────────────────────────
    story.append(Table(
        [[_p(f"📊 {grade} 全班數學表現分析報告", s["title"])]],
        colWidths=[W_PT],
        style=TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), C_PRIMARY),
            ("ROWPADDING", (0, 0), (-1, -1), 14),
            ("ROUNDEDCORNERS", [6]),
        ]),
    ))
    story.append(Spacer(1, 10))

    total_s = class_results.get("total_students", 0)
    avg = class_results.get("class_average", 0)
    now_str = datetime.now().strftime("%Y-%m-%d")
    notes_str = notes or "—"

    meta = [
        ["年級", grade, "學生人數", str(total_s)],
        ["全班平均分", f"{avg:.1f}%", "備註", notes_str],
        ["報告日期", now_str, "課程參考", "《數學課程指引》2017"],
    ]
    meta_rows = [[_p(c, s["th"] if j % 2 == 0 else s["td"]) for j, c in enumerate(row)] for row in meta]
    story.append(_build_table(meta_rows, [3.5 * cm, 5.5 * cm, 3.5 * cm, 5.5 * cm], s))
    story.append(Spacer(1, 14))

    # ── Overall diagnosis ────────────────────────────────────────────
    if insights and not insights.get("parse_error"):
        diag = insights.get("overall_diagnosis", "")
        if diag:
            story.append(_p("🔬 AI 全班診斷摘要", s["h1"]))
            story.append(_p(diag, s["body"]))
            story.append(Spacer(1, 8))

    # ── Score distribution chart ─────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=1, color=C_PRIMARY))
    story.append(_p("第一部分：全班成績概覽", s["h1"]))

    dist = class_results.get("class_distribution", {})
    col_map = {
        "優秀(≥85%)": "#43a047",
        "良好(70-84%)": "#1e88e5",
        "一般(55-69%)": "#f9a825",
        "需要改善(<55%)": "#e53935",
    }
    if dist:
        df_dist = pd.DataFrame([
            {"等級": k, "人數": v, "color": col_map.get(k, "#90a4ae")}
            for k, v in dist.items() if v > 0
        ])
        fig_pie = go.Figure(go.Pie(
            labels=df_dist["等級"], values=df_dist["人數"],
            marker_colors=df_dist["color"],
            textinfo="label+value+percent", hole=0.35,
        ))
        fig_pie.update_layout(
            title="全班成績等級分佈", showlegend=False,
            margin=dict(l=10, r=10, t=40, b=10),
        )
        img_pie = _chart(fig_pie, 9, 7)
        if img_pie:
            story.append(img_pie)
        else:
            dist_rows = [["等級", "人數", "百分比"]] + [
                [k, str(v), f"{round(100*v/total_s)}%" if total_s else "—"]
                for k, v in dist.items()
            ]
            dist_para = [[_p(c, s["th"] if i == 0 else s["td"]) for c in row]
                         for i, row in enumerate(dist_rows)]
            story.append(_build_table(dist_para, [7 * cm, 3 * cm, 3 * cm], s))

    # Metric summary bar
    weak_q = class_results.get("weak_questions", [])
    need_help = dist.get("需要改善(<55%)", 0)
    summary_rows = [
        [_p("全班平均分", s["th"]), _p(f"{avg:.1f}%", s["td"]),
         _p("需要關注學生", s["th"]), _p(f"{need_help} 人", s["td"])],
        [_p("弱題數目（<60%）", s["th"]), _p(str(len(weak_q)), s["td"]),
         _p("分析學生總數", s["th"]), _p(str(total_s), s["td"])],
    ]
    story.append(_build_table(summary_rows, [4 * cm, 4 * cm, 4 * cm, 4 * cm], s))
    story.append(Spacer(1, 14))

    # ── Student ranking table ────────────────────────────────────────
    story.append(PageBreak())
    story.append(_p("第二部分：學生成績排名", s["h1"]))
    ranking = class_results.get("student_ranking", [])
    if ranking:
        rank_header = [_p(h, s["th"]) for h in ["排名", "學生", "得分率", "得分", "表現等級"]]
        rank_rows = [rank_header]
        lv_icon = {"優秀(≥85%)": "●", "良好(70-84%)": "○", "一般(55-69%)": "△", "需要改善(<55%)": "▲"}
        for r in ranking:
            icon = lv_icon.get(r.get("performance_level", ""), "")
            rank_rows.append([
                _p(str(r.get("rank", "")), s["td"]),
                _p(r.get("student_name", ""), s["td"]),
                _p(f"{r.get('percentage', 0):.1f}%", s["td"]),
                _p(f"{r.get('total_marks_awarded','—')} / {r.get('total_marks_possible','—')}", s["td"]),
                _p(f"{icon} {r.get('performance_level','')}", s["td"]),
            ])
        story.append(_build_table(
            rank_rows,
            [2 * cm, 5 * cm, 3 * cm, 3.5 * cm, 4.5 * cm],
            s,
        ))

    # ── Strand radar chart ───────────────────────────────────────────
    story.append(PageBreak())
    story.append(_p("第三部分：各課程範疇表現", s["h1"]))
    strand_stats = class_results.get("strand_stats", [])
    if strand_stats:
        cats = [ss["strand"] for ss in strand_stats]
        vals = [ss["class_average_rate"] for ss in strand_stats]
        if len(cats) >= 3:
            fig_radar = go.Figure(go.Scatterpolar(
                r=vals + [vals[0]], theta=cats + [cats[0]],
                fill="toself",
                fillcolor="rgba(102,126,234,0.25)",
                line=dict(color="rgb(102,126,234)", width=2),
            ))
            fig_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                showlegend=False, margin=dict(l=50, r=50, t=40, b=40),
            )
            img_radar = _chart(fig_radar, 12, 9)
            if img_radar:
                story.append(img_radar)

        strand_header = [_p(h, s["th"]) for h in ["範疇", "全班正確率", "狀態", "涉及題目"]]
        strand_rows = [strand_header]
        for ss in strand_stats:
            rate = ss["class_average_rate"]
            status = ss["status"]
            color_cell = C_DANGER if status == "弱項" else C_WARNING if status == "一般" else C_SUCCESS
            strand_rows.append([
                _p(ss["strand"], s["td"]),
                _p(f"{rate}%", s["td"]),
                _p(status, s["td"]),
                _p("、".join(ss.get("questions", [])[:8]), s["small"]),
            ])
        t = _build_table(strand_rows, [5 * cm, 3 * cm, 2.5 * cm, 7.5 * cm], s)
        story.append(t)

    # ── Per-question correct rate chart ─────────────────────────────
    story.append(PageBreak())
    story.append(_p("第四部分：逐題全班正確率", s["h1"]))
    q_stats = class_results.get("question_stats", [])
    if q_stats:
        q_disp = sorted(q_stats, key=lambda x: x["question_ref"])
        q_refs = [q["question_ref"] for q in q_disp]
        q_rates = [q["class_correct_rate"] for q in q_disp]
        q_colors = ["#e53935" if r < 40 else "#f9a825" if r < 60 else "#43a047" for r in q_rates]

        fig_bar = go.Figure(go.Bar(
            x=q_refs, y=q_rates,
            marker_color=q_colors,
            text=[f"{r}%" for r in q_rates],
            textposition="outside",
        ))
        fig_bar.add_hline(y=60, line_dash="dash", line_color="red",
                          annotation_text="60% 基準")
        fig_bar.update_layout(
            title="各題全班正確率",
            yaxis=dict(range=[0, 110], title="正確率 (%)"),
            xaxis_title="題目",
            margin=dict(l=10, r=10, t=50, b=40),
        )
        img_bar = _chart(fig_bar, 16, 8)
        if img_bar:
            story.append(img_bar)

        # Question table
        q_header = [_p(h, s["th"]) for h in ["題目", "考核主題", "範疇", "全班正確率", "正確人數", "常見錯誤"]]
        q_rows = [q_header]
        for q in q_disp:
            rate = q["class_correct_rate"]
            style_cell = s["warn"] if rate < 60 else s["td"]
            q_rows.append([
                _p(q["question_ref"], s["td"]),
                _p(q.get("topic", ""), s["small"]),
                _p(q.get("strand", ""), s["small"]),
                _p(f"{rate}%", style_cell),
                _p(f"{q['class_correct_count']} / {total_s}", s["td"]),
                _p("；".join(q.get("common_errors", [])[:2]) or "—", s["small"]),
            ])
        story.append(_build_table(
            q_rows,
            [2 * cm, 4 * cm, 3.5 * cm, 2.5 * cm, 2.5 * cm, 3.5 * cm],
            s,
        ))

    # ── Weak question ranking ────────────────────────────────────────
    if weak_q:
        story.append(PageBreak())
        story.append(_p("第五部分：弱題排行榜（全班正確率 < 60%）", s["h1"]))
        story.append(_p(
            f"共找出 {len(weak_q)} 條弱題，以下按正確率由低至高排列。正確率低於 40% 為嚴重弱項。",
            s["body"],
        ))

        wq_header = [_p(h, s["th"]) for h in ["排名", "題目", "全班正確率", "考核主題", "範疇", "常見錯誤"]]
        wq_rows = [wq_header]
        for q in weak_q:
            rate = q["class_correct_rate"]
            flag = "🔴" if rate < 40 else "🟡"
            wq_rows.append([
                _p(str(q.get("rank", "")), s["td"]),
                _p(q["question_ref"], s["td"]),
                _p(f"{flag} {rate}%", s["warn"]),
                _p(q.get("topic", ""), s["small"]),
                _p(q.get("strand", ""), s["small"]),
                _p("；".join(q.get("common_errors", [])[:2]) or "—", s["small"]),
            ])
        story.append(_build_table(
            wq_rows,
            [2 * cm, 2.5 * cm, 3 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm],
            s,
        ))

    # ── AI teaching recommendations ──────────────────────────────────
    if insights and not insights.get("parse_error"):
        story.append(PageBreak())
        story.append(_p("第六部分：AI 教學建議", s["h1"]))

        ws_analysis = insights.get("weak_strand_analysis", [])
        if ws_analysis:
            story.append(_p("弱項範疇深度分析", s["h2"]))
            for ws in ws_analysis:
                story.append(_p(
                    f"【{ws.get('strand','')}】全班正確率 {ws.get('class_average_rate','')}%",
                    s["body"],
                ))
                for issue in ws.get("key_issues", []):
                    story.append(_p(f"• {issue}", s["body"]))
                if ws.get("misconception"):
                    story.append(_p(f"可能的概念誤解：{ws['misconception']}", s["warn"]))
                if ws.get("curriculum_link"):
                    story.append(_p(f"課程連結：{ws['curriculum_link']}", s["small"]))
                story.append(Spacer(1, 6))

        recs = insights.get("teaching_recommendations", [])
        if recs:
            story.append(_p("補救教學建議", s["h2"]))
            rec_header = [_p(h, s["th"]) for h in ["優先級", "針對範疇", "教學策略", "建議時間"]]
            rec_rows = [rec_header]
            for rec in sorted(recs, key=lambda r: {"高": 0, "中": 1, "低": 2}.get(r.get("priority", "低"), 2)):
                rec_rows.append([
                    _p(rec.get("priority", ""), s["td"]),
                    _p(rec.get("strand", ""), s["td"]),
                    _p(rec.get("strategy", ""), s["small"]),
                    _p(rec.get("timeline", ""), s["small"]),
                ])
            story.append(_build_table(rec_rows, [2.5 * cm, 4 * cm, 7 * cm, 4.5 * cm], s))

        if insights.get("attention_students_note"):
            story.append(Spacer(1, 8))
            story.append(_p("個別關注學生", s["h2"]))
            story.append(_p(insights["attention_students_note"], s["body"]))

        if insights.get("positive_findings"):
            story.append(Spacer(1, 8))
            story.append(_p("全班亮點", s["h2"]))
            story.append(_p(insights["positive_findings"], s["success"]))

    # ── Footer ───────────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=1, color=C_GREY))
    story.append(_p(f"報告生成日期：{now_str} | 由 Qwen AI 驅動 | 小學數學學生表現分析系統", s["caption"]))
    story.append(_p("課程參考：香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版", s["caption"]))

    doc.build(story)
    return buf.getvalue()
