"""
中華基督教會基慈小學 · 數學學生表現分析系統
Streamlit web application — powered by Qwen AI
"""

import json
import math
import os
import tempfile

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from analyzer import MathAnalyzer, aggregate_student_results
from curriculum_hk import CURRICULUM_STRANDS
from file_processor import (
    FileProcessor,
    check_pdf_support,
    get_pdf_page_count,
    split_student_papers,
)
from html_exporter import build_student_html_report
from pdf_exporter import build_pdf, build_student_report
from practice_html import build_practice_worksheets_html

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="中華基督教會基慈小學 · 數學學生表現分析系統",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
.main-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 22px 28px;
    border-radius: 12px;
    color: white;
    margin-bottom: 24px;
}
.main-header h1 { margin: 0 0 6px 0; font-size: 1.8rem; }
.main-header p  { margin: 0; opacity: 0.9; font-size: 0.95rem; }

.card {
    background: #f8f9fa;
    padding: 14px 16px;
    border-radius: 8px;
    border-left: 4px solid #667eea;
    margin: 6px 0;
}
.card-red   { border-left-color: #e53935; background: #fff3f3; }
.card-green { border-left-color: #43a047; background: #f0fff4; }
.card-blue  { border-left-color: #1e88e5; background: #f0f8ff; }
.card-yellow{ border-left-color: #f9a825; background: #fffde7; }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(
    """
<div class="main-header">
  <h1>📊 中華基督教會基慈小學 · 數學學生表現分析系統</h1>
  <p>上傳全班學生試卷 PDF · AI 逐份批改 · 自動生成全班弱點診斷報告 · 基於香港小學數學課程綱要</p>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 系統設定")
    api_key = st.text_input(
        "🔑 Qwen International API Key",
        type="password",
        placeholder="sk-...",
        help="前往 dashscope.aliyuncs.com 取得 API Key",
    )

    st.markdown("---")
    st.markdown("### 📚 分析功能")
    st.markdown(
        """
- � 全班試卷 PDF 一鍵分割
- 🤖 AI 逐份批改（視覺識別）
- 📈 全班成績分佈圖
- 🔴 弱題熱圖分析
- 🎯 個人及全班弱點識別
- 📅 補救教學建議
- 📥 PDF 報告匯出
"""
    )

    st.markdown("---")
    st.markdown("### 📂 支援格式")
    st.markdown(
        """
**學生試卷：** PDF（全班合併）  
**答案鍵：** PDF、JPG、PNG
"""
    )

    st.markdown("---")
    st.caption(check_pdf_support())
    st.caption("由 Qwen AI 驅動")

# ---------------------------------------------------------------------------
# Guard — require API key
# ---------------------------------------------------------------------------
if not api_key:
    st.info("👈 請在左側欄輸入您的 **Qwen API Key** 以開始使用。")
    st.stop()

# ---------------------------------------------------------------------------
# Mode selector — removed (single mode only)
# ---------------------------------------------------------------------------

st.markdown("---")

# ===========================================================================
# Student paper batch analysis
# ===========================================================================
if True:

    st.markdown("### 📁 上傳全班試卷")
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### 📄 全班試卷 PDF（必填）")
        student_pdf_file = st.file_uploader(
            "將全班學生試卷掃描合併成一份 PDF 上傳",
            type=["pdf"],
            key="student_pdf",
            help="例如：全班25份試卷，每份4頁，共100頁。AI 會自動分割並逐份批改。",
        )
        if student_pdf_file:
            st.success(f"✅ 已選擇：{student_pdf_file.name}")

    with col_b:
        st.markdown("#### 📋 答案鍵（可選，提高評分準確度）")
        answer_key_file = st.file_uploader(
            "上傳試卷答案鍵（PDF / JPG / PNG）",
            type=["pdf", "jpg", "jpeg", "png"],
            key="answer_key",
            help="如提供答案鍵，AI 可更準確地批改並給分。沒有也可自動評核。",
        )
        if answer_key_file:
            st.success(f"✅ 已選擇：{answer_key_file.name}")

    # ── Configuration ────────────────────────────────────────────────
    st.markdown("### ⚙️ 試卷設定")
    cfg_col1, cfg_col2, cfg_col3 = st.columns(3)

    with cfg_col1:
        grade = st.selectbox("年級", ["P1", "P2", "P3", "P4", "P5", "P6"], index=3, key="grade_s")

    with cfg_col2:
        pages_per_student = st.number_input(
            "每位學生的試卷頁數",
            min_value=1, max_value=20, value=4,
            help="例如每份試卷4頁，25名學生 = 100頁 PDF",
        )

    with cfg_col3:
        class_label = st.text_input("備註（可選）", placeholder="例：2024-25 上學期", key="label_s")

    # Preview estimated student count
    if student_pdf_file:
        pdf_bytes_preview = student_pdf_file.read()
        student_pdf_file.seek(0)
        total_pages = get_pdf_page_count(pdf_bytes_preview)
        est_students = math.ceil(total_pages / pages_per_student)
        st.info(
            f"📄 共 **{total_pages}** 頁  ·  每人 **{pages_per_student}** 頁  "
            f"·  估計 **{est_students}** 位學生"
        )

    # Optional student names
    st.markdown("#### 🏷️ 學生姓名（可選）")
    names_text = st.text_area(
        "每行輸入一個學生名字（按順序對應 PDF 頁序）",
        height=120,
        placeholder="陳大文\n李小明\n黃美玲\n（留空則自動命名為學生1、學生2…）",
        key="names",
    )

    # ── Analyse button ────────────────────────────────────────────────
    st.markdown("---")
    _, btn_col, _ = st.columns([1, 2, 1])
    with btn_col:
        analyse_btn = st.button(
            "🔍 開始批改全班試卷",
            type="primary",
            use_container_width=True,
            disabled=not student_pdf_file,
        )

    # ── Run analysis ──────────────────────────────────────────────────
    if analyse_btn:
        st.session_state.pop("student_results", None)
        st.session_state.pop("class_agg", None)
        st.session_state.pop("class_insights", None)
        st.session_state.pop("s_html_bytes", None)
        st.session_state.pop("s_html_stem", None)

        analyzer = MathAnalyzer(api_key)
        processor = FileProcessor()

        # Parse student names
        student_names = [n.strip() for n in names_text.strip().splitlines() if n.strip()]

        progress = st.progress(0)
        status = st.empty()
        error_log = []

        try:
            # Step 1: Split PDF
            status.info("✂️ 正在分割試卷 PDF…")
            pdf_bytes = student_pdf_file.read()
            student_chunks = split_student_papers(pdf_bytes, int(pages_per_student))
            total_students = len(student_chunks)
            progress.progress(5)

            # Step 2: Optional — extract answer key question schema
            question_schema = []
            if answer_key_file:
                status.info("📋 正在分析答案鍵…")
                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=os.path.splitext(answer_key_file.name)[1],
                ) as tmp:
                    tmp.write(answer_key_file.read())
                    tmp_path = tmp.name
                try:
                    key_data = processor.process_exam(tmp_path)
                    key_result = analyzer.analyze_exam(key_data, grade)
                    question_schema = key_result.get("question_analysis", [])
                    # Simplify schema for prompting (keep only essential fields)
                    question_schema = [
                        {
                            "question_ref": q.get("question_ref", ""),
                            "topic": q.get("topic", ""),
                            "strand": q.get("strand", ""),
                            "marks": q.get("marks"),
                            "correct_answer": q.get("correct_answer", ""),
                            "solution_method": q.get("solution_method", ""),
                        }
                        for q in question_schema
                    ]
                    status.success(f"✅ 已從答案鍵識別 {len(question_schema)} 題")
                finally:
                    os.unlink(tmp_path)
                progress.progress(15)

            # Step 3: Analyze each student
            all_student_results = []
            base_prog = 15
            prog_per_student = (85 - base_prog) / max(total_students, 1)

            for chunk in student_chunks:
                idx = chunk["student_index"]
                name = (
                    student_names[idx - 1]
                    if idx - 1 < len(student_names)
                    else f"學生{idx}"
                )
                pages = chunk["page_range"]
                status.info(
                    f"🤖 正在批改 **{name}** 的試卷（第 {pages[0]}–{pages[1]} 頁）"
                    f"　{idx}/{total_students}"
                )
                try:
                    result = analyzer.analyze_student_paper(
                        chunk["images"], question_schema, grade, name
                    )
                    result["student_name"] = name   # always use the roster name
                    result["student_index"] = idx
                    all_student_results.append(result)
                except Exception as exc:
                    error_log.append(f"{name}：{exc}")
                    all_student_results.append({
                        "student_name": name,
                        "student_index": idx,
                        "parse_error": True,
                        "error": str(exc),
                    })
                progress.progress(min(99, int(base_prog + idx * prog_per_student)))

            # Step 4: Aggregate
            status.info("📊 正在計算全班統計數據…")
            expected_qs = [q["question_ref"] for q in question_schema] if question_schema else []
            class_agg = aggregate_student_results(all_student_results, expected_qs)

            # Step 5: AI insights
            status.info("🧠 AI 正在生成教學診斷建議…")
            class_insights = analyzer.generate_class_insights(class_agg, grade)

            progress.progress(100)
            status.success(f"✅ 完成！成功批改 {len([r for r in all_student_results if not r.get('parse_error')])} / {total_students} 份試卷")

            # Save to session state
            st.session_state["student_results"] = all_student_results
            st.session_state["class_agg"] = class_agg
            st.session_state["class_insights"] = class_insights
            st.session_state["s_grade"] = grade
            st.session_state["s_label"] = class_label
            if error_log:
                st.warning("⚠️ 部分試卷分析失敗：\n" + "\n".join(error_log))

        except Exception as exc:
            progress.empty()
            status.empty()
            st.error(f"❌ 分析時發生錯誤：{exc}")
            st.exception(exc)

    # ── Display results ───────────────────────────────────────────────
    if "class_agg" not in st.session_state:
        st.stop()

    agg = st.session_state["class_agg"]
    insights = st.session_state.get("class_insights", {})
    all_results = st.session_state.get("student_results", [])
    s_grade = st.session_state.get("s_grade", "")
    s_label = st.session_state.get("s_label", "")
    label_str = f"（{s_label}）" if s_label else ""

    if agg.get("error"):
        st.error(agg["error"])
        st.stop()

    st.markdown(f"---\n## 📊 {s_grade}{label_str} 全班數學表現分析報告")

    rtabs = st.tabs([
        "📋 整體概覽",
        "🏅 學生成績",
        "✏️ 自動批改",
        "📝 逐題分析",
        "🔥 弱點熱圖",
        "🎯 弱點診斷",
        "💡 教學建議",
        "� 弱點練習",
        "�📥 匯出報告",
    ])

    # ── Tab 0: Overview ───────────────────────────────────────────────
    with rtabs[0]:
        total_s = agg.get("total_students", 0)
        avg = agg.get("class_average", 0)
        dist = agg.get("class_distribution", {})
        weak_q = agg.get("weak_questions", [])
        strand_stats = agg.get("strand_stats", [])

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("分析學生數", f"{total_s} 人")
        m2.metric("全班平均分", f"{avg:.1f}%")
        m3.metric("弱題數目（<60%）", len(weak_q))
        need_help = dist.get("需要改善(<55%)", 0)
        m4.metric("需要關注學生", f"{need_help} 人")

        if insights and not insights.get("parse_error"):
            st.info(f"🔬 **診斷摘要：** {insights.get('overall_diagnosis', '')}")

        # Score distribution pie + bar
        col_pie, col_bar = st.columns(2)
        with col_pie:
            if dist:
                df_pie = pd.DataFrame([
                    {"表現等級": k, "人數": v} for k, v in dist.items() if v > 0
                ])
                color_map = {
                    "優秀(≥85%)": "#43a047",
                    "良好(70-84%)": "#1e88e5",
                    "一般(55-69%)": "#f9a825",
                    "需要改善(<55%)": "#e53935",
                }
                fig = px.pie(
                    df_pie, names="表現等級", values="人數",
                    color="表現等級", color_discrete_map=color_map,
                    title="全班成績等級分佈",
                )
                fig.update_traces(textinfo="label+value+percent")
                st.plotly_chart(fig, use_container_width=True)

        with col_bar:
            # Histogram of student percentages
            pcts = [r.get("percentage", 0) for r in all_results if not r.get("parse_error")]
            if pcts:
                fig = px.histogram(
                    x=pcts, nbins=10,
                    labels={"x": "得分率 (%)", "y": "學生人數"},
                    title="全班得分率分佈直方圖",
                    color_discrete_sequence=["#667eea"],
                )
                fig.add_vline(x=avg, line_dash="dash", line_color="red",
                              annotation_text=f"平均 {avg:.1f}%")
                fig.update_xaxes(range=[0, 100])
                st.plotly_chart(fig, use_container_width=True)

        # Strand radar
        if strand_stats:
            st.markdown("### 📈 各課程範疇全班正確率")
            cats = [s["strand"] for s in strand_stats]
            vals = [s["class_average_rate"] for s in strand_stats]
            if len(cats) >= 3:
                fig = go.Figure(go.Scatterpolar(
                    r=vals + [vals[0]],
                    theta=cats + [cats[0]],
                    fill="toself",
                    fillcolor="rgba(102,126,234,0.25)",
                    line=dict(color="rgb(102,126,234)", width=2),
                ))
                fig.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                    showlegend=False, height=380,
                    margin=dict(l=60, r=60, t=40, b=40),
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                df_strand = pd.DataFrame(strand_stats)
                fig = px.bar(
                    df_strand, x="strand", y="class_average_rate",
                    color="class_average_rate", color_continuous_scale="RdYlGn",
                    range_y=[0, 100], labels={"strand": "範疇", "class_average_rate": "正確率 (%)"},
                    title="各課程範疇全班正確率",
                )
                fig.add_hline(y=60, line_dash="dash", line_color="red", annotation_text="60% 基準線")
                st.plotly_chart(fig, use_container_width=True)

    # ── Tab 1: Student ranking ────────────────────────────────────────
    with rtabs[1]:
        ranking = agg.get("student_ranking", [])
        if ranking:
            st.markdown(f"### 🏅 全班成績排名（共 {len(ranking)} 位學生）")
            sicon = {"優秀(≥85%)": "🟢", "良好(70-84%)": "🔵", "一般(55-69%)": "🟡", "需要改善(<55%)": "🔴", "分析失敗": "❌"}
            rows = [
                {
                    "排名": s["rank"],
                    "學生": s["student_name"],
                    "得分率": f"{s['percentage']:.1f}%",
                    "得分": f"{s.get('total_marks_awarded','—')} / {s.get('total_marks_possible','—')}",
                    "表現等級": f"{sicon.get(s['performance_level'],'⚪')} {s['performance_level']}",
                }
                for s in ranking
            ]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Bar chart sorted by score
            df_rank = pd.DataFrame([
                {"學生": s["student_name"], "得分率": s["percentage"]}
                for s in sorted(ranking, key=lambda x: x["percentage"])
            ])
            fig = px.bar(
                df_rank, x="得分率", y="學生", orientation="h",
                color="得分率", color_continuous_scale="RdYlGn", range_x=[0, 100],
                title="全班學生得分率排行",
                labels={"得分率": "得分率 (%)", "學生": ""},
            )
            fig.add_vline(x=agg["class_average"], line_dash="dash", line_color="blue",
                          annotation_text=f"平均 {agg['class_average']}%")
            fig.add_vline(x=60, line_dash="dot", line_color="red",
                          annotation_text="60% 基準")
            fig.update_layout(height=max(300, 22 * len(df_rank)), margin=dict(l=120))
            st.plotly_chart(fig, use_container_width=True)

    # ── Tab 2: Auto-marking (wrong answers per student) ──────────────
    with rtabs[2]:
        student_results = agg.get("student_results", [])
        q_stats = agg.get("question_stats", [])
        if student_results:
            st.markdown("### ✏️ 自動批改 — 各學生答錯題目")
            st.caption("只列出每位學生答錯的題目，方便老師用紅筆在紙本工作紙上批改。")

            for student in student_results:
                name = student.get("student_name", "未知")
                if student.get("parse_error"):
                    st.warning(f"**{name}** — 分析失敗，無法取得答題結果")
                    continue

                q_results = student.get("question_results", [])
                wrong = [q for q in q_results if q.get("is_correct") is False]

                pct = student.get("percentage", 0)
                total_q = len(q_results)
                wrong_count = len(wrong)
                correct_count = total_q - wrong_count

                if not wrong:
                    st.success(f"**{name}** — ✅ 全部答對（{total_q}/{total_q}）")
                    continue

                color = "red" if pct < 55 else "orange" if pct < 70 else "blue"
                with st.expander(
                    f"❌ {name}　—　答錯 {wrong_count} 題 / 共 {total_q} 題　（得分率 {pct:.0f}%）",
                    expanded=(wrong_count >= 3),
                ):
                    rows = []
                    for q in wrong:
                        ref = q.get("question_ref", "")
                        rows.append({
                            "題目": ref,
                            "考核主題": q.get("topic", ""),
                            "學生答案": q.get("student_answer", "—"),
                            "正確答案": q.get("correct_answer", "—"),
                            "得分": f"{q.get('marks_awarded', 0)} / {q.get('marks_possible', '')}",
                            "錯誤類型": q.get("error_type", "") or "",
                            "錯誤說明": q.get("error_description", "") or "",
                        })
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Summary table: student × wrong question quick reference
            st.markdown("---")
            st.markdown("### 📋 全班答錯題目一覽表")
            st.caption("每格顯示 ❌ 代表該學生答錯該題，空白代表答對或未作答。")
            all_refs = [q["question_ref"] for q in q_stats] if q_stats else []
            if all_refs:
                summary_rows = []
                for student in student_results:
                    name = student.get("student_name", "未知")
                    if student.get("parse_error"):
                        continue
                    q_map = {
                        str(q.get("question_ref", "")): q.get("is_correct")
                        for q in student.get("question_results", [])
                    }
                    row = {"學生": name}
                    for ref in all_refs:
                        val = q_map.get(str(ref))
                        row[ref] = "❌" if val is False else ""
                    summary_rows.append(row)
                if summary_rows:
                    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)
        else:
            st.info("尚未有學生分析結果。")

    # ── Tab 3: Per-question stats ─────────────────────────────────────
    with rtabs[3]:
        q_stats = agg.get("question_stats", [])
        if q_stats:
            st.markdown(f"### 📝 逐題全班正確率（共 {len(q_stats)} 題）")

            # q_stats already in schema/natural order from aggregate_student_results
            q_display = q_stats

            rows = [
                {
                    "題目": q["question_ref"],
                    "考核主題": q.get("topic", ""),
                    "範疇": q.get("strand", ""),
                    "全班正確率": f"{q['class_correct_rate']}%",
                    "正確人數": f"{q['class_correct_count']} / {agg.get('valid_students', agg['total_students'])}",
                    "平均得分": q.get("class_average_marks") or "—",
                    "常見錯誤": "；".join(q.get("common_errors", [])[:2]) or "—",
                }
                for q in q_display
            ]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Correct-rate bar chart
            df_q = pd.DataFrame([
                {"題目": q["question_ref"], "正確率": q["class_correct_rate"]}
                for q in q_display
            ])
            fig = px.bar(
                df_q, x="題目", y="正確率",
                color="正確率", color_continuous_scale="RdYlGn", range_y=[0, 100],
                title="各題全班正確率",
                labels={"正確率": "正確率 (%)"},
            )
            fig.add_hline(y=60, line_dash="dash", line_color="red",
                          annotation_text="60% 基準線")
            st.plotly_chart(fig, use_container_width=True)

    # ── Tab 4: Heatmap ────────────────────────────────────────────────
    with rtabs[4]:
        q_stats = agg.get("question_stats", [])
        student_results = agg.get("student_results", [])
        if q_stats and student_results:
            st.markdown("### 🔥 學生 × 題目 答對熱圖")
            st.caption("🟢 答對　🔴 答錯　⬜ 未作答")

            all_refs = [q["question_ref"] for q in q_stats]  # already in correct order
            student_names_ordered = [s.get("student_name", f"學生{i+1}") for i, s in enumerate(student_results)]

            z_matrix = []
            text_matrix = []
            for student in student_results:
                q_map = {
                    str(q.get("question_ref", "")): q.get("is_correct")
                    for q in student.get("question_results", [])
                }
                row_z = []
                row_t = []
                for ref in all_refs:
                    val = q_map.get(str(ref))
                    if val is True:
                        row_z.append(1)
                        row_t.append("✓")
                    elif val is False:
                        row_z.append(0)
                        row_t.append("✗")
                    else:
                        row_z.append(0.5)
                        row_t.append("—")
                z_matrix.append(row_z)
                text_matrix.append(row_t)

            fig = go.Figure(data=go.Heatmap(
                z=z_matrix,
                x=all_refs,
                y=student_names_ordered,
                colorscale=[[0, "#e53935"], [0.45, "#ffb300"], [0.55, "#ffb300"], [1, "#43a047"]],
                showscale=False,
                text=text_matrix,
                texttemplate="%{text}",
                xgap=2, ygap=2,
            ))
            fig.update_layout(
                xaxis_title="題目",
                yaxis_title="學生",
                height=max(350, 26 * len(student_names_ordered)),
                margin=dict(l=100, r=20, t=40, b=60),
                font_size=11,
            )
            st.plotly_chart(fig, use_container_width=True)

            # Per-question correct rate below heatmap
            st.markdown("#### 各題全班正確率")
            rate_cols = st.columns(min(len(all_refs), 8))
            for i, ref in enumerate(all_refs):
                q = next((q for q in q_stats if q["question_ref"] == ref), None)
                if q:
                    rate = q["class_correct_rate"]
                    color = "🔴" if rate < 40 else "🟡" if rate < 60 else "🟢"
                    rate_cols[i % len(rate_cols)].metric(ref, f"{color} {rate}%")

    # ── Tab 5: Weak area diagnosis ────────────────────────────────────
    with rtabs[5]:
        weak_q = agg.get("weak_questions", [])
        strand_stats = agg.get("strand_stats", [])

        if weak_q:
            st.markdown(f"### 🔴 弱題排行榜（正確率 < 60%，共 {len(weak_q)} 題）")
            rows = [
                {
                    "排名": q["rank"],
                    " ": "🔴" if q["class_correct_rate"] < 40 else "🟡",
                    "題目": q["question_ref"],
                    "全班正確率": f"{q['class_correct_rate']}%",
                    "正確人數": f"{q['class_correct_count']} / {agg.get('valid_students', agg['total_students'])}",
                    "考核主題": q.get("topic", ""),
                    "範疇": q.get("strand", ""),
                    "常見錯誤": "；".join(q.get("common_errors", [])[:2]) or "—",
                }
                for q in weak_q
            ]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Bar chart for weak questions only
            df_wq = pd.DataFrame([
                {"題目": q["question_ref"], "正確率": q["class_correct_rate"]}
                for q in weak_q
            ])
            fig = px.bar(
                df_wq, x="題目", y="正確率",
                color="正確率", color_continuous_scale="RdYlGn",
                range_y=[0, 100], title="弱題正確率",
                labels={"正確率": "正確率 (%)"},
            )
            fig.add_hline(y=60, line_dash="dash", line_color="red")
            st.plotly_chart(fig, use_container_width=True)

        if strand_stats:
            st.markdown("### 📊 各課程範疇弱點")
            for s in strand_stats:
                icon = "🔴" if s["status"] == "弱項" else "🟡" if s["status"] == "一般" else "✅"
                rate = s["class_average_rate"]
                bar_len = int(rate / 5)
                bar_str = "█" * bar_len + "░" * (20 - bar_len)
                st.markdown(
                    f'<div class="card {"card-red" if s["status"]=="弱項" else "card-yellow" if s["status"]=="一般" else "card-green"}">'
                    f'{icon} <strong>{s["strand"]}</strong>　{rate}%　'
                    f'<code style="font-size:0.8em">{bar_str}</code>　'
                    f'<em>（涉及題目：{", ".join(s["questions"][:6])}{"…" if len(s["questions"])>6 else ""}）</em>'
                    f"</div>",
                    unsafe_allow_html=True,
                )

        if insights and not insights.get("parse_error"):
            ws_analysis = insights.get("weak_strand_analysis", [])
            if ws_analysis:
                st.markdown("### 🧠 AI 弱點深度分析")
                for ws in ws_analysis:
                    with st.expander(
                        f"🔍 {ws.get('strand','')}　（全班正確率：{ws.get('class_average_rate','')}%）"
                    ):
                        for issue in ws.get("key_issues", []):
                            st.markdown(f"- {issue}")
                        if ws.get("misconception"):
                            st.warning(f"🧩 可能的概念誤解：{ws['misconception']}")
                        if ws.get("curriculum_link"):
                            st.info(f"📚 課程連結：{ws['curriculum_link']}")

            et = insights.get("error_type_analysis", {})
            if et:
                st.markdown("### 🔎 錯誤類型分析")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**🧩 概念性誤解**")
                    st.markdown(et.get("conceptual", "—"))
                with c2:
                    st.markdown("**🔢 程序性錯誤**")
                    st.markdown(et.get("procedural", "—"))

    # ── Tab 6: Teaching recommendations ──────────────────────────────
    with rtabs[6]:
        if insights and not insights.get("parse_error"):
            recs = insights.get("teaching_recommendations", [])
            if recs:
                st.markdown("### 📅 補救教學建議")
                p_order = {"高": 0, "中": 1, "低": 2}
                recs_sorted = sorted(recs, key=lambda r: p_order.get(r.get("priority", "低"), 2))
                for rec in recs_sorted:
                    icon = "🔴" if rec.get("priority") == "高" else "🟡" if rec.get("priority") == "中" else "🟢"
                    with st.expander(
                        f"{icon} {rec.get('strand','')} — {rec.get('strategy','')}"
                    ):
                        st.markdown(f"**優先級：** {rec.get('priority','')}　**建議時間：** {rec.get('timeline','')}")
                        acts = rec.get("activities", [])
                        if acts:
                            st.markdown("**教學活動：**")
                            for a in acts:
                                st.markdown(f"- {a}")

            if insights.get("attention_students_note"):
                st.markdown("### 👀 需要個別關注的學生")
                st.warning(insights["attention_students_note"])

            if insights.get("positive_findings"):
                st.markdown("### 💪 全班亮點")
                st.success(insights["positive_findings"])
        else:
            st.info("教學建議需要先完成AI分析才能顯示。")

    # ── Tab 7: Practice Question Generator ─────────────────────────
    with rtabs[7]:
        st.markdown("### 📝 弱點針對練習 — 因材施教，生成針對性練習題")
        st.caption("根據每位學生的答錯題目和弱點，由 AI 自動生成相似題型的練習題；全對的同學亦會獲得鞏固延伸練習。")

        student_results = agg.get("student_results", [])
        students_with_errors = []
        students_perfect = []
        for s in student_results:
            if s.get("parse_error"):
                continue
            wrong = [q for q in s.get("question_results", []) if q.get("is_correct") is False]
            if wrong:
                students_with_errors.append((s, wrong))
            else:
                all_qs = s.get("question_results", [])
                if all_qs:
                    students_perfect.append((s, all_qs))

        if not students_with_errors and not students_perfect:
            st.info("沒有可用的學生數據。")
        else:
            # ── Common settings ───────────────────────────────────────
            cfg_c1, cfg_c2 = st.columns(2)
            with cfg_c1:
                num_q = st.number_input("每人題目數量", min_value=1, max_value=15, value=5, key="pq_num")
            with cfg_c2:
                difficulty = st.selectbox("難度", ["簡單", "適中", "進階"], index=1, key="pq_diff")

            # ── Tabs: individual vs batch ─────────────────────────────
            mode_tab1, mode_tab2 = st.tabs(["👤 個別生成", "📋 批量生成全班"])

            # ── Individual mode ───────────────────────────────────────
            with mode_tab1:
                student_options = {
                    f"{s.get('student_name', '未知')}　（答錯 {len(w)} 題，得分率 {s.get('percentage', 0):.0f}%）": i
                    for i, (s, w) in enumerate(students_with_errors)
                }
                selected_label = st.selectbox(
                    "選擇學生", options=list(student_options.keys()), key="pq_student_sel"
                )

                sel_idx = student_options[selected_label]
                sel_student, sel_wrong = students_with_errors[sel_idx]
                sel_name = sel_student.get("student_name", "未知")

                with st.expander(f"📋 {sel_name} 的答錯題目（共 {len(sel_wrong)} 題）", expanded=False):
                    wrong_rows = []
                    for q in sel_wrong:
                        wrong_rows.append({
                            "題目": q.get("question_ref", ""),
                            "主題": q.get("topic", ""),
                            "範疇": q.get("strand", ""),
                            "學生答案": q.get("student_answer", "—"),
                            "正確答案": q.get("correct_answer", "—"),
                            "錯誤類型": q.get("error_type", ""),
                        })
                    st.dataframe(pd.DataFrame(wrong_rows), use_container_width=True, hide_index=True)

                pq_state_key = f"pq_result_{sel_name}"

                if st.button(
                    f"🤖 為 {sel_name} 生成 {num_q} 道針對練習題",
                    use_container_width=True,
                    key="pq_gen_btn",
                ):
                    with st.spinner(f"AI 正在為 {sel_name} 設計練習題…"):
                        try:
                            pq_analyzer = MathAnalyzer(api_key)
                            pq_result = pq_analyzer.generate_practice_questions(
                                student_name=sel_name,
                                grade=s_grade,
                                weak_questions=sel_wrong,
                                num_questions=num_q,
                                difficulty=difficulty,
                            )
                            st.session_state[pq_state_key] = pq_result
                        except Exception as e:
                            st.error(f"❌ 生成失敗：{e}")
                            st.exception(e)

                pq_result = st.session_state.get(pq_state_key)
                if pq_result:
                    st.markdown("---")
                    ws = pq_result.get("weakness_summary", "")
                    if ws:
                        st.info(f"**弱點概述：** {ws}")

                    questions = pq_result.get("practice_questions", [])
                    for q in questions:
                        qn = q.get("question_number", "")
                        qtype = q.get("question_type", "")
                        target = q.get("targeted_weakness", "")
                        with st.expander(
                            f"第 {qn} 題（{qtype}）— 針對：{target}", expanded=True
                        ):
                            st.markdown(
                                f"**範疇：** {q.get('strand', '')}　|　**主題：** {q.get('topic', '')}"
                            )
                            st.markdown("---")
                            st.markdown(f"**📖 題目：**\n\n{q.get('question_text', '')}")
                            hint = q.get("hints", "")
                            if hint:
                                st.caption(f"💡 提示：{hint}")
                            with st.expander("🔑 查看答案及解題步驟", expanded=False):
                                st.markdown(f"**答案：** {q.get('answer', '')}")
                                steps = q.get("solution_steps", [])
                                if steps:
                                    st.markdown("**解題步驟：**")
                                    for si, step in enumerate(steps, 1):
                                        st.markdown(f"{si}. {step}")
                                expl = q.get("explanation", "")
                                if expl:
                                    st.caption(f"📌 設計理由：{expl}")

                    tips = pq_result.get("study_tips", [])
                    if tips:
                        st.markdown("### 📚 學習建議")
                        for tip in tips:
                            st.markdown(f"- {tip}")

                    # ── HTML download buttons ─────────────────────────
                    st.markdown("---")
                    st.markdown("#### 🖨️ 下載練習題（可直接列印）")
                    dl_c1, dl_c2 = st.columns(2)
                    with dl_c1:
                        html_student = build_practice_worksheets_html(
                            [pq_result], grade=s_grade, show_answers=False
                        )
                        st.download_button(
                            f"📄 學生版（A4，無答案）",
                            data=html_student.encode("utf-8"),
                            file_name=f"練習題_{sel_name}_{s_grade}_學生版.html",
                            mime="text/html",
                            key="pq_dl_student",
                            use_container_width=True,
                        )
                    with dl_c2:
                        html_teacher = build_practice_worksheets_html(
                            [pq_result], grade=s_grade, show_answers=True
                        )
                        st.download_button(
                            f"📋 老師版（含答案及解題步驟）",
                            data=html_teacher.encode("utf-8"),
                            file_name=f"練習題_{sel_name}_{s_grade}_老師版.html",
                            mime="text/html",
                            key="pq_dl_teacher",
                            use_container_width=True,
                        )
                    st.caption("💡 下載後在瀏覽器開啟，按 Ctrl+P（Mac: Cmd+P）或點頁面上的「🖨️ 列印」鍵即可列印。")

            # ── Batch mode ────────────────────────────────────────────
            with mode_tab2:
                total_err = len(students_with_errors)
                total_perf = len(students_perfect)
                total_all = total_err + total_perf

                st.info(
                    f"全班共 **{total_all}** 位學生（**{total_err}** 位有答錯、**{total_perf}** 位全對）。\n\n"
                    f"按下按鈕後，AI 會一次過為全班生成練習題：答錯的同學生成**弱點針對練習**，"
                    f"全對的同學生成**鞏固延伸練習**，完成後可一鍵下載全班練習題 HTML，直接送印。"
                )

                # Show combined student list
                with st.expander(f"📋 全班學生名單（{total_all} 人）", expanded=False):
                    name_rows = []
                    for s, w in students_with_errors:
                        name_rows.append({
                            "學生": s.get("student_name", "未知"),
                            "類型": f"⚠️ 答錯 {len(w)} 題",
                            "得分率": f"{s.get('percentage', 0):.0f}%",
                            "練習類型": "弱點針對練習",
                        })
                    for s, _ in students_perfect:
                        name_rows.append({
                            "學生": s.get("student_name", "未知"),
                            "類型": "✅ 全對",
                            "得分率": f"{s.get('percentage', 0):.0f}%",
                            "練習類型": "鞏固延伸練習",
                        })
                    st.dataframe(pd.DataFrame(name_rows), use_container_width=True, hide_index=True)

                batch_key = "pq_batch_results"
                batch_running_key = "pq_batch_done"

                if st.button(
                    f"🚀 一鍵為全班 {total_all} 位學生生成練習題（每人 {num_q} 題）",
                    use_container_width=True,
                    type="primary",
                    key="pq_batch_btn",
                ):
                    batch_results = []
                    progress_bar = st.progress(0, text="準備中…")
                    # Build a live progress table
                    progress_rows = []
                    for s, w in students_with_errors:
                        progress_rows.append({
                            "學生": s.get("student_name", "未知"),
                            "類型": f"弱點練習（答錯 {len(w)} 題）",
                            "狀態": "⏳ 等待中",
                        })
                    for s, _ in students_perfect:
                        progress_rows.append({
                            "學生": s.get("student_name", "未知"),
                            "類型": "鞏固延伸練習（全對）",
                            "狀態": "⏳ 等待中",
                        })
                    progress_table = st.empty()
                    progress_table.dataframe(
                        pd.DataFrame(progress_rows),
                        use_container_width=True,
                        hide_index=True,
                    )
                    pq_analyzer = MathAnalyzer(api_key)
                    errors_log = []

                    # ── Phase 1: error students (weakness practice) ───
                    for idx, (s_data, s_wrong) in enumerate(students_with_errors):
                        sname = s_data.get("student_name", f"學生{idx+1}")
                        progress_rows[idx]["狀態"] = "🔄 生成中…"
                        progress_table.dataframe(
                            pd.DataFrame(progress_rows),
                            use_container_width=True,
                            hide_index=True,
                        )
                        pct = (idx + 1) / total_all
                        progress_bar.progress(pct, text=f"正在處理：{sname}（{idx+1}/{total_all}）")
                        try:
                            result = pq_analyzer.generate_practice_questions(
                                student_name=sname,
                                grade=s_grade,
                                weak_questions=s_wrong,
                                num_questions=num_q,
                                difficulty=difficulty,
                            )
                            if result.get("parse_error"):
                                result = pq_analyzer.generate_practice_questions(
                                    student_name=sname,
                                    grade=s_grade,
                                    weak_questions=s_wrong,
                                    num_questions=num_q,
                                    difficulty=difficulty,
                                )
                            if result.get("parse_error"):
                                errors_log.append(f"{sname}: AI 回覆格式錯誤（已重試一次）")
                                progress_rows[idx]["狀態"] = "❌ 失敗"
                            else:
                                n_qs = len(result.get("practice_questions", []))
                                progress_rows[idx]["狀態"] = f"✅ 完成（{n_qs} 題）"
                            result["_gen_type"] = "weakness"
                            batch_results.append(result)
                        except Exception as e:
                            errors_log.append(f"{sname}: {e}")
                            batch_results.append({
                                "student_name": sname,
                                "parse_error": True,
                                "error": str(e),
                                "_gen_type": "weakness",
                            })
                            progress_rows[idx]["狀態"] = "❌ 失敗"
                        progress_table.dataframe(
                            pd.DataFrame(progress_rows),
                            use_container_width=True,
                            hide_index=True,
                        )

                    # ── Phase 2: perfect students (consolidation) ─────
                    for idx2, (s_data, s_qs) in enumerate(students_perfect):
                        row_idx = total_err + idx2
                        sname = s_data.get("student_name", f"學生{row_idx+1}")
                        overall_idx = total_err + idx2 + 1
                        progress_rows[row_idx]["狀態"] = "🔄 生成中…"
                        progress_table.dataframe(
                            pd.DataFrame(progress_rows),
                            use_container_width=True,
                            hide_index=True,
                        )
                        pct = overall_idx / total_all
                        progress_bar.progress(pct, text=f"正在處理：{sname}（{overall_idx}/{total_all}）")
                        try:
                            result = pq_analyzer.generate_consolidation_questions(
                                student_name=sname,
                                grade=s_grade,
                                all_questions=s_qs,
                                num_questions=num_q,
                                difficulty="進階",
                            )
                            if result.get("parse_error"):
                                result = pq_analyzer.generate_consolidation_questions(
                                    student_name=sname,
                                    grade=s_grade,
                                    all_questions=s_qs,
                                    num_questions=num_q,
                                    difficulty="進階",
                                )
                            if result.get("parse_error"):
                                errors_log.append(f"{sname}: AI 回覆格式錯誤（已重試一次）")
                                progress_rows[row_idx]["狀態"] = "❌ 失敗"
                            else:
                                n_qs = len(result.get("practice_questions", []))
                                progress_rows[row_idx]["狀態"] = f"✅ 完成（{n_qs} 題）"
                            result["_gen_type"] = "consolidation"
                            batch_results.append(result)
                        except Exception as e:
                            errors_log.append(f"{sname}: {e}")
                            batch_results.append({
                                "student_name": sname,
                                "parse_error": True,
                                "error": str(e),
                                "_gen_type": "consolidation",
                            })
                            progress_rows[row_idx]["狀態"] = "❌ 失敗"
                        progress_table.dataframe(
                            pd.DataFrame(progress_rows),
                            use_container_width=True,
                            hide_index=True,
                        )

                    st.session_state[batch_key] = batch_results
                    st.session_state[batch_running_key] = True
                    n_ok = len([r for r in batch_results if not r.get("parse_error")])
                    n_fail = total_all - n_ok
                    if n_fail:
                        progress_bar.progress(1.0, text=f"⚠️ 完成！成功 {n_ok} 份，失敗 {n_fail} 份")
                    else:
                        progress_bar.progress(
                            1.0,
                            text=f"✅ 全部完成！{total_err} 份弱點練習 + {total_perf} 份鞏固練習",
                        )
                    if errors_log:
                        for err in errors_log:
                            st.warning(f"⚠️ {err}")

                # Show download buttons once batch is done
                batch_results = st.session_state.get(batch_key, [])
                if batch_results:
                    ok_results = [r for r in batch_results if not r.get("parse_error")]
                    n_weakness = len([r for r in ok_results if r.get("_gen_type") == "weakness"])
                    n_consol = len([r for r in ok_results if r.get("_gen_type") == "consolidation"])
                    st.markdown(
                        f"#### 🖨️ 下載全班練習題（{len(ok_results)} 位學生："
                        f"{n_weakness} 份弱點練習 + {n_consol} 份鞏固練習）"
                    )
                    st.caption(
                        "每位學生獨立一頁 A4。瀏覽器開啟後點「🖨️ 列印全部」即可一次列印全班。"
                    )

                    dl_b1, dl_b2 = st.columns(2)
                    with dl_b1:
                        html_all_student = build_practice_worksheets_html(
                            ok_results, grade=s_grade, show_answers=False
                        )
                        st.download_button(
                            f"📄 全班學生版（{len(ok_results)} 頁，無答案）",
                            data=html_all_student.encode("utf-8"),
                            file_name=f"全班練習題_{s_grade}_學生版.html",
                            mime="text/html",
                            key="pq_batch_dl_student",
                            use_container_width=True,
                            type="primary",
                        )
                    with dl_b2:
                        html_all_teacher = build_practice_worksheets_html(
                            ok_results, grade=s_grade, show_answers=True
                        )
                        st.download_button(
                            f"📋 全班老師版（{len(ok_results)} 頁，含答案）",
                            data=html_all_teacher.encode("utf-8"),
                            file_name=f"全班練習題_{s_grade}_老師版.html",
                            mime="text/html",
                            key="pq_batch_dl_teacher",
                            use_container_width=True,
                        )

                    # Per-student preview
                    st.markdown("---")
                    st.markdown("#### 📋 各學生練習題預覽")
                    for r in ok_results:
                        sname = r.get("student_name", "學生")
                        qs = r.get("practice_questions", [])
                        ws = r.get("weakness_summary", "")
                        gen_type = r.get("_gen_type", "weakness")
                        type_label = "🌟 鞏固" if gen_type == "consolidation" else "⚠️ 弱點"
                        with st.expander(f"{type_label} {sname}（{len(qs)} 題）", expanded=False):
                            if ws:
                                st.info(f"{'練習方向' if gen_type == 'consolidation' else '弱點概述'}：{ws}")
                            for q in qs:
                                st.markdown(
                                    f"**第 {q.get('question_number')} 題** "
                                    f"（{q.get('question_type', '')}）— "
                                    f"{q.get('question_text', '')[:80]}…"
                                )

    # ── Tab 8: Export ─────────────────────────────────────────────────
    with rtabs[8]:
        st.markdown("### 📥 匯出分析報告")

        st.markdown("#### 🌐 HTML 互動報告（完整還原分析介面）")
        st.caption("匯出為 HTML 網頁檔案，包含所有互動圖表，在瀏覽器中開啟即可查看，效果與本頁分析介面一致。")
        if st.button("🔄 生成 HTML 報告", use_container_width=True, key="s_html_btn"):
            with st.spinner("正在生成 HTML 報告…"):
                try:
                    html_str = build_student_html_report(agg, insights, s_grade, s_label)
                    st.session_state["s_html_bytes"] = html_str.encode("utf-8")
                    st.session_state["s_html_stem"] = f"{s_grade}_{s_label or 'class'}"
                except Exception as e:
                    st.error(f"❌ HTML 生成失敗：{e}")
                    st.exception(e)

        if st.session_state.get("s_html_bytes"):
            html_stem = st.session_state.get("s_html_stem", "report")
            st.download_button(
                "⬇️ 點擊下載 HTML 報告",
                data=st.session_state["s_html_bytes"],
                file_name=f"student_report_{html_stem}.html",
                mime="text/html",
                use_container_width=True,
                key="s_html_dl",
            )
            st.success(f"✅ HTML 報告已生成 ({len(st.session_state['s_html_bytes']):,} bytes)　— 下載後用瀏覽器開啟即可查看")

        st.markdown("---")
        st.markdown("#### 📂 JSON 原始資料")
        export_data = {"aggregated": agg, "insights": insights, "grade": s_grade, "notes": s_label}
        st.download_button(
            "📥 下載 JSON 資料",
            data=json.dumps(export_data, ensure_ascii=False, indent=2).encode("utf-8-sig"),
            file_name=f"student_analysis_{s_grade}_{s_label or 'class'}.json",
            mime="application/json",
        )

        with st.expander("📋 展開查看原始分析數據"):
            st.json(export_data)

