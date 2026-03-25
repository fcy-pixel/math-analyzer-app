"""
小學數學學生表現分析系統
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
    page_title="小學數學學生表現分析系統",
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
  <h1>📊 小學數學學生表現分析系統</h1>
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
**AQP 報告：** Excel、CSV、PDF
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
# Mode selector
# ---------------------------------------------------------------------------
mode = st.radio(
    "選擇分析模式",
    ["📝 學生試卷批量分析（新）", "📊 AQP + 答案版分析"],
    horizontal=True,
    label_visibility="collapsed",
)

st.markdown("---")

# ===========================================================================
# MODE 1 — Student paper batch analysis
# ===========================================================================
if mode == "📝 學生試卷批量分析（新）":

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
        st.caption("根據每位學生的答錯題目和弱點，由 AI 自動生成相似題型的練習題，可一鍵列印全班練習卷（每人獨立一頁 A4）。")

        student_results = agg.get("student_results", [])
        students_with_errors = []
        for s in student_results:
            if s.get("parse_error"):
                continue
            wrong = [q for q in s.get("question_results", []) if q.get("is_correct") is False]
            if wrong:
                students_with_errors.append((s, wrong))

        if not students_with_errors:
            st.success("🎉 所有學生都全對！無需生成弱點練習題。")
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
                st.info(
                    f"共 **{total_err}** 位學生有答錯題目。按下按鈕後，AI 會依次為每位學生生成練習題，"
                    f"完成後可一鍵下載全班練習題 HTML，直接送印。"
                )

                # Show student list
                with st.expander(f"📋 需要生成練習題的學生名單（{total_err} 人）", expanded=False):
                    name_rows = [
                        {
                            "學生": s.get("student_name", "未知"),
                            "答錯題數": len(w),
                            "得分率": f"{s.get('percentage', 0):.0f}%",
                        }
                        for s, w in students_with_errors
                    ]
                    st.dataframe(pd.DataFrame(name_rows), use_container_width=True, hide_index=True)

                batch_key = "pq_batch_results"
                batch_running_key = "pq_batch_done"

                if st.button(
                    f"🚀 為全部 {total_err} 位學生生成練習題（每人 {num_q} 題）",
                    use_container_width=True,
                    type="primary",
                    key="pq_batch_btn",
                ):
                    batch_results = []
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    pq_analyzer = MathAnalyzer(api_key)
                    errors_log = []

                    for idx, (s_data, s_wrong) in enumerate(students_with_errors):
                        sname = s_data.get("student_name", f"學生{idx+1}")
                        status_text.info(
                            f"⏳ 正在為 **{sname}** 生成練習題… （{idx+1} / {total_err}）"
                        )
                        try:
                            result = pq_analyzer.generate_practice_questions(
                                student_name=sname,
                                grade=s_grade,
                                weak_questions=s_wrong,
                                num_questions=num_q,
                                difficulty=difficulty,
                            )
                            batch_results.append(result)
                        except Exception as e:
                            errors_log.append(f"{sname}: {e}")
                            batch_results.append({
                                "student_name": sname,
                                "parse_error": True,
                                "error": str(e),
                            })
                        progress_bar.progress((idx + 1) / total_err)

                    st.session_state[batch_key] = batch_results
                    st.session_state[batch_running_key] = True
                    status_text.success(
                        f"✅ 已完成全部 {total_err} 位學生的練習題生成！"
                        + (f"（{len(errors_log)} 位生成失敗）" if errors_log else "")
                    )
                    if errors_log:
                        for err in errors_log:
                            st.warning(f"⚠️ {err}")

                # Show download buttons once batch is done
                batch_results = st.session_state.get(batch_key, [])
                if batch_results:
                    ok_results = [r for r in batch_results if not r.get("parse_error")]
                    st.markdown(f"#### 🖨️ 下載全班練習題（{len(ok_results)} 位學生）")
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
                        with st.expander(f"📄 {sname}（{len(qs)} 題）", expanded=False):
                            if ws:
                                st.info(f"弱點概述：{ws}")
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

    st.stop()  # end of MODE 1 — prevents MODE 2 from running


# ===========================================================================
# MODE 2 — AQP + Answer key (original functionality)
# Reached only when mode == "📊 AQP + 答案版分析"
# ===========================================================================

# ---------------------------------------------------------------------------
# File upload section
# ---------------------------------------------------------------------------
st.markdown("### 📁 上傳文件")
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("#### 📋 AQP 成績報告（全年級）")
    aqp_file = st.file_uploader(
        "支援 Excel / CSV / PDF",
        type=["xlsx", "xls", "csv", "pdf"],
        key="aqp",
        help="上傳全年級 AQP 報告，反映整個年級所有班別的共同表現",
    )
    if aqp_file:
        st.success(f"✅ 已選擇：{aqp_file.name}")

with col_right:
    st.markdown("#### 📄 試卷答案版（可選）")
    exam_file = st.file_uploader(
        "支援 PDF / JPG / PNG",
        type=["pdf", "jpg", "jpeg", "png"],
        key="exam",
        help="上傳試卷的答案版（參考答案卷），AI 將分析每題考核概念及預測學生出錯原因",
    )
    if exam_file:
        st.success(f"✅ 已選擇：{exam_file.name}")

# ---------------------------------------------------------------------------
# Class info
# ---------------------------------------------------------------------------
st.markdown("### 📚 年級資料")
c1, c2 = st.columns(2)
with c1:
    grade = st.selectbox("年級", ["P1", "P2", "P3", "P4", "P5", "P6"], index=3)
with c2:
    class_label = st.text_input("備註（可選）", placeholder="例：2024-25 上學期")

st.markdown("---")

# ---------------------------------------------------------------------------
# Analyse button
# ---------------------------------------------------------------------------
_, btn_col, _ = st.columns([1, 2, 1])
with btn_col:
    analyse_btn = st.button(
        "🔍 開始深度分析",
        type="primary",
        use_container_width=True,
        disabled=(not aqp_file and not exam_file),
    )

# ---------------------------------------------------------------------------
# Run analysis
# ---------------------------------------------------------------------------
if analyse_btn:
    results: dict = {}
    progress = st.progress(0)
    status = st.empty()

    try:
        analyzer = MathAnalyzer(api_key)
        processor = FileProcessor()

        # --- AQP ---
        if aqp_file:
            status.info("📋 正在處理 AQP 報告（全年級整體報告）…")
            progress.progress(10)
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=os.path.splitext(aqp_file.name)[1]
            ) as tmp:
                tmp.write(aqp_file.read())
                tmp_path = tmp.name
            try:
                aqp_data = processor.process_aqp(tmp_path)
            finally:
                os.unlink(tmp_path)

            page_count = aqp_data.get("page_count", "?")  
            status.info(f"🤖 AI 正在逐頁分析 AQP 報告（共 {page_count} 頁，使用 qwen-vl-max 視覺分析）…")
            progress.progress(30)
            results["aqp"] = analyzer.analyze_aqp(aqp_data, grade)

        # --- Exam ---
        if exam_file:
            status.info("📄 正在處理試卷…")
            progress.progress(45)
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=os.path.splitext(exam_file.name)[1]
            ) as tmp:
                tmp.write(exam_file.read())
                tmp_path = tmp.name
            try:
                exam_data = processor.process_exam(tmp_path)
            finally:
                os.unlink(tmp_path)

            page_count = exam_data.get("page_count", "?")
            status.info(f"🤖 AI 正在逐頁分析試卷答案版（共 {page_count} 頁，每批 6 頁，分析每題考核概念）…")
            progress.progress(65)
            results["exam"] = analyzer.analyze_exam(exam_data, grade)

        # --- Combined ---
        if "aqp" in results and "exam" in results:
            status.info("🔗 正在交叉比對全年級 AQP 弱點與試卷答案版，推論全年級學生出錯根因…")
            progress.progress(85)
            results["combined"] = analyzer.combined_analysis(
                results["aqp"], results["exam"], grade
            )

        progress.progress(100)
        status.success("✅ 分析完成！")

        st.session_state["results"] = results
        st.session_state["grade"] = grade
        st.session_state["class_label"] = class_label
        st.session_state.pop("aqp_pdf_bytes", None)
        st.session_state.pop("aqp_pdf_stem", None)

    except Exception as exc:
        progress.empty()
        status.empty()
        st.error(f"❌ 分析時發生錯誤：{exc}")
        st.exception(exc)

# ---------------------------------------------------------------------------
# Display results
# ---------------------------------------------------------------------------
if "results" not in st.session_state:
    st.stop()

results = st.session_state["results"]
grade = st.session_state.get("grade", grade)
class_label = st.session_state.get("class_label", "")
class_label_str = f"（{class_label}）" if class_label else ""

st.markdown(f"---\n## 📊 {grade}{class_label_str} 全級數學分析報告")

# Build tab list dynamically
tab_labels = ["📋 整體概覽"]
if "aqp" in results:
    tab_labels.append("📊 AQP 報告")
if "exam" in results:
    tab_labels.append("📄 試卷分析")
if "combined" in results:
    tab_labels.append("🎯 綜合分析")
tab_labels.append("💡 改善建議")
tab_labels.append("📥 匯出報告")

tabs = st.tabs(tab_labels)
tab_idx = 0

# ============================================================
# TAB: 整體概覽
# ============================================================
with tabs[tab_idx]:
    tab_idx += 1

    primary = results.get("aqp") or results.get("exam", {})
    if primary.get("parse_error"):
        st.warning("⚠️ 無法解析結構化數據，顯示原始回應：")
        st.text(primary.get("raw_response", ""))
        st.stop()

    # Metric cards
    overall = primary.get("overall_performance", {})
    if overall:
        m1, m2, m3, m4 = st.columns(4)
        level = overall.get("performance_level", "—")
        level_icon = {"優秀": "🟢", "良好": "🔵", "一般": "🟡", "需要改善": "🔴"}.get(
            level, "⚪"
        )
        m1.metric("整體表現", f"{level_icon} {level}")

        score = overall.get("class_average_percentage") or overall.get("score_percentage")
        if score is not None:
            m2.metric("全年級平均", f"{score}%")
        else:
            m2.metric("強項範疇", len(primary.get("class_strong_areas", primary.get("strong_areas", []))))

        weak_list = primary.get("class_weak_areas", primary.get("weak_areas", []))
        m3.metric("全年級弱點數", len(weak_list))
        m4.metric("覆蓋頁數", primary.get("page_count", "—"))

        st.markdown(f"> {overall.get('summary', '')}")

    # Radar chart
    strand_data = primary.get("strand_analysis", [])
    if strand_data:
        st.markdown("### 📈 各範疇全年級表現雷達圖")
        cats = [s.get("strand", "") for s in strand_data]
        vals = []
        for s in strand_data:
            raw_score = s.get("class_score") or s.get("score")
            status = s.get("status", "")
            if raw_score is not None:
                try:
                    vals.append(min(100.0, float(str(raw_score).replace("%", ""))))
                    continue
                except (ValueError, TypeError):
                    pass
            vals.append({"強項": 85, "一般": 58, "弱項": 30}.get(status, 50))

        if cats and vals:
            fig = go.Figure(
                go.Scatterpolar(
                    r=vals + [vals[0]],
                    theta=cats + [cats[0]],
                    fill="toself",
                    fillcolor="rgba(102,126,234,0.25)",
                    line=dict(color="rgb(102,126,234)", width=2),
                    name="全年級表現",
                )
            )
            fig.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                showlegend=False,
                height=420,
                margin=dict(l=60, r=60, t=40, b=40),
            )
            st.plotly_chart(fig, use_container_width=True)

    # Class weak areas
    weak_areas = primary.get("class_weak_areas", primary.get("weak_areas", []))
    if weak_areas:
        st.markdown("### ⚠️ 全年級主要弱點摘要")
        sev_icon = {"嚴重": "🔴", "中等": "🟡", "輕微": "🟢"}
        for area in weak_areas:
            icon = sev_icon.get(area.get("severity", ""), "⚪")
            desc = area.get("description") or area.get("likely_misconception", "")
            st.markdown(
                f'<div class="card card-red">'
                f'{icon} <strong>{area.get("topic","")}</strong>'
                f' <em>({area.get("strand","")})</em>'
                f"<br><small>{desc}</small></div>",
                unsafe_allow_html=True,
            )


# ============================================================
# TAB: AQP 報告
# ============================================================
if "aqp" in results:
    with tabs[tab_idx]:
        tab_idx += 1
        aqp = results["aqp"]

        if aqp.get("parse_error"):
            st.text(aqp.get("raw_response", ""))
        else:
            st.caption("📌 此為全年級整體 AQP 報告分析，反映整個年級所有班別學生的共同表現")

            # Strand table + bar chart
            strand_data = aqp.get("strand_analysis", [])
            if strand_data:
                st.markdown("### 📊 全年級各範疇表現")
                sicon = {"強項": "✅", "一般": "➡️", "弱項": "⚠️"}
                rows = [
                    {
                        "範疇": s.get("strand", ""),
                        "狀態": f"{sicon.get(s.get('status',''), '')} {s.get('status','')}",
                        "全年級得分率": s.get("class_score", "—"),
                        "表現描述": s.get("performance", ""),
                        "普遍困難主題": "、".join(s.get("specific_topics_struggled", [])),
                    }
                    for s in strand_data
                ]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                # Bar chart (numeric scores only)
                numeric_rows = []
                for s in strand_data:
                    try:
                        val = float(str(s.get("class_score", "")).replace("%", ""))
                        numeric_rows.append({"範疇": s.get("strand", ""), "全年級得分率": val})
                    except (ValueError, TypeError):
                        pass
                if numeric_rows:
                    df_bar = pd.DataFrame(numeric_rows)
                    fig = px.bar(
                        df_bar, x="範疇", y="全年級得分率",
                        color="全年級得分率", color_continuous_scale="RdYlGn", range_y=[0, 100],
                    )
                    fig.update_layout(coloraxis_showscale=False)
                    st.plotly_chart(fig, use_container_width=True)

            # Weak question ranking table (sorted by correct rate ascending)
            weak_questions = aqp.get("weak_questions", [])
            qperf = aqp.get("question_performance", [])

            # Build weak ranking from dedicated field, or fallback to question_performance
            ranking_source = weak_questions if weak_questions else [
                q for q in qperf
                if q.get("class_correct_rate") is not None
                and isinstance(q.get("class_correct_rate"), (int, float))
                and float(q["class_correct_rate"]) < 60
            ]
            if ranking_source:
                st.markdown("### 🔴 弱題排行榜（正確率由低至高）")
                ranking_rows = []
                for q in ranking_source:
                    rate = q.get("correct_rate") or q.get("class_correct_rate")
                    try:
                        rate_val = float(str(rate).replace("%", ""))
                        rate_str = f"{rate_val:.0f}%"
                        rate_bar = "🔴" if rate_val < 40 else "🟡" if rate_val < 60 else "🟢"
                    except (TypeError, ValueError):
                        rate_str = str(rate) if rate is not None else "—"
                        rate_bar = "⚪"
                    ranking_rows.append({
                        "排名": q.get("rank", ""),
                        "難度 / 正確率": f"{rate_bar} {rate_str}",
                        "題目": q.get("question_ref", ""),
                        "考核主題": q.get("topic", ""),
                        "範疇": q.get("strand", ""),
                        "常見錯誤": q.get("common_error") or q.get("common_errors", ""),
                    })
                try:
                    ranking_rows.sort(
                        key=lambda x: float(
                            x["難度 / 正確率"].replace("🔴", "").replace("🟡", "").replace("🟢", "").replace("⚪", "").strip().replace("%", "")
                        )
                    )
                except (ValueError, TypeError):
                    pass
                for i, row in enumerate(ranking_rows, 1):
                    row["排名"] = i
                st.dataframe(pd.DataFrame(ranking_rows), use_container_width=True, hide_index=True)
                st.caption(f"共找出 {len(ranking_rows)} 條弱題（全年級正確率低於 60%）")

            # Full per-question class performance
            if qperf:
                st.markdown("### 📋 各題全年級正確率（完整列表）")
                rows = [
                    {
                        "題目": q.get("question_ref", ""),
                        "考核主題": q.get("topic", ""),
                        "範疇": q.get("strand", ""),
                        "全年級正確率": f"{q.get('class_correct_rate', '—')}%" if q.get('class_correct_rate') is not None else "—",
                        "難度": q.get("difficulty", ""),
                        "常見錯誤": q.get("common_errors", ""),
                    }
                    for q in qperf
                ]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Class weak areas with misconceptions
            class_weak = aqp.get("class_weak_areas", aqp.get("weak_areas", []))
            if class_weak:
                st.markdown("### ⚠️ 全年級弱點及概念誤解")
                sev_icon = {"嚴重": "🔴", "中等": "🟡", "輕微": "🟢"}
                for a in class_weak:
                    icon = sev_icon.get(a.get("severity", ""), "⚪")
                    with st.expander(f"{icon} {a.get('topic','')} （{a.get('strand','')}）"):
                        st.markdown(f"**弱點描述：** {a.get('description','')}")
                        if a.get("likely_misconception"):
                            st.warning(f"🧠 可能的概念誤解：{a['likely_misconception']}")
                        if a.get("affected_question_types"):
                            st.markdown("**涉及題型：** " + "、".join(a["affected_question_types"]))

            # Teaching implications
            implications = aqp.get("teaching_implications", [])
            if implications:
                st.markdown("### 🏫 教學啟示")
                for impl in implications:
                    with st.expander(f"📌 {impl.get('issue', '')}"):
                        st.markdown(f"**佐證：** {impl.get('evidence', '')}")
                        st.info(f"💡 建議教學策略：{impl.get('suggested_teaching_strategy', '')}")

            # Strong areas
            strong_areas = aqp.get("class_strong_areas", aqp.get("strong_areas", []))
            if strong_areas:
                st.markdown("### 💪 全年級強項")
                for a in strong_areas:
                    st.markdown(
                        f'<div class="card card-green">✅ <strong>{a.get("topic","")}</strong>'
                        f' ({a.get("strand","")})<br><small>{a.get("description","")}</small></div>',
                        unsafe_allow_html=True,
                    )


# ============================================================
# TAB: 試卷分析
# ============================================================
if "exam" in results:
    with tabs[tab_idx]:
        tab_idx += 1
        exam = results["exam"]

        if exam.get("parse_error"):
            st.text(exam.get("raw_response", ""))
        else:
            overview = exam.get("exam_overview", {})
            if overview:
                m1, m2, m3 = st.columns(3)
                m1.metric("試卷難度", overview.get("estimated_difficulty", "—"))
                m2.metric("涵蓋主題數", len(overview.get("topics_covered", [])))
                m3.metric("涉及範疇數", len(overview.get("strands_tested", [])))
                covered = overview.get("topics_covered", [])
                if covered:
                    st.markdown("**涵蓋主題：** " + " · ".join(f"`{t}`" for t in covered))

            # Question-level table
            questions = exam.get("question_analysis", [])
            if questions:
                st.markdown(f"### 📝 試卷逐題分析（共 {len(questions)} 題）")
                picon = {"正確": "✅", "錯誤": "❌", "部分正確": "⚠️", "未能判斷": "❓"}
                rows = [
                    {
                        "頁": q.get("page", ""),
                        "題目": q.get("question_ref", ""),
                        "考核主題": q.get("topic", ""),
                        "範疇": q.get("strand", ""),
                        "難度": q.get("difficulty", ""),
                        "表現": f"{picon.get(q.get('correctness',''), '❓')} {q.get('correctness','')}",
                        "觀察到的錯誤": q.get("error_observed") or "—",
                    }
                    for q in questions
                ]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                # Performance pie chart
                perf_counts = {}
                for q in questions:
                    k = q.get("correctness", "未能判斷")
                    perf_counts[k] = perf_counts.get(k, 0) + 1
                df_pie = pd.DataFrame(
                    [{"表現": k, "題數": v} for k, v in perf_counts.items()]
                )
                if not df_pie.empty:
                    fig = px.pie(
                        df_pie, names="表現", values="題數",
                        color="表現",
                        color_discrete_map={"正確": "#43a047", "錯誤": "#e53935",
                                            "部分正確": "#f9a825", "未能判斷": "#90a4ae"},
                        title="答題表現分佈",
                    )
                    st.plotly_chart(fig, use_container_width=True)

            # Error patterns
            patterns = exam.get("error_patterns", [])
            if patterns:
                st.markdown("### 🔍 錯誤模式分析")
                for p in patterns:
                    freq = p.get("frequency", "")
                    icon = "🔴" if freq == "頻繁" else "🟡"
                    affected = "、".join(p.get("affected_questions", []))
                    st.markdown(
                        f'<div class="card card-yellow">{icon} <strong>{p.get("pattern","")}</strong>'
                        f'<br><small>相關概念：{p.get("related_concept","")}'
                        f'{ "　涉及題目：" + affected if affected else ""}</small></div>',
                        unsafe_allow_html=True,
                    )

            # Weak areas
            weak = exam.get("weak_areas", [])
            if weak:
                st.markdown("### ⚠️ 弱點")
                sev_icon = {"嚴重": "🔴", "中等": "🟡", "輕微": "🟢"}
                for a in weak:
                    icon = sev_icon.get(a.get("severity", ""), "⚪")
                    st.markdown(
                        f'<div class="card card-red">{icon} <strong>{a.get("topic","")}</strong>'
                        f' ({a.get("strand","")})<br><small>{a.get("evidence","")}</small></div>',
                        unsafe_allow_html=True,
                    )


# ============================================================
# TAB: 綜合分析
# ============================================================
if "combined" in results:
    with tabs[tab_idx]:
        tab_idx += 1
        combined = results["combined"]

        if combined.get("parse_error"):
            st.text(combined.get("raw_response", ""))
        else:
            # ── Diagnostic summary ──────────────────────────────────
            diag = combined.get("diagnostic_summary", {})
            if diag:
                st.markdown("### 🔬 深度學習診斷")
                st.error(f"**核心診斷：** {diag.get('overall_diagnosis','')}")
                st.info(f"**AQP 與試卷關聯：** {diag.get('aqp_exam_correlation','')}")

                # Key weak question cross-reference table
                kw = diag.get("key_weak_questions", [])
                if kw:
                    st.markdown("#### 📊 AQP 弱題 × 試卷題目 對照表")
                    kw_rows = [
                        {
                            "AQP 弱題": q.get("aqp_question", ""),
                            "AQP 正確率": f"{q.get('aqp_correct_rate','')}%" if q.get('aqp_correct_rate') is not None else "—",
                            "對應試卷題號": q.get("exam_question", ""),
                            "試卷題目內容": q.get("exam_question_content", ""),
                            "關聯說明": q.get("connection", ""),
                        }
                        for q in kw
                    ]
                    st.dataframe(pd.DataFrame(kw_rows), use_container_width=True, hide_index=True)

                mc = diag.get("misconception_vs_procedural", {})
                if mc:
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**🧠 概念性誤解**")
                        for ci in mc.get("conceptual_issues", []):
                            st.markdown(f"- {ci}")
                    with c2:
                        st.markdown("**🔢 程序性錯誤**")
                        for pi in mc.get("procedural_issues", []):
                            st.markdown(f"- {pi}")

            # ── Root cause per question ──────────────────────────────
            rca = combined.get("question_root_cause_analysis", [])
            if rca:
                st.markdown(f"### 🔍 逐題出錯根因分析（{len(rca)} 題）")
                cicon = {"正確": "✅", "錯誤": "❌", "部分正確": "⚠️"}
                etype_color = {
                    "概念性誤解": "card-red",
                    "程序性錯誤": "card-yellow",
                    "粗心大意": "card-blue",
                    "未完成課程": "card-red",
                    "語文理解困難": "card-yellow",
                }
                for q in rca:
                    corr = q.get("correctness", "")
                    if corr == "正確":
                        continue  # skip correct questions
                    css = etype_color.get(q.get("error_type", ""), "card")
                    icon = cicon.get(corr, "❓")
                    with st.expander(
                        f"{icon} 題 {q.get('question_ref','')} │ {q.get('topic','')} （{q.get('strand','')}）"
                    ):
                        if q.get("question_content"):
                            st.markdown(
                                f'<div class="card card-blue">📝 <strong>試卷題目：</strong>{q["question_content"]}</div>',
                                unsafe_allow_html=True,
                            )
                        st.markdown(f"**觀察到的錯誤：** {q.get('error_observed','—')}")
                        st.markdown(
                            f'<div class="card {css}">'
                            f"🎯 <strong>根本原因：</strong>{q.get('root_cause','')}<br>"
                            f"<small>錯誤類型：{q.get('error_type','')}　│　"
                            f"教學缺口：{q.get('teaching_gap','')}</small></div>",
                            unsafe_allow_html=True,
                        )
                        if q.get("aqp_evidence"):
                            st.caption(f"📊 AQP 佐證：{q['aqp_evidence']}")

            # ── Consolidated weak areas ──────────────────────────────
            weak_areas = combined.get("consolidated_weak_areas", [])
            if weak_areas:
                st.markdown("### ⚠️ 重點需改善範疇")
                p_icon = {"緊急": "🔴", "重要": "🟡", "一般": "🟢"}
                for a in weak_areas:
                    icon = p_icon.get(a.get("priority_level", ""), "⚪")
                    topics_str = "、".join(a.get("topics", []))
                    with st.expander(f"{icon} {a.get('strand','')} — {topics_str}"):
                        st.markdown(f"**深層原因：** {a.get('root_cause_analysis','')}")
                        c1, c2 = st.columns(2)
                        if a.get("aqp_evidence"):
                            c1.markdown(f"**AQP 佐證：** {a['aqp_evidence']}")
                        if a.get("exam_evidence"):
                            c2.markdown(f"**試卷佐證：** {a['exam_evidence']}")
                        if a.get("intervention_type"):
                            st.info(f"💡 介入方式：{a['intervention_type']}")

            # ── Priority interventions ───────────────────────────────
            priority = combined.get("priority_interventions", [])
            if priority:
                st.markdown("### 🚨 優先介入行動")
                for item in priority:
                    rank = item.get("rank", "")
                    with st.expander(f"#{rank} {item.get('weakness','')}"):
                        st.markdown(f"**優先原因：** {item.get('reason','')}")
                        st.success(f"⚡ 即時行動（下星期）：{item.get('immediate_action','')}")
                        if item.get("resources"):
                            st.markdown(f"**教材建議：** {item['resources']}")

            # ── Remediation plan ────────────────────────────────────
            plan = combined.get("remediation_plan", [])
            if plan:
                st.markdown("### 📅 補救教學計劃")
                for phase in plan:
                    with st.expander(f"📅 {phase.get('phase','')} — {phase.get('target_weakness','')}"):
                        st.markdown(f"**教學方法：** {phase.get('teaching_approach','')}")
                        acts = phase.get("practice_activities", [])
                        if acts:
                            st.markdown("**練習活動：**")
                            for act in acts:
                                st.markdown(f"- {act}")
                        if phase.get("success_criteria"):
                            st.markdown(f"**達成標準：** {phase['success_criteria']}")
                        if phase.get("assessment"):
                            st.markdown(f"**評估方法：** {phase['assessment']}")

            # ── Parent-teacher report ────────────────────────────────
            pt = combined.get("parent_teacher_report", {})
            if pt:
                st.markdown("### 👨‍👩‍👧 家校溝通")
                for finding in pt.get("key_findings", []):
                    st.markdown(f"- 📌 {finding}")
                home = pt.get("home_support", [])
                if home:
                    st.markdown("**家庭支援建議：**")
                    for h in home:
                        st.markdown(f"- 🏠 {h}")
                if pt.get("follow_up_timeline"):
                    st.info(f"📆 跟進時間表：{pt['follow_up_timeline']}")


# ============================================================
# TAB: 改善建議
# ============================================================
with tabs[tab_idx]:
    tab_idx += 1

    # Gather all recommendations across all analyses
    all_recs = []
    for key in ("aqp", "exam", "combined"):
        if key in results:
            for r in results[key].get("recommendations", []):
                r["_source"] = key
                all_recs.append(r)

    if not all_recs:
        st.info("建議已整合在「綜合分析」標籤頁中。")
    else:
        p_order = {"高": 0, "緊急": 0, "中": 1, "重要": 1, "低": 2, "一般": 2}
        all_recs.sort(key=lambda x: p_order.get(x.get("priority", "低"), 2))

        st.markdown("### 💡 優先改善建議")

        for level, label, css_class in [
            ({"高", "緊急"}, "🔴 高優先級", "card-red"),
            ({"中", "重要"}, "🟡 中優先級", "card-yellow"),
            ({"低", "一般"}, "🟢 一般建議", "card-green"),
        ]:
            group = [r for r in all_recs if r.get("priority", "低") in level]
            if not group:
                continue
            st.markdown(f"#### {label}")
            for rec in group:
                area = rec.get("area", "")
                action = rec.get("action") or rec.get("specific_action", "")
                resources = rec.get("resources") or rec.get("suggested_exercises", "")
                resource_html = f"<br><small>💡 建議練習：{resources}</small>" if resources else ""
                st.markdown(
                    f'<div class="card {css_class}"><strong>📌 {area}</strong><br>{action}{resource_html}</div>',
                    unsafe_allow_html=True,
                )


# ============================================================
# TAB: 匯出報告
# ============================================================
with tabs[tab_idx]:
    st.markdown("### 📥 匯出分析報告")

    # ── PDF ──────────────────────────────────────────────────
    st.markdown("#### 📄 PDF 報告（包含圖表）")
    st.caption("📌 包含 AQP 分析、試卷逐題分析、綜合診斷及改善建議 · 內嵌雷達圖、正確率柱狀圖及難度圈圖")
    if st.button("🔄 生成 PDF 報告", use_container_width=True, type="primary"):
        with st.spinner("正在生成 PDF，請稍候…"):
            try:
                pdf_bytes = build_pdf(results, grade, class_label)
                st.session_state["aqp_pdf_bytes"] = pdf_bytes
                st.session_state["aqp_pdf_stem"] = f"{grade}_{class_label or 'all'}"
            except Exception as e:
                st.error(f"❌ PDF 生成失敗：{e}")
                st.exception(e)

    if st.session_state.get("aqp_pdf_bytes"):
        file_stem = st.session_state.get("aqp_pdf_stem", "report")
        st.download_button(
            label="⬇️ 點擊下載 PDF 報告",
            data=st.session_state["aqp_pdf_bytes"],
            file_name=f"math_report_{file_stem}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.success(f"✅ PDF 已生成，共 {len(st.session_state['aqp_pdf_bytes']):,} bytes")

    st.markdown("---")

    # ── JSON ──────────────────────────────────────────────────
    st.markdown("#### 📂 其他格式")
    json_str = json.dumps(results, ensure_ascii=False, indent=2)
    st.download_button(
        label="📥 下載完整 JSON 報告",
        data=json_str.encode("utf-8-sig"),
        file_name=f"math_analysis_{class_label or 'class'}_{grade}.json",
        mime="application/json",
    )

    # Plain-text report
    def _build_text_report(res: dict, class_lbl: str, grd: str) -> str:
        lines = [
            "=" * 60,
            "小學數學全年級表現分析報告",
            f"年級：{grd}  班別：{class_lbl or '（全年級）'}",
            "=" * 60,
        ]
        sections = {
            "aqp": "AQP 報告分析",
            "exam": "試卷分析",
            "combined": "綜合分析",
        }
        for key, title in sections.items():
            if key not in res or res[key].get("parse_error"):
                continue
            data = res[key]
            lines += ["", f"{'=' * 20} {title} {'=' * 20}"]

            overall = data.get("overall_performance", {})
            if overall:
                lines.append(f"整體表現：{overall.get('performance_level','')}")
                lines.append(f"概述：{overall.get('summary','')}")

            for w in data.get("class_weak_areas", data.get("weak_areas", [])):
                desc = w.get("description") or w.get("evidence", "")
                lines.append(
                    f"  ⚠ {w.get('topic','')} ({w.get('strand','')}): {desc}"
                )

            for r in data.get("recommendations", []):
                action = r.get("action") or r.get("specific_action", "")
                lines.append(
                    f"  [{r.get('priority','')}] {r.get('area','')}: {action}"
                )

        lines += ["", "=" * 60, "報告由 Qwen AI 生成 | 小學數學全年級表現分析系統"]
        lines += ["課程參考：香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版"]
        return "\n".join(lines)

    text_report = _build_text_report(results, class_label, grade)
    st.download_button(
        label="📄 下載文字報告 (.txt)",
        data=text_report.encode("utf-8-sig"),
        file_name=f"math_report_{class_label or 'class'}_{grade}.txt",
        mime="text/plain",
    )

    st.markdown("---")
    st.markdown("#### 📋 原始 JSON 預覽")
    with st.expander("展開查看原始分析數據"):
        st.json(results)
