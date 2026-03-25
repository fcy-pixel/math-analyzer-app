"""
practice_html.py — Build printable A4 practice worksheets (one per student)
for Hong Kong primary school math teachers.

Usage
-----
from practice_html import build_practice_worksheets_html

html_str = build_practice_worksheets_html(all_pq_results, grade="P4", show_answers=False)
# → student copy (no answers)

html_str = build_practice_worksheets_html(all_pq_results, grade="P4", show_answers=True)
# → teacher / answer copy
"""

import html
from datetime import date
from typing import Dict, List, Optional


# ---------------------------------------------------------------------------
# CSS — A4 print-optimised, CJK-friendly, no external requests
# ---------------------------------------------------------------------------
_CSS = """
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: "Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC",
                 "Source Han Sans TC", sans-serif;
    font-size: 12pt;
    color: #1a1a1a;
    background: #d8dce0;
  }

  /* ── Print controls (hidden on print) ─────────────────────────── */
  .print-controls {
    background: #1e3a5f;
    color: #fff;
    text-align: center;
    padding: 14px 20px;
    position: sticky;
    top: 0;
    z-index: 100;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 12px;
    flex-wrap: wrap;
  }
  .print-controls p { font-size: 13px; opacity: 0.85; margin-right: 8px; }
  .btn {
    padding: 8px 22px;
    border: none;
    border-radius: 5px;
    font-size: 13px;
    cursor: pointer;
    font-family: inherit;
    font-weight: 600;
  }
  .btn-primary { background: #f0a500; color: #1a1a1a; }
  .btn-primary:hover { background: #e09400; }
  .btn-secondary { background: transparent; color: #fff; border: 1.5px solid rgba(255,255,255,0.6); }
  .btn-secondary:hover { background: rgba(255,255,255,0.15); }

  /* ── A4 page ───────────────────────────────────────────────────── */
  .page {
    width: 210mm;
    min-height: 297mm;
    margin: 10mm auto;
    background: white;
    padding: 14mm 16mm 22mm 16mm;
    box-shadow: 0 4px 20px rgba(0,0,0,0.2);
    position: relative;
    page-break-after: always;
  }

  /* ── Worksheet header ──────────────────────────────────────────── */
  .ws-header {
    border-bottom: 3px double #1e3a5f;
    padding-bottom: 10px;
    margin-bottom: 12px;
  }
  .ws-super {
    text-align: center;
    font-size: 9.5pt;
    color: #1e3a5f;
    letter-spacing: 0.5px;
    margin-bottom: 4px;
  }
  .ws-title {
    text-align: center;
    font-size: 17pt;
    font-weight: 700;
    color: #1e3a5f;
    margin-bottom: 12px;
  }
  .ws-fields {
    display: flex;
    gap: 8px;
  }
  .ws-field {
    flex: 1;
    border-bottom: 1.5px solid #555;
    padding: 2px 0 3px 0;
    font-size: 11pt;
    min-height: 26px;
  }
  .ws-field-label {
    font-size: 9pt;
    color: #555;
    margin-right: 3px;
  }

  /* ── Weakness note ─────────────────────────────────────────────── */
  .weakness-note {
    background: #fff8e1;
    border-left: 4px solid #e09400;
    border-radius: 0 4px 4px 0;
    padding: 6px 12px;
    margin: 10px 0 4px 0;
    font-size: 10pt;
    color: #5d3a00;
    line-height: 1.5;
  }

  /* ── Question block ────────────────────────────────────────────── */
  .q-block {
    border: 1px solid #bbb;
    border-radius: 5px;
    margin: 11px 0;
    overflow: hidden;
    break-inside: avoid;
  }
  .q-head {
    background: #1e3a5f;
    color: #fff;
    padding: 5px 12px;
    font-size: 10pt;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
  .q-num { font-weight: 700; font-size: 12pt; }
  .q-type-tag {
    display: inline-block;
    background: rgba(255,255,255,0.2);
    padding: 1px 8px;
    border-radius: 3px;
    font-size: 9pt;
    margin-left: 6px;
  }
  .q-topic { font-size: 9pt; opacity: 0.8; }

  .q-body { padding: 10px 14px 8px 14px; }
  .q-text {
    font-size: 12.5pt;
    line-height: 1.75;
    margin-bottom: 8px;
    white-space: pre-wrap;
  }
  .hint-box {
    background: #e3f2fd;
    border-radius: 4px;
    padding: 4px 10px;
    font-size: 9.5pt;
    color: #1a3c5c;
    margin-bottom: 8px;
  }
  .work-space {
    border: 1px dashed #bbb;
    border-radius: 4px;
    background: #fafafa;
    min-height: 50px;
    padding: 6px 10px;
    font-size: 9pt;
    color: #aaa;
  }

  /* ── Answer section ────────────────────────────────────────────── */
  .ans-section {
    background: #fffde7;
    border-top: 1px dashed #ffd54f;
    padding: 8px 14px 10px 14px;
    font-size: 10.5pt;
  }
  .ans-label {
    font-weight: 700;
    color: #b06000;
    margin-bottom: 5px;
  }
  .steps-list {
    list-style: none;
    padding: 0;
    counter-reset: step;
    margin: 4px 0 0 0;
  }
  .steps-list li {
    counter-increment: step;
    padding: 2px 0 2px 20px;
    font-size: 10.5pt;
    position: relative;
    line-height: 1.6;
  }
  .steps-list li::before {
    content: counter(step) ". ";
    position: absolute;
    left: 0;
    font-weight: 700;
    color: #1e3a5f;
  }

  /* ── Study tips ────────────────────────────────────────────────── */
  .tips-box {
    margin-top: 14px;
    background: #e8f5e9;
    border: 1px solid #a5d6a7;
    border-radius: 5px;
    padding: 8px 14px;
    font-size: 10pt;
  }
  .tips-title {
    font-weight: 700;
    color: #2e7d32;
    margin-bottom: 6px;
    font-size: 11pt;
  }
  .tip-item { padding: 2px 0; line-height: 1.5; }
  .tip-item::before { content: "📌 "; }

  /* ── Page footer ───────────────────────────────────────────────── */
  .pg-footer {
    position: absolute;
    bottom: 10mm;
    left: 16mm;
    right: 16mm;
    border-top: 1px solid #ddd;
    padding-top: 4px;
    display: flex;
    justify-content: space-between;
    font-size: 8pt;
    color: #999;
  }

  /* ── Print rules ───────────────────────────────────────────────── */
  @media print {
    body { background: white; }
    .print-controls { display: none !important; }
    .page {
      width: 100%;
      margin: 0;
      padding: 12mm 14mm 22mm 14mm;
      box-shadow: none;
      min-height: unset;
    }
  }
</style>
"""


def _e(text) -> str:
    """HTML-escape a value."""
    if text is None:
        return ""
    return html.escape(str(text))


def _build_student_page(pq_result: Dict, grade: str, show_answers: bool) -> str:
    """Return the HTML for one student's A4 worksheet page."""
    name = _e(pq_result.get("student_name", "學生"))
    weakness_summary = _e(pq_result.get("weakness_summary", ""))
    questions = pq_result.get("practice_questions", [])
    tips = pq_result.get("study_tips", [])
    today = date.today().strftime("%Y年%m月%d日")
    copy_label = "【老師版 · 含答案及解題步驟】" if show_answers else "【學生練習版】"
    total_marks = len(questions) * 2

    parts: List[str] = ['<div class="page">']

    # Header
    parts.append(f"""
<div class="ws-header">
  <div class="ws-super">小學數學 弱點針對練習 · {_e(grade)} · {copy_label}</div>
  <div class="ws-title">📝 數學弱點鞏固練習題</div>
  <div class="ws-fields">
    <div class="ws-field"><span class="ws-field-label">姓名：</span>{name}</div>
    <div class="ws-field"><span class="ws-field-label">班別：</span>&nbsp;</div>
    <div class="ws-field"><span class="ws-field-label">日期：</span>{today}</div>
    <div class="ws-field"><span class="ws-field-label">得分：</span>_____ / {total_marks}</div>
  </div>
</div>""")

    # Weakness summary
    if weakness_summary:
        parts.append(f"""
<div class="weakness-note">🎯 <strong>練習重點：</strong>{weakness_summary}</div>""")

    # Questions
    for q in questions:
        qn = q.get("question_number", "")
        qtype = _e(q.get("question_type", ""))
        strand = _e(q.get("strand", ""))
        topic = _e(q.get("topic", ""))
        qtext = _e(q.get("question_text", "").replace("\\n", "\n"))
        hint = q.get("hints", "")
        answer = _e(q.get("answer", ""))
        steps = q.get("solution_steps", [])

        parts.append(f"""
<div class="q-block">
  <div class="q-head">
    <span>
      <span class="q-num">第 {qn} 題</span>
      <span class="q-type-tag">{qtype}</span>
    </span>
    <span class="q-topic">{strand}&nbsp;·&nbsp;{topic}</span>
  </div>
  <div class="q-body">
    <div class="q-text">{qtext}</div>""")

        if hint:
            parts.append(f'\n    <div class="hint-box">💡 提示：{_e(hint)}</div>')

        parts.append('\n    <div class="work-space">（計算工作空間）</div>')
        parts.append('\n  </div>')

        if show_answers:
            steps_html = ""
            if steps:
                items = "".join(f"<li>{_e(s)}</li>" for s in steps)
                steps_html = f'<ol class="steps-list">{items}</ol>'
            parts.append(f"""
  <div class="ans-section">
    <div class="ans-label">✅ 正確答案：{answer}</div>
    {steps_html}
  </div>""")

        parts.append('\n</div>')  # end q-block

    # Study tips
    if tips:
        tip_items = "".join(f'<div class="tip-item">{_e(t)}</div>' for t in tips)
        parts.append(f"""
<div class="tips-box">
  <div class="tips-title">📚 學習建議</div>
  {tip_items}
</div>""")

    # Page footer
    parts.append(f"""
<div class="pg-footer">
  <span>{name}　{_e(grade)}</span>
  <span>弱點針對練習 — 小學數學分析系統</span>
  <span>{today}</span>
</div>""")

    parts.append('\n</div>')  # end .page
    return "".join(parts)


def build_practice_worksheets_html(
    all_pq_results: List[Dict],
    grade: str,
    show_answers: bool = False,
) -> str:
    """
    Build a single self-contained HTML file with one A4 page per student.
    Includes a sticky toolbar with a 「列印全部」 button.
    Use show_answers=True for the teacher / answer copy.

    Parameters
    ----------
    all_pq_results : list of dicts from MathAnalyzer.generate_practice_questions()
    grade          : e.g. "P4"
    show_answers   : False → student copy  |  True → teacher copy with answers
    """
    valid_results = [r for r in all_pq_results if r and not r.get("parse_error")]

    if not valid_results:
        return "<html><body><p>沒有可用的練習題資料。</p></body></html>"

    copy_label = "老師版（含答案）" if show_answers else "學生練習版"
    n = len(valid_results)
    doc_title = f"數學弱點練習 · {grade} · {copy_label}"
    today = date.today().strftime("%Y年%m月%d日")

    student_pages = "\n".join(
        _build_student_page(r, grade, show_answers) for r in valid_results
    )

    toggle_btn = ""
    toggle_js = ""
    if show_answers:
        toggle_btn = '<button class="btn btn-secondary" id="toggleBtn" onclick="toggleAnswers()">隱藏答案</button>'
        toggle_js = """
<script>
function toggleAnswers() {
  var secs = document.querySelectorAll('.ans-section');
  var btn  = document.getElementById('toggleBtn');
  var hide = secs.length && secs[0].style.display !== 'none';
  secs.forEach(function(el) { el.style.display = hide ? 'none' : ''; });
  btn.textContent = hide ? '顯示答案' : '隱藏答案';
}
</script>"""

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{_e(doc_title)}</title>
  {_CSS}
  {toggle_js}
</head>
<body>

<div class="print-controls">
  <p>共 <strong>{n}</strong> 位學生的練習題 · {_e(copy_label)} · {today}</p>
  <button class="btn btn-primary" onclick="window.print()">🖨️ 列印全部（{n} 頁）</button>
  {toggle_btn}
</div>

{student_pages}

</body>
</html>"""
