"""
Qwen API integration for primary school mathematics performance analysis.
Uses Qwen International API (OpenAI-compatible endpoint).

Key design:
- AQP report   → class-wide weak area extraction (全級弱點)
- Exam paper   → full question scan, page-by-page via vision batching
- Combined     → cross-reference: WHY students make mistakes per question
"""

import json
import re
from typing import Dict, List, Optional


def _natural_sort_key(s: str) -> list:
    """Sort key that handles embedded numbers naturally: Q2 < Q10, 1a < 1b < 2a."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", str(s))]

from openai import OpenAI

from curriculum_hk import CURRICULUM_STRANDS, GRADE_LEARNING_OBJECTIVES, get_grade_curriculum
from file_processor import VISION_BATCH_SIZE, get_image_batches

# ---------------------------------------------------------------------------
# Qwen International API configuration
# ---------------------------------------------------------------------------
QWEN_BASE_URL = "https://dashscope-intl.aliyuncs.com/compatible-mode/v1"
TEXT_MODEL = "qwen-max"
VISION_MODEL = "qwen-vl-max"


class MathAnalyzer:
    """Analyze student math performance using Qwen AI."""

    def __init__(self, api_key: str):
        if not api_key or not api_key.strip():
            raise ValueError("請提供有效的 Qwen API Key。")
        self.client = OpenAI(api_key=api_key.strip(), base_url=QWEN_BASE_URL)

    # ------------------------------------------------------------------
    # 1. AQP class-wide report analysis
    # ------------------------------------------------------------------

    def analyze_aqp(self, aqp_data: Dict, grade: str, student_name: str = "") -> Dict:
        """
        Analyse a CLASS-WIDE AQP report.
        The AQP report summarises the whole class / cohort performance.
        We identify shared weak areas across the class.
        """
        curriculum_info = get_grade_curriculum(grade)
        images = aqp_data.get("images", [])
        text_data = aqp_data.get("text_summary", "")

        system_msg = (
            "你是一位擁有豐富經驗的香港小學數學科主任，專門解讀 AQP（全港基本能力評估）"
            "成績報告，並能從全年級數據中準確找出學生普遍弱點。"
            "你熟悉香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版，"
            "能將 AQP 弱點準確對應到相關課程範疇和學習目標。"
        )

        # Pre-compute so no backslash inside f-string expressions (Python 3.9)
        text_section = ("報告文字內容：\n" + text_data) if text_data else ""

        prompt = f"""以下是{grade}年級的 AQP（全港基本能力評估計劃）**全年級整體**成績報告。
注意：這是全年級報告，反映整個年級所有班別學生的共同表現，並非單一學生或班別的成績。

{text_section}

香港課程發展議會《數學課程指引》（2017）{grade}年級課程綱要：
{curriculum_info}

請逐頁仔細閱讀報告中的所有數據、圖表、分項分數，然後完成以下分析：
1. 逐題列出每題的全年級正確率（百分比），並找出正確率低於60%的弱題
2. 將弱題按正確率由低至高排列，列出弱題排行榜
3. 找出全年級學生在哪些課程範疇和題型上表現最弱，並對應到《數學課程指引》（2017）的相關學習目標
4. 找出全年級的強項範疇
5. 推斷全班學生在哪些概念上存在共同的知識盲點或誤解
6. 提供具體數據（正確率%、得分率%）作為每項弱點的佐證

**重要：弱題分析必須以具體數字作證據，例如「第5題正確率僅23%」而非籠統描述。每項弱點須對應到《數學課程指引》（2017）的具體課程範疇。**

以**純 JSON** 格式回應（不要加 markdown 代碼塊），結構如下：
{{
  "report_scope": "全級報告",
  "overall_performance": {{
    "summary": "全班整體表現概述（3-5句，必須包含具體數據如全班平均分、各範疇正確率）",
    "class_average_percentage": 數字或 null,
    "performance_level": "優秀 / 良好 / 一般 / 需要改善"
  }},
  "strand_analysis": [
    {{
      "strand": "課程範疇名稱",
      "class_score": 全班平均分或得分率（數字）或 null,
      "performance": "全班在此範疇的整體表現描述（包含具體分數或正確率）",
      "status": "強項 / 一般 / 弱項",
      "specific_topics_struggled": ["全班普遍困難的具體主題1", "主題2"]
    }}
  ],
  "question_performance": [
    {{
      "question_ref": "題號（如：第1題、Q2、Section A Q3）",
      "topic": "考核主題",
      "strand": "課程範疇",
      "class_correct_rate": 全班正確率百分比（數字）或 null,
      "difficulty": "容易 / 中等 / 困難",
      "common_errors": "全班學生常見的錯誤描述"
    }}
  ],
  "weak_questions": [
    {{
      "rank": 排名數字（1=最弱）,
      "question_ref": "題號",
      "correct_rate": 正確率百分比（數字），
      "topic": "考核主題",
      "strand": "課程範疇",
      "common_error": "學生常見的具體錯誤",
      "severity": "嚴重（<40%） / 中等（40-59%） / 留意（60-69%）"
    }}
  ],
  "class_weak_areas": [
    {{
      "topic": "弱點主題",
      "strand": "所屬課程範疇",
      "description": "全班在此主題的具體弱點描述（包含數據佐證，如：第X題正確率僅Y%）",
      "likely_misconception": "學生可能存在的概念誤解或知識盲點",
      "severity": "嚴重 / 中等 / 輕微",
      "affected_question_types": ["涉及的題目類型1", "類型2"],
      "data_evidence": "具體數據佐證（正確率、得分率等）"
    }}
  ],
  "class_strong_areas": [
    {{
      "topic": "強項主題",
      "strand": "所屬課程範疇",
      "description": "全班在此主題的優秀表現描述（包含數據）"
    }}
  ],
  "teaching_implications": [
    {{
      "issue": "教學問題",
      "evidence": "AQP 報告中的佐證數據（必須包含具體題號和正確率）",
      "suggested_teaching_strategy": "建議的教學策略"
    }}
  ],
  "recommendations": [
    {{
      "priority": "高 / 中 / 低",
      "area": "需改善範疇",
      "action": "具體建議行動",
      "resources": "建議練習類型或資源"
    }}
  ]
}}"""

        merge_instruction = f"""你是香港小學數學科主任。
以下是對{grade}年級 AQP 報告各批次頁面的分析結果。
請整合所有批次的資料，生成一份完整的全班 AQP 分析。
**必須保留所有具體數字（正確率%、得分率%）作為弱點佐證。**
**weak_questions 必須列出所有正確率低於60%的題目，按正確率由低至高排列。**

各批次分析：
{{{{batch_results}}}}

請以**純 JSON** 格式輸出，結構如下：
{{{{
  "report_scope": "全級報告",
  "overall_performance": {{"summary":"","class_average_percentage":null,"performance_level":""}},
  "strand_analysis": [],
  "question_performance": [],
  "weak_questions": [],
  "class_weak_areas": [],
  "class_strong_areas": [],
  "teaching_implications": [],
  "recommendations": []
}}}}"""

        if images:
            raw = self._vision_batched(
                base_prompt=prompt,
                images_b64=images,
                system_msg=system_msg,
                merge_instruction=merge_instruction,
            )
        else:
            raw = self._text(prompt, system_msg)

        return self._parse_json(raw)

    # ------------------------------------------------------------------
    # 2. Exam paper — full question scan
    # ------------------------------------------------------------------

    def analyze_exam(self, exam_data: Dict, grade: str, student_name: str = "") -> Dict:
        """
        Full exam paper analysis using qwen-vl-max.
        The uploaded paper is an ANSWER VERSION (答案版) for reference.
        ALL pages are scanned in batches. Every question is catalogued.
        """
        curriculum_info = get_grade_curriculum(grade)
        images = exam_data.get("images", [])
        text_fallback = exam_data.get("text", "")
        page_count = exam_data.get("page_count", len(images) or 1)

        system_msg = (
            "你是一位專業的香港小學數學教師，正在分析一份數學試卷的答案版（參考答案卷）。"
            "這份試卷是供教師參考用的答案版，不是學生的作答卷。"
            "你熟悉香港課程發展議會《數學課程指引》（小一至六年級）2017 年修訂版，"
            "能將每道題目準確對應到相關課程範疇、學習目標及年級要求。"
            "你的任務是逐題辨認題目內容、對應課程學習目標、正確解題方法，以及評估每題的難度，"
            "務必不漏掉任何一題。"
        )

        per_batch_prompt = """請仔細閱讀這批試卷答案版頁面（第 {start}–{end} 頁，共 {total} 頁的一部分）。

注意：這是試卷的**答案版（參考答案卷）**，不是學生作答卷。
請對這幾頁上的**每一道題目**進行分析，包括：
- 題目編號（如：第1題、(a)、Q3 等）
- 題目內容的完整描述
- 題目考核的數學概念、課程範疇及對應《數學課程指引》（2017）的學習目標
- 正確答案或正確解題方法（從答案版中讀取）
- 題目分值（如標示）
- 題目難度評估
- 預計學生容易在此題犯的錯誤（根據題目設計及課程要求推斷）

以**純 JSON** 格式回應，結構如下：
{{
  "pages_scanned": ["{start}至{end}"],
  "questions_found": [
    {{
      "question_ref": "題目編號",
      "page": 頁碼,
      "question_description": "題目內容的完整描述",
      "topic": "考核的數學主題",
      "strand": "課程範疇",
      "marks": 分值或 null,
      "correct_answer": "正確答案或正確解題步驟（從答案版讀取）",
      "solution_method": "解題方法描述",
      "difficulty": "容易 / 中等 / 困難",
      "predicted_errors": "預計學生在此題容易犯的錯誤（根據題目設計推斷）"
    }}
  ]
}}"""

        if images:
            batches = get_image_batches(images, VISION_BATCH_SIZE)
            batch_jsons = []

            for i, batch in enumerate(batches):
                start_page = i * VISION_BATCH_SIZE + 1
                end_page = min(start_page + len(batch) - 1, page_count)
                p = per_batch_prompt.format(
                    start=start_page, end=end_page, total=page_count
                )
                raw = self._vision(p, batch, system_msg)
                parsed = self._parse_json(raw)
                batch_jsons.append(parsed)

            raw = self._synthesize_exam(batch_jsons, grade, curriculum_info, page_count)
        else:
            raw = self._text(
                self._build_exam_text_prompt(text_fallback, grade, curriculum_info),
                system_msg,
            )

        return self._parse_json(raw)

    def _synthesize_exam(
        self, batch_jsons: List[Dict], grade: str, curriculum_info: str, page_count: int
    ) -> str:
        """Merge per-batch question lists into a complete exam analysis (answer-key version)."""
        all_questions = []
        for b in batch_jsons:
            all_questions.extend(b.get("questions_found", []))

        questions_json = json.dumps(all_questions, ensure_ascii=False, indent=2)

        prompt = f"""你是一位香港小學數學科主任。
以下是從一份{grade}年級數學試卷**答案版**（共 {page_count} 頁）逐批掃描得到的所有題目列表：

{questions_json}

香港{grade}年級數學課程綱要：
{curriculum_info}

注意：這份試卷是**答案版（參考答案卷）**，不是學生作答卷。
根據上述題目清單，生成完整的試卷結構分析報告，重點說明：
1. 每題考核的數學概念和難度
2. 試卷的整體結構和分配
3. 根據題目設計，預測學生可能在哪些題目上出錯，及可能的出錯原因

請以**純 JSON** 格式回應（不要加 markdown 代碼塊），結構如下：
{{
  "exam_overview": {{
    "total_questions": 題目總數,
    "total_pages": {page_count},
    "topics_covered": ["涵蓋主題列表"],
    "estimated_difficulty": "容易 / 中等 / 困難",
    "strands_tested": ["課程範疇列表"],
    "marks_distribution": {{"範疇名": 分值}}
  }},
  "question_analysis": [
    {{
      "question_ref": "題目編號",
      "page": 頁碼,
      "question_description": "題目內容的完整描述",
      "topic": "考核主題",
      "strand": "課程範疇",
      "marks": 分值或 null,
      "correct_answer": "正確答案（從答案版讀取）",
      "solution_method": "正確解題方法描述",
      "difficulty": "容易 / 中等 / 困難",
      "predicted_errors": "根據題目設計，預測學生容易犯的錯誤"
    }}
  ],
  "predicted_error_patterns": [
    {{
      "pattern": "預測的錯誤模式描述",
      "related_concept": "相關數學概念",
      "affected_questions": ["題號1", "題號2"],
      "reason": "為何學生容易在此類題目出錯"
    }}
  ],
  "challenging_areas": [
    {{
      "topic": "較具挑戰性的主題",
      "strand": "所屬課程範疇",
      "questions": ["相關題號"],
      "reason": "為何此主題對學生有挑戰性",
      "severity": "高 / 中 / 低"
    }}
  ],
  "accessible_areas": [
    {{
      "topic": "相對容易的主題",
      "strand": "所屬課程範疇",
      "questions": ["相關題號"]
    }}
  ],
  "recommendations": [
    {{
      "priority": "高 / 中 / 低",
      "area": "需加強範疇",
      "specific_action": "具體教學或練習建議",
      "suggested_exercises": "建議題型或練習方向"
    }}
  ]
}}"""
        return self._text(
            prompt,
            "你是專業香港小學數學科主任，熟悉香港課程發展議會《數學課程指引》（2017），擅長分析試卷結構並按課程要求預測學生學習困難。",
        )

    def _build_exam_text_prompt(self, text: str, grade: str, curriculum_info: str) -> str:
        return f"""你是一位香港小學數學科主任，正在分析以下{grade}年級數學試卷答案版的文字內容。

注意：這是試卷的**答案版（參考答案卷）**，不是學生作答卷。

試卷內容：
{text}

香港{grade}年級數學課程綱要：
{curriculum_info}

請逐題分析每題的考核概念、正確答案、解題方法及預測學生可能出錯之處。
以**純 JSON** 格式回應：
{{
  "exam_overview": {{"total_questions":null,"topics_covered":[],"estimated_difficulty":"","strands_tested":[],"marks_distribution":{{}}}},
  "question_analysis": [],
  "predicted_error_patterns": [],
  "challenging_areas": [],
  "accessible_areas": [],
  "recommendations": []
}}"""

    # ------------------------------------------------------------------
    # 3. Combined cross-reference analysis
    # ------------------------------------------------------------------

    def combined_analysis(
        self,
        aqp_result: Dict,
        exam_result: Dict,
        grade: str,
        student_name: str = "",
    ) -> Dict:
        """
        Cross-reference AQP class-wide weak areas with specific exam questions
        to reason WHY students make mistakes on each question.
        """
        aqp_weak = aqp_result.get("class_weak_areas", aqp_result.get("weak_areas", []))
        aqp_strands = aqp_result.get("strand_analysis", [])
        aqp_qperf = aqp_result.get("question_performance", [])
        aqp_weak_questions = aqp_result.get("weak_questions", [])
        exam_questions = exam_result.get("question_analysis", [])
        exam_predicted_errors = exam_result.get("predicted_error_patterns", exam_result.get("error_patterns", []))
        exam_challenging = exam_result.get("challenging_areas", exam_result.get("weak_areas", []))

        prompt = f"""你是一位資深香港小學數學科主任，正在為{grade}年級進行全年級深度學習診斷。

你有以下資料：
- AQP 報告為**全年級**數據，反映整個年級所有班別學生的共同表現
- 試卷為**答案版（參考答案卷）**，用以了解每題考核的概念和難度，並推斷學生出錯的根本原因

【A】AQP 弱題排行榜（按正確率由低至高）：
{json.dumps(aqp_weak_questions, ensure_ascii=False, indent=2)}

【B】AQP 全班各題正確率（完整列表）：
{json.dumps(aqp_qperf, ensure_ascii=False, indent=2)}

【C】AQP 全班弱點分析：
{json.dumps(aqp_weak, ensure_ascii=False, indent=2)}

【D】AQP 各範疇表現：
{json.dumps(aqp_strands, ensure_ascii=False, indent=2)}

【E】試卷逐題分析（答案版——包含每題題目內容、正確答案、解題方法、預測學生出錯情況）：
{json.dumps(exam_questions, ensure_ascii=False, indent=2)}

【F】試卷預測的學生錯誤模式（根據題目設計推斷）：
{json.dumps(exam_predicted_errors, ensure_ascii=False, indent=2)}

【G】試卷較具挑戰性的題目範疇：
{json.dumps(exam_challenging, ensure_ascii=False, indent=2)}

你的任務：
1. **交叉比對**：對每條在 AQP 中正確率低的弱題，在試卷答案版中找出考查同一概念的題目，引用試卷題目的具體內容（question_description）和正確答案（correct_answer）作為佐證
2. **根因分析**：結合 AQP 全年級正確率數據及試卷答案版的題目設計，推論全年級學生出錯的根本原因——是概念不理解、步驟錯誤，還是題目理解困難？
3. **數據佐證**：每項弱點分析**必須引用**具體數字，例如「AQP 第5題全年級正確率僅23%，而試卷第3題（考查分數加減法，正確答案為3/8）的題目設計亦要求相同概念」
4. 區分概念性誤解與程序性錯誤
5. 制定具體補救教學計劃

以**純 JSON** 格式回應（不要加 markdown 代碼塊），結構如下：
{{
  "diagnostic_summary": {{
    "overall_diagnosis": "全班學習問題的根本診斷（3-5句，必須包含具體數據如弱題正確率）",
    "aqp_exam_correlation": "AQP 弱點與試卷失分的關聯分析（引用具體題號和正確率作佐證）",
    "key_weak_questions": [
      {{
        "aqp_question": "AQP 弱題題號",
        "aqp_correct_rate": 正確率百分比,
        "exam_question": "對應試卷題號",
        "exam_question_content": "試卷題目內容描述（直接引用 question_description）",
        "connection": "兩題考查相同概念的說明"
      }}
    ],
    "misconception_vs_procedural": {{
      "conceptual_issues": ["概念性誤解1（附數據佐證）", "概念性誤解2"],
      "procedural_issues": ["程序性錯誤1（附數據佐證）", "程序性錯誤2"]
    }}
  }},
  "question_root_cause_analysis": [
    {{
      "question_ref": "試卷題目編號",
      "question_content": "試卷題目內容描述（直接引用）",
      "topic": "考核主題",
      "strand": "課程範疇",
      "correctness": "正確 / 錯誤 / 部分正確",
      "error_observed": "試卷觀察到的具體錯誤",
      "aqp_evidence": "對應 AQP 題目的具體數據（如：AQP 第X題全年級正確率Y%）",
      "root_cause": "出錯的根本原因（從概念理解、知識遷移、計算步驟等角度分析）",
      "error_type": "概念性誤解 / 程序性錯誤 / 粗心大意 / 未完成課程 / 語文理解困難",
      "teaching_gap": "此錯誤反映的教學缺口"
    }}
  ],
  "consolidated_weak_areas": [
    {{
      "strand": "課程範疇",
      "topics": ["具體主題1", "具體主題2"],
      "aqp_evidence": "AQP 全年級數據佐證（具體題號和正確率）",
      "exam_evidence": "試卷答案版佐證（具體題號、題目內容摘要及正確答案）",
      "root_cause_analysis": "深層原因：為何學生在此範疇持續出錯",
      "priority_level": "緊急 / 重要 / 一般",
      "intervention_type": "概念重教 / 練習強化 / 解題策略訓練 / 語文支援"
    }}
  ],
  "remediation_plan": [
    {{
      "phase": "第一階段（第1-2週）",
      "target_weakness": "本階段針對的弱點（列出具體 AQP 弱題題號）",
      "teaching_approach": "具體教學方法",
      "practice_activities": ["練習活動1", "練習活動2"],
      "success_criteria": "達成標準（可量化）",
      "assessment": "評估方法"
    }},
    {{
      "phase": "第二階段（第3-4週）",
      "target_weakness": "本階段針對的弱點",
      "teaching_approach": "具體教學方法",
      "practice_activities": ["練習活動1", "練習活動2"],
      "success_criteria": "達成標準",
      "assessment": "評估方法"
    }},
    {{
      "phase": "第三階段（第5-6週）",
      "target_weakness": "本階段針對的弱點",
      "teaching_approach": "具體教學方法",
      "practice_activities": ["練習活動1", "練習活動2"],
      "success_criteria": "達成標準",
      "assessment": "評估方法"
    }}
  ],
  "priority_interventions": [
    {{
      "rank": 1,
      "weakness": "最優先處理的弱點（附 AQP 正確率數據）",
      "reason": "為何優先（結合 AQP 嚴重程度及試卷失分比重）",
      "immediate_action": "下星期即時行動",
      "resources": "建議練習題型或教材"
    }}
  ],
  "parent_teacher_report": {{
    "key_findings": ["重要發現1（附具體數據）", "重要發現2", "重要發現3"],
    "home_support": ["家庭支援建議1", "家庭支援建議2"],
    "follow_up_timeline": "建議跟進時間表"
  }}
}}"""

        raw = self._text(
            prompt,
            "你是資深香港小學數學科主任，熟悉香港課程發展議會《數學課程指引》（2017）及全港基本能力評估（AQP）框架，"
            "擅長整合全年級 AQP 數據與試卷答案版進行精準弱點診斷，分析必須以具體數字和題目內容為證據。",
        )
        return self._parse_json(raw)

    # ------------------------------------------------------------------
    # 4. Student exam paper batch analysis
    # ------------------------------------------------------------------

    def analyze_student_paper(
        self,
        images: List[str],
        question_schema: List[Dict],
        grade: str,
        student_name: str = "學生",
    ) -> Dict:
        """
        Analyze ONE student's scanned exam paper using vision AI.

        question_schema: list of question dicts from answer-key analysis.
            If empty, the AI will infer questions and score without an answer key.
        Returns a per-question scored result dict.
        """
        n_pages = len(images)
        system_msg = (
            "你是一位有豐富批改經驗的香港小學數學教師，"
            "熟悉香港課程發展議會《數學課程指引》（2017），"
            "擅長辨認學生的手寫答案及識別常見數學錯誤類型。"
            "重要評分規則（必須嚴格遵守）：\n"
            "(1) 假分數（improper fraction）必須直接給滿分。例如答案是 1 3/4，學生寫 7/4，直接給滿分，is_correct 設為 true。\n"
            "(2) 帶分數與假分數互換一律正確：7/4 = 1 3/4、13/5 = 2 3/5，只要數值相等就給滿分。\n"
            "(3) 絕對不可因為學生用假分數作答而扣分、標記錯誤或備註為非標準答案。"
        )

        out_schema = (
            '{"student_name":"...","total_marks_awarded":數字,"total_marks_possible":數字,'
            '"percentage":數字,"performance_level":"優秀(≥85%) / 良好(70-84%) / 一般(55-69%) / 需要改善(<55%)",'
            '"question_results":[{"question_ref":"題號","topic":"考核主題",'
            '"strand":"課程範疇","marks_possible":分值,"marks_awarded":得分,'
            '"is_correct":true/false,"student_answer":"學生作答",'
            '"error_type":"概念性誤解/程序性錯誤/粗心大意/未作答/null",'
            '"error_description":"錯誤描述或null"}],"overall_remarks":"簡短評語"}'
        )

        if question_schema:
            schema_json = json.dumps(question_schema, ensure_ascii=False, indent=2)
            prompt = f"""你正在批改 {student_name} 的 {grade} 年級數學試卷（共 {n_pages} 頁）。

本次試卷各題正確答案：
{schema_json}

請仔細閱讀每頁學生作答，逐題：
1. 辨認學生答案（手寫可能字跡潦草，請盡力判讀）
2. 根據答案鍵評正
3. 如答錯，填寫error_type（概念性誤解/程序性錯誤/粗心大意/未作答）和error_description
4. 答案鍵內每題必須評分，找不到視為「未作答」
5. **假分數必須直接給滿分**：若答案鍵為帶分數（如 1 3/4），學生寫假分數（如 7/4），必須直接給該題滿分，is_correct 設為 true，反之亦然。7/4 和 1 3/4 是完全等價的正確答案，不可扣分、不可標注錯誤、不可加任何備註

只輸出純JSON（不加markdown代碼塊）：
{out_schema}"""
        else:
            prompt = f"""你正在批改 {student_name} 的 {grade} 年級數學試卷（共 {n_pages} 頁）。

沒有答案鍵，請：
1. 識別所有題目（包括子題(a)(b)(c)等）
2. 根據數學知識判斷學生的作答是否正確
3. 標注課程範疇（數與代數 / 度量 / 圖形與空間 / 數據處理）
4. 描述錯誤（如有），每題分值若題目未標示則估算1分
5. **假分數必須直接給滿分**：假分數（如 7/4）與帶分數（如 1 3/4）是完全等價的正確答案，只要數值相等就直接給滿分，is_correct 設為 true，不可扣分或標注任何錯誤

只輸出純JSON（不加markdown代碼塊）：
{out_schema}"""

        # ── Single batch (fast path) ───────────────────────────────────
        if len(images) <= VISION_BATCH_SIZE:
            result = self._parse_json(self._vision(prompt, images, system_msg))
            if isinstance(result, dict):
                result["student_name"] = student_name  # always set
            return result

        # ── Multi-batch: each student page range scanned separately ──────
        batches = get_image_batches(images, VISION_BATCH_SIZE)
        all_q_results: List[Dict] = []
        total_awarded: float = 0.0
        total_possible: float = 0.0

        for i, batch in enumerate(batches):
            start_p = i * VISION_BATCH_SIZE + 1
            end_p = min(start_p + len(batch) - 1, n_pages)
            batch_prompt = (
                prompt
                + f"\n\n【你正在分析第 {start_p}\u2013{end_p} 頁，共 {n_pages} 頁，"
                + "本批次只包含此頁範圍內的題目，請勿遺漏任何一題】"
            )
            raw = self._vision(batch_prompt, batch, system_msg)
            parsed = self._parse_json(raw)
            if not parsed.get("parse_error"):
                all_q_results.extend(parsed.get("question_results", []))
                try:
                    total_awarded += float(parsed.get("total_marks_awarded") or 0)
                    total_possible += float(parsed.get("total_marks_possible") or 0)
                except (TypeError, ValueError):
                    pass

        # Deduplicate by question_ref (later batch wins if AI re-scans a page)
        seen_refs: Dict[str, Dict] = {}
        for q in all_q_results:
            seen_refs[str(q.get("question_ref", ""))] = q
        merged = list(seen_refs.values())

        if not total_possible:
            total_possible = sum(float(q.get("marks_possible") or 1) for q in merged)
        if not total_awarded:
            total_awarded = sum(float(q.get("marks_awarded") or 0) for q in merged)

        pct = round(100 * total_awarded / total_possible, 1) if total_possible else 0.0
        level = (
            "優秀(≥85%)" if pct >= 85
            else "良好(70-84%)" if pct >= 70
            else "一般(55-69%)" if pct >= 55
            else "需要改善(<55%)"
        )
        return {
            "student_name": student_name,
            "total_marks_awarded": total_awarded,
            "total_marks_possible": total_possible,
            "percentage": pct,
            "performance_level": level,
            "question_results": merged,
        }

    # ------------------------------------------------------------------
    # 5. Class-wide insight generation (text model, no vision)
    # ------------------------------------------------------------------

    def generate_class_insights(self, aggregated: Dict, grade: str) -> Dict:
        """
        Generate qualitative teaching insights from aggregated class stats.
        Uses the text model only (cheaper than vision).
        """
        curriculum_info = get_grade_curriculum(grade)
        weak_q = aggregated.get("weak_questions", [])
        strand_stats = aggregated.get("strand_stats", [])
        class_avg = aggregated.get("class_average", 0)
        total = aggregated.get("total_students", 0)
        dist = aggregated.get("class_distribution", {})

        prompt = f"""你是一位資深香港小學數學科主任，正在分析 {grade} 年級全班 {total} 位學生的試卷評分結果。

【全班整體數據】
全班平均分：{class_avg:.1f}%
成績分佈：{json.dumps(dist, ensure_ascii=False)}

【弱題排行（正確率最低）】
{json.dumps(weak_q[:10], ensure_ascii=False, indent=2)}

【各課程範疇分析】
{json.dumps(strand_stats, ensure_ascii=False, indent=2)}

香港 {grade} 年級數學課程綱要：
{curriculum_info}

請根據以上數據生成深度診斷，必須：
1. 找出全班最弱的2-3個課程範疇（數據佐證）
2. 分析常見錯誤類型（概念性誤解 vs 程序性錯誤）
3. 結合《數學課程指引》（2017）提出3個具體補救教學建議
4. 描述需要額外關注的學生群組特徵（不提名）

只輸出純JSON（不加markdown代碼塊）：
{{
  "overall_diagnosis": "全班學習問題核心診斷（3-5句，包含具體數據）",
  "weak_strand_analysis": [
    {{
      "strand": "課程範疇",
      "class_average_rate": 正確率,
      "key_issues": ["主要問題1", "問題2"],
      "misconception": "可能的概念誤解",
      "curriculum_link": "對應《數學課程指引》（2017）的學習目標"
    }}
  ],
  "error_type_analysis": {{
    "conceptual": "概念性誤解描述（附題號和正確率）",
    "procedural": "程序性錯誤描述（附題號和正確率）"
  }},
  "teaching_recommendations": [
    {{
      "priority": "高 / 中 / 低",
      "strand": "針對範疇",
      "strategy": "具體教學策略",
      "activities": ["教學活動1", "活動2"],
      "timeline": "建議時間（如：1週內）"
    }}
  ],
  "attention_students_note": "需要個別關注的學生群組特徵",
  "positive_findings": "全班優秀表現和可鼓勵之處"
}}"""

        raw = self._text(
            prompt,
            "你是資深香港小學數學科主任，熟悉香港課程發展議會《數學課程指引》（2017），"
            "擅長從全班數據中找出教學缺口並提出具體改善方案。",
        )
        return self._parse_json(raw)

    # ------------------------------------------------------------------
    # 7. Generate practice questions targeting a student's weak areas
    # ------------------------------------------------------------------

    def generate_practice_questions(
        self,
        student_name: str,
        grade: str,
        weak_questions: List[Dict],
        num_questions: int = 5,
        difficulty: str = "適中",
    ) -> Dict:
        """
        Generate practice questions that target a specific student's
        weak topics / error types, based on their wrong answers.

        Parameters
        ----------
        student_name : str
            Student display name.
        grade : str
            e.g. "P4"
        weak_questions : list[dict]
            Each dict should have: question_ref, topic, strand,
            correct_answer, student_answer, error_type, error_description,
            marks_possible.
        num_questions : int
            How many new practice questions to generate.
        difficulty : str
            "簡單" / "適中" / "進階"
        """
        curriculum_info = get_grade_curriculum(grade)

        weak_summary = json.dumps(weak_questions, ensure_ascii=False, indent=2)

        prompt = f"""你是一位經驗豐富的香港小學數學科老師，正在為 {grade} 年級的學生 {student_name} 設計針對性練習題。

以下是該學生在測驗中答錯的題目，包括錯誤類型和具體錯誤描述：

{weak_summary}

香港 {grade} 年級數學課程綱要：
{curriculum_info}

請根據該學生的弱點，生成 {num_questions} 道針對性練習題，要求：
1. 每道題必須針對該學生的某個特定弱點或錯誤類型
2. 題目難度：{difficulty}（在該年級範圍內）
3. 題目類型需多樣化（計算題、應用題、填充題等）
4. 每道題附有詳細解題步驟和答案
5. 每道題說明針對哪個弱點，以及為什麼這道題能幫助學生改善
6. 題目要貼近香港小學數學課程和日常生活情境

只輸出純JSON（不加markdown代碼塊）：
{{
  "student_name": "{student_name}",
  "grade": "{grade}",
  "weakness_summary": "該學生主要弱點概述（2-3句）",
  "practice_questions": [
    {{
      "question_number": 1,
      "targeted_weakness": "針對的弱點（對應原錯題）",
      "strand": "課程範疇",
      "topic": "具體主題",
      "question_type": "題目類型（計算題/應用題/填充題/選擇題）",
      "question_text": "完整題目文字（可包含多個小題）",
      "hints": "給學生的提示（選填）",
      "solution_steps": ["步驟1", "步驟2", "步驟3"],
      "answer": "正確答案",
      "explanation": "為什麼這道題能幫助改善該弱點"
    }}
  ],
  "study_tips": ["學習建議1", "學習建議2", "學習建議3"]
}}"""

        raw = self._text(
            prompt,
            "你是資深香港小學數學教師，精通因材施教、針對學生弱點設計練習題。"
            "你熟悉香港課程發展議會《數學課程指引》（2017），"
            "能設計符合課程要求且貼近學生生活的數學練習題。",
        )
        return self._parse_json(raw)

    def _text(self, prompt: str, system_msg: str = "") -> str:
        messages = []
        if system_msg:
            messages.append({"role": "system", "content": system_msg})
        messages.append({"role": "user", "content": prompt})

        resp = self.client.chat.completions.create(
            model=TEXT_MODEL,
            messages=messages,
            temperature=0.2,
            max_tokens=8192,
        )
        return resp.choices[0].message.content

    def _vision(self, prompt: str, images_b64: List[str], system_msg: str = "") -> str:
        """Single vision API call — one batch of pages."""
        messages = []
        if system_msg:
            messages.append({"role": "system", "content": system_msg})

        content = []
        for b64 in images_b64:
            # Auto-detect: JPEG starts with /9j in base64, PNG with iVBOR
            mime = "image/jpeg" if b64.startswith("/9j") else "image/png"
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:{mime};base64,{b64}"},
            })
        content.append({"type": "text", "text": prompt})
        messages.append({"role": "user", "content": content})

        resp = self.client.chat.completions.create(
            model=VISION_MODEL,
            messages=messages,
            temperature=0.2,
            max_tokens=4096,
        )
        return resp.choices[0].message.content

    def _vision_batched(
        self,
        base_prompt: str,
        images_b64: List[str],
        system_msg: str,
        merge_instruction: str,
    ) -> str:
        """
        Process a large set of images in batches of VISION_BATCH_SIZE,
        then synthesise all batch results into a single response via the
        text model.
        """
        if len(images_b64) <= VISION_BATCH_SIZE:
            return self._vision(base_prompt, images_b64, system_msg)

        batches = get_image_batches(images_b64, VISION_BATCH_SIZE)
        total = len(images_b64)
        batch_outputs = []

        for i, batch in enumerate(batches):
            start_p = i * VISION_BATCH_SIZE + 1
            end_p = min(start_p + len(batch) - 1, total)
            page_note = f"\n\n【你正在分析第 {start_p}–{end_p} 頁，共 {total} 頁】"
            raw = self._vision(base_prompt + page_note, batch, system_msg)
            batch_outputs.append(f"=== 第 {start_p}–{end_p} 頁分析結果 ===\n{raw}")

        combined = "\n\n".join(batch_outputs)
        synth_prompt = merge_instruction.replace("{batch_results}", combined)
        return self._text(synth_prompt, system_msg)

    # ------------------------------------------------------------------
    # JSON parsing
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_json(text: str) -> Dict:
        """Extract JSON from model response, tolerating extra prose."""
        # 1. Direct parse
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # 2. Extract from markdown code blocks
        m = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass

        # 3. Find outermost { … }
        m = re.search(r"(\{[\s\S]*\})", text)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass

        # 4. Return raw text so UI can still display something
        return {"raw_response": text, "parse_error": True}


# ---------------------------------------------------------------------------
# Pure-Python aggregation — no API call
# ---------------------------------------------------------------------------

def aggregate_student_results(
    all_student_results: List[Dict],
    expected_questions: Optional[List[str]] = None,
) -> Dict:
    """
    Aggregate per-student results into class-wide statistics.

    Parameters
    ----------
    all_student_results : list of per-student dicts (including failed ones)
    expected_questions  : ordered list of question_refs from the answer key.
                          When provided, question_stats preserves that order
                          and correct_rate denominator = n_valid.
    """
    n_total = len(all_student_results)
    valid = [
        r for r in all_student_results
        if not r.get("parse_error") and r.get("question_results")
    ]
    n_valid = len(valid)

    if not valid:
        return {"error": "沒有有效的學生分析結果"}

    percentages = [
        float(r["percentage"]) for r in valid
        if r.get("percentage") is not None
    ]
    class_avg = sum(percentages) / len(percentages) if percentages else 0.0

    dist = {"優秀(≥85%)": 0, "良好(70-84%)": 0, "一般(55-69%)": 0, "需要改善(<55%)": 0}
    for p in percentages:
        if p >= 85:
            dist["優秀(≥85%)"] += 1
        elif p >= 70:
            dist["良好(70-84%)"] += 1
        elif p >= 55:
            dist["一般(55-69%)"] += 1
        else:
            dist["需要改善(<55%)"] += 1

    # ── Per-question stats ──────────────────────────────────────────────
    # Pre-seed question_data using expected_questions so order is preserved
    question_data: Dict[str, dict] = {}
    if expected_questions:
        for ref in expected_questions:
            question_data[ref] = {
                "topic": "", "strand": "",
                "marks_possible": 1,
                "correct": 0, "marks_awarded": [], "errors": [],
            }

    for student in valid:
        for q in student.get("question_results", []):
            ref = str(q.get("question_ref", "")).strip()
            if not ref:
                continue
            if ref not in question_data:
                question_data[ref] = {
                    "topic": q.get("topic", ""),
                    "strand": q.get("strand", ""),
                    "marks_possible": q.get("marks_possible") or 1,
                    "correct": 0, "marks_awarded": [], "errors": [],
                }
            else:
                # Back-fill topic/strand from first student who answered it
                if not question_data[ref]["topic"]:
                    question_data[ref]["topic"] = q.get("topic", "")
                if not question_data[ref]["strand"]:
                    question_data[ref]["strand"] = q.get("strand", "")
            d = question_data[ref]
            if q.get("is_correct"):
                d["correct"] += 1
            ma = q.get("marks_awarded")
            if ma is not None:
                try:
                    d["marks_awarded"].append(float(ma))
                except (TypeError, ValueError):
                    pass
            err = q.get("error_description")
            if err and err not in ("null", None, ""):
                d["errors"].append(str(err))

    # Denominator = n_valid so that questions missed by some students' AI
    # are treated as wrong (not excluded), giving accurate class-wide rates.
    question_stats: List[Dict] = []
    for ref, d in question_data.items():
        correct_rate = round(100 * d["correct"] / n_valid) if n_valid else 0
        avg_marks = (
            round(sum(d["marks_awarded"]) / len(d["marks_awarded"]), 1)
            if d["marks_awarded"] else None
        )
        unique_errors: List[str] = []
        seen_err: set = set()
        for e in d["errors"]:
            k = e[:40]
            if k not in seen_err:
                seen_err.add(k)
                unique_errors.append(e)
                if len(unique_errors) >= 3:
                    break
        question_stats.append({
            "question_ref": ref,
            "topic": d["topic"],
            "strand": d["strand"],
            "marks_possible": d["marks_possible"],
            "class_correct_count": d["correct"],
            "class_correct_rate": correct_rate,
            "class_average_marks": avg_marks,
            "common_errors": unique_errors,
        })

    # Sort: schema order if provided, else natural sort by question_ref
    if expected_questions:
        order = {ref: i for i, ref in enumerate(expected_questions)}
        question_stats.sort(key=lambda x: order.get(x["question_ref"], 9999))
    else:
        question_stats.sort(key=lambda x: _natural_sort_key(x["question_ref"]))

    # ── Strand stats ────────────────────────────────────────────────────
    strand_data: Dict[str, dict] = {}
    for q in question_stats:
        s = (q.get("strand") or "其他").strip()
        if s not in strand_data:
            strand_data[s] = {"rates": [], "questions": []}
        strand_data[s]["rates"].append(q["class_correct_rate"])
        strand_data[s]["questions"].append(q["question_ref"])

    strand_stats = []
    for s, d in strand_data.items():
        avg_rate = round(sum(d["rates"]) / len(d["rates"])) if d["rates"] else 0
        strand_stats.append({
            "strand": s,
            "class_average_rate": avg_rate,
            "questions": d["questions"],
            "status": "弱項" if avg_rate < 60 else "一般" if avg_rate < 75 else "強項",
        })
    strand_stats.sort(key=lambda x: x["class_average_rate"])

    # ── Weak questions: separate sorted list (by rate asc) ───────────────
    weak_questions = sorted(
        [q for q in question_stats if q["class_correct_rate"] < 60],
        key=lambda x: x["class_correct_rate"],
    )
    for i, q in enumerate(weak_questions, 1):
        q["rank"] = i

    # ── Student ranking — ALL students (including failed papers) ─────────
    student_ranking = []
    for s in all_student_results:
        if s.get("parse_error"):
            student_ranking.append({
                "student_name": s.get("student_name", ""),
                "percentage": 0.0,
                "total_marks_awarded": "—",
                "total_marks_possible": "—",
                "performance_level": "分析失敗",
            })
        else:
            student_ranking.append({
                "student_name": s.get("student_name", ""),
                "percentage": float(s.get("percentage", 0)),
                "total_marks_awarded": s.get("total_marks_awarded", 0),
                "total_marks_possible": s.get("total_marks_possible", 0),
                "performance_level": s.get("performance_level", ""),
            })
    student_ranking.sort(key=lambda x: x["percentage"], reverse=True)
    for i, s in enumerate(student_ranking, 1):
        s["rank"] = i

    return {
        "total_students": n_total,      # total including failed papers
        "valid_students": n_valid,       # successfully analyzed
        "class_average": round(class_avg, 1),
        "class_distribution": dist,
        "student_results": all_student_results,  # ALL — for heatmap
        "question_stats": question_stats,        # sorted (schema order or natural)
        "strand_stats": strand_stats,
        "weak_questions": weak_questions,
        "student_ranking": student_ranking,      # ALL students
    }
