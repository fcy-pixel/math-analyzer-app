"""Integration test for student batch analysis features."""
import sys

# ── Import checks ────────────────────────────────────────────────────────────
try:
    from file_processor import split_student_papers, get_pdf_page_count
    from analyzer import aggregate_student_results
    from pdf_exporter import build_student_report
    print("imports: OK")
except ImportError as e:
    print(f"IMPORT ERROR: {e}")
    sys.exit(1)

# ── aggregate_student_results ────────────────────────────────────────────────
mock_students = [
    {
        "student_name": "陳大文",
        "percentage": 85,
        "total_marks_awarded": 68,
        "total_marks_possible": 80,
        "performance_level": "良好(70-84%)",
        "question_results": [
            {"question_ref": "Q1", "topic": "分數", "strand": "數與代數",
             "marks_possible": 2, "marks_awarded": 2, "is_correct": True, "error_description": None},
            {"question_ref": "Q2", "topic": "面積", "strand": "度量",
             "marks_possible": 2, "marks_awarded": 0, "is_correct": False, "error_description": "忘記乘以2"},
            {"question_ref": "Q3", "topic": "四則運算", "strand": "數與代數",
             "marks_possible": 2, "marks_awarded": 2, "is_correct": True, "error_description": None},
        ],
    },
    {
        "student_name": "李小明",
        "percentage": 50,
        "total_marks_awarded": 40,
        "total_marks_possible": 80,
        "performance_level": "需要改善(<55%)",
        "question_results": [
            {"question_ref": "Q1", "topic": "分數", "strand": "數與代數",
             "marks_possible": 2, "marks_awarded": 0, "is_correct": False, "error_description": "直接加分母"},
            {"question_ref": "Q2", "topic": "面積", "strand": "度量",
             "marks_possible": 2, "marks_awarded": 2, "is_correct": True, "error_description": None},
            {"question_ref": "Q3", "topic": "四則運算", "strand": "數與代數",
             "marks_possible": 2, "marks_awarded": 0, "is_correct": False, "error_description": "計算次序錯誤"},
        ],
    },
    {
        "student_name": "黃美玲",
        "percentage": 75,
        "total_marks_awarded": 60,
        "total_marks_possible": 80,
        "performance_level": "良好(70-84%)",
        "question_results": [
            {"question_ref": "Q1", "topic": "分數", "strand": "數與代數",
             "marks_possible": 2, "marks_awarded": 2, "is_correct": True, "error_description": None},
            {"question_ref": "Q2", "topic": "面積", "strand": "度量",
             "marks_possible": 2, "marks_awarded": 0, "is_correct": False, "error_description": "公式錯誤"},
            {"question_ref": "Q3", "topic": "四則運算", "strand": "數與代數",
             "marks_possible": 2, "marks_awarded": 2, "is_correct": True, "error_description": None},
        ],
    },
]

agg = aggregate_student_results(mock_students)
assert agg["total_students"] == 3, f"Expected 3 students, got {agg['total_students']}"
avg = agg["class_average"]
assert 68 <= avg <= 72, f"avg={avg}"
assert len(agg["question_stats"]) == 3, f"Expected 3 question_stats, got {len(agg['question_stats'])}"
print(f"aggregate_student_results: OK — {agg['total_students']} students, avg={agg['class_average']}%")
print(f"  distribution: {agg['class_distribution']}")
print(f"  weak_questions: {[q['question_ref'] for q in agg['weak_questions']]}")
print(f"  strand_stats: {[(s['strand'], s['class_average_rate']) for s in agg['strand_stats']]}")

# ── build_student_report ─────────────────────────────────────────────────────
mock_insights = {
    "overall_diagnosis": "全班在分數加減法上普遍薄弱，Q1和Q3正確率分別為67%和67%。",
    "weak_strand_analysis": [
        {
            "strand": "度量",
            "class_average_rate": 33,
            "key_issues": ["面積公式應用"],
            "misconception": "混淆周長和面積計算方法",
            "curriculum_link": "P4 基本圖形面積",
        }
    ],
    "error_type_analysis": {
        "conceptual": "部分學生不理解通分概念，Q1有33%學生直接加分母",
        "procedural": "Q3計算次序錯誤，屬程序性錯誤",
    },
    "teaching_recommendations": [
        {
            "priority": "高",
            "strand": "度量",
            "strategy": "重教長方形和正方形面積公式，配合具體圖形操作",
            "activities": ["用方格紙計算面積", "小組練習"],
            "timeline": "1週內",
        }
    ],
    "attention_students_note": "約33%學生需要額外輔導，以分數運算為主要困難",
    "positive_findings": "所有學生均能完成試卷，積極嘗試每一題",
}

pdf_bytes = build_student_report(agg, mock_insights, "P4", "2024-25 上學期")
assert len(pdf_bytes) > 50000, f"PDF too small: {len(pdf_bytes)} bytes"
print(f"build_student_report: OK — {len(pdf_bytes):,} bytes")

with open("/tmp/test_student_report.pdf", "wb") as f:
    f.write(pdf_bytes)
print("Saved to /tmp/test_student_report.pdf")

print("\n✅ All tests passed!")
