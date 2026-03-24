"""
學生 QR Code 批量生成器
Batch Student QR Code Generator — A4 Printable Sheets
"""

import io
import zipfile

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="學生 QR Code 生成器",
    page_icon="🔲",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Register CJK font (STSong-Light 支援中文)
# ---------------------------------------------------------------------------
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
.main-header {
    background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
    padding: 20px 28px;
    border-radius: 12px;
    color: white;
    margin-bottom: 24px;
}
.main-header h1 { margin: 0 0 6px 0; font-size: 1.7rem; }
.main-header p  { margin: 0; opacity: 0.85; font-size: 0.9rem; }
.tip-box {
    background: #e8f0fe;
    border-left: 4px solid #1a73e8;
    padding: 10px 14px;
    border-radius: 0 8px 8px 0;
    font-size: 0.88rem;
}
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
  <h1>🔲 學生 QR Code 批量生成器</h1>
  <p>上載 CSV / 手動輸入學生名稱，一鍵生成可貼在手冊上的 A4 打印版 QR Code</p>
</div>
""",
    unsafe_allow_html=True,
)

# ===========================================================================
# Helpers
# ===========================================================================


def make_qr_image(data: str, box_size: int = 12) -> Image.Image:
    """Return a PIL RGBA QR code image."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    return img


def generate_pdf(
    students: list[dict],
    cols: int,
    rows: int,
    qr_mm: float,
    label_font_size: int,
    show_class: bool,
    show_cut_lines: bool,
    school_name: str,
) -> bytes:
    """
    Build an A4 PDF with QR code cards.

    Each entry in `students` is a dict with keys:
      - name       (str)
      - class_name (str, optional)
      - qr_data    (str)  — the actual content encoded in QR
    """
    buffer = io.BytesIO()
    page_w, page_h = A4  # points

    margin_x = 10 * mm
    margin_y = 10 * mm
    header_h = 8 * mm if school_name else 0

    usable_w = page_w - 2 * margin_x
    usable_h = page_h - 2 * margin_y - header_h

    card_w = usable_w / cols
    card_h = usable_h / rows

    # QR image size in points
    qr_pt = min(qr_mm * mm, card_w - 6 * mm, card_h - 18 * mm)
    cards_per_page = cols * rows

    c = canvas.Canvas(buffer, pagesize=A4)

    def draw_page_header(page_num: int, total_pages: int):
        if school_name:
            c.setFont("STSong-Light", 9)
            c.setFillColor(colors.HexColor("#555555"))
            c.drawString(margin_x, page_h - margin_y + 2 * mm, school_name)
            c.drawRightString(
                page_w - margin_x,
                page_h - margin_y + 2 * mm,
                f"第 {page_num} 頁 / 共 {total_pages} 頁",
            )

    total_pages = max(1, -(-len(students) // cards_per_page))  # ceiling division

    for idx, student in enumerate(students):
        page_idx = idx % cards_per_page
        if page_idx == 0 and idx > 0:
            c.showPage()
        current_page = idx // cards_per_page + 1
        if page_idx == 0:
            draw_page_header(current_page, total_pages)

        col = page_idx % cols
        row = page_idx // cols

        card_x = margin_x + col * card_w
        card_y = page_h - margin_y - header_h - (row + 1) * card_h

        # --- Card border / cut lines ---
        if show_cut_lines:
            c.setStrokeColor(colors.HexColor("#aaaaaa"))
            c.setLineWidth(0.4)
            c.setDash(3, 3)
        else:
            c.setStrokeColor(colors.HexColor("#cccccc"))
            c.setLineWidth(0.5)
            c.setDash()

        padding = 1.5 * mm
        c.rect(
            card_x + padding,
            card_y + padding,
            card_w - 2 * padding,
            card_h - 2 * padding,
        )
        c.setDash()

        # --- QR code image ---
        qr_img = make_qr_image(student["qr_data"])
        img_buf = io.BytesIO()
        qr_img.save(img_buf, format="PNG")
        img_buf.seek(0)

        name_area_h = label_font_size * 1.4 + (label_font_size * 1.2 if show_class and student.get("class_name") else 0)
        name_area_h = (name_area_h / 72) * 25.4 * mm  # convert pt→mm→points

        qr_x = card_x + (card_w - qr_pt) / 2
        qr_y = card_y + name_area_h + 3 * mm

        c.drawImage(ImageReader(img_buf), qr_x, qr_y, width=qr_pt, height=qr_pt)

        # --- Labels ---
        centre_x = card_x + card_w / 2
        text_y = card_y + 3 * mm

        c.setFillColor(colors.black)

        if show_class and student.get("class_name"):
            c.setFont("STSong-Light", label_font_size - 1)
            c.drawCentredString(centre_x, text_y, student["class_name"])
            text_y += (label_font_size + 2) * 0.352778 * mm  # pt → mm → points (approx)

        c.setFont("STSong-Light", label_font_size)
        c.drawCentredString(centre_x, text_y, student["name"])

    # Last page header if only one page of students
    if len(students) <= cards_per_page:
        draw_page_header(1, total_pages)

    c.save()
    buffer.seek(0)
    return buffer.read()


def generate_zip(students: list[dict]) -> bytes:
    """Return a ZIP of individual QR PNG images."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for s in students:
            img = make_qr_image(s["qr_data"])
            img_buf = io.BytesIO()
            img.save(img_buf, format="PNG")
            safe_name = s["name"].replace("/", "_").replace("\\", "_")
            zf.writestr(f"{safe_name}.png", img_buf.getvalue())
    buf.seek(0)
    return buf.read()


# ===========================================================================
# Sidebar — Settings
# ===========================================================================
with st.sidebar:
    st.header("⚙️ 設定")

    st.subheader("版面排列")
    layout_choice = st.selectbox(
        "每頁排列方式",
        options=["3 × 4（每頁 12 個）", "4 × 5（每頁 20 個）", "2 × 3（每頁 6 個，較大）"],
        index=0,
    )
    layout_map = {
        "3 × 4（每頁 12 個）": (3, 4),
        "4 × 5（每頁 20 個）": (4, 5),
        "2 × 3（每頁 6 個，較大）": (2, 3),
    }
    cols, rows = layout_map[layout_choice]

    qr_size_mm = st.slider("QR Code 大小（mm）", min_value=30, max_value=60, value=45, step=5)
    label_font_size = st.slider("名字字型大小（pt）", min_value=8, max_value=16, value=11, step=1)

    st.subheader("內容選項")
    qr_content_mode = st.radio(
        "QR Code 內容",
        options=["只有名字", "名字 + 班別", "學號（需要 CSV 有學號欄）"],
        index=0,
    )
    show_class_label = st.checkbox("在 QR 下方顯示班別", value=True)
    show_cut_lines = st.checkbox("顯示裁剪虛線", value=True)

    st.subheader("頁首（選填）")
    school_name = st.text_input("學校名稱（印在每頁頁首）", value="")

# ===========================================================================
# Main — Input method
# ===========================================================================
tab_csv, tab_manual = st.tabs(["📂 上載 CSV", "✏️ 手動輸入"])

students_raw: list[dict] = []

# ---------------------------------------------------------------------------
# Tab 1 — CSV upload
# ---------------------------------------------------------------------------
with tab_csv:
    st.markdown(
        """
<div class="tip-box">
💡 <b>CSV 格式說明：</b> 檔案須包含 <code>姓名</code>（或 <code>Name</code>）欄位。<br>
可選欄位：<code>班別</code>（或 <code>Class</code>）、<code>學號</code>（或 <code>ID</code>）。<br>
第一行為表頭。
</div>
""",
        unsafe_allow_html=True,
    )
    st.download_button(
        "⬇️ 下載 CSV 範本",
        data="姓名,班別,學號\n陳大文,1A,001\n李小明,1A,002\n黃美玲,1B,003\n",
        file_name="student_template.csv",
        mime="text/csv",
    )

    uploaded_file = st.file_uploader("上載 CSV 檔案", type=["csv"])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            # Normalise column names
            col_map = {}
            for col in df.columns:
                lower = col.strip().lower()
                if lower in ("姓名", "name", "名字", "student name", "學生姓名"):
                    col_map[col] = "name"
                elif lower in ("班別", "class", "班級", "class name"):
                    col_map[col] = "class_name"
                elif lower in ("學號", "id", "student id", "學生編號", "編號"):
                    col_map[col] = "student_id"
            df = df.rename(columns=col_map)

            if "name" not in df.columns:
                st.error("❌ 找不到姓名欄位，請確保 CSV 包含「姓名」或「Name」欄。")
            else:
                df = df.dropna(subset=["name"])
                df["name"] = df["name"].astype(str).str.strip()
                df = df[df["name"] != ""]
                st.success(f"✅ 成功讀取 **{len(df)}** 名學生")
                st.dataframe(df, use_container_width=True, height=220)

                for _, row in df.iterrows():
                    entry = {
                        "name": row["name"],
                        "class_name": str(row.get("class_name", "")).strip() if "class_name" in df.columns else "",
                        "student_id": str(row.get("student_id", "")).strip() if "student_id" in df.columns else "",
                    }
                    if qr_content_mode == "只有名字":
                        entry["qr_data"] = entry["name"]
                    elif qr_content_mode == "名字 + 班別":
                        parts = [entry["name"]]
                        if entry["class_name"]:
                            parts.append(entry["class_name"])
                        entry["qr_data"] = " | ".join(parts)
                    else:  # 學號
                        entry["qr_data"] = entry["student_id"] or entry["name"]
                    students_raw.append(entry)
        except Exception as e:
            st.error(f"讀取 CSV 時發生錯誤：{e}")

# ---------------------------------------------------------------------------
# Tab 2 — Manual input
# ---------------------------------------------------------------------------
with tab_manual:
    st.info("每行輸入一位學生，格式：**姓名** 或 **姓名,班別**（例如：陳大文,1A）")
    manual_text = st.text_area(
        "學生名單",
        placeholder="陳大文,1A\n李小明,1B\n黃美玲,2A",
        height=250,
    )
    if manual_text.strip():
        for line in manual_text.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            parts = [p.strip() for p in line.split(",")]
            name = parts[0]
            class_name = parts[1] if len(parts) > 1 else ""
            if name:
                if qr_content_mode == "名字 + 班別" and class_name:
                    qr_data = f"{name} | {class_name}"
                else:
                    qr_data = name
                students_raw.append(
                    {"name": name, "class_name": class_name, "student_id": "", "qr_data": qr_data}
                )
        if students_raw:
            st.success(f"✅ 已輸入 **{len(students_raw)}** 名學生")

# ===========================================================================
# Preview & Generate
# ===========================================================================
if students_raw:
    st.divider()
    st.subheader(f"👁️ 預覽（共 {len(students_raw)} 名學生）")

    total_pages = max(1, -(-len(students_raw) // (cols * rows)))
    st.caption(
        f"將生成 **{total_pages}** 頁 A4（每頁 {cols}×{rows}={cols*rows} 個），"
        f"共 {len(students_raw)} 個 QR Code"
    )

    # Show small preview grid of first few QR codes
    preview_n = min(len(students_raw), cols * 2)
    preview_cols = st.columns(min(preview_n, cols))
    for i, student in enumerate(students_raw[:preview_n]):
        with preview_cols[i % cols]:
            img = make_qr_image(student["qr_data"], box_size=6)
            st.image(img, caption=student["name"], width=130)

    st.divider()
    col_a, col_b = st.columns(2)

    with col_a:
        if st.button("📄 生成 A4 PDF（可打印）", type="primary", use_container_width=True):
            with st.spinner("正在生成 PDF，請稍候…"):
                try:
                    pdf_bytes = generate_pdf(
                        students=students_raw,
                        cols=cols,
                        rows=rows,
                        qr_mm=qr_size_mm,
                        label_font_size=label_font_size,
                        show_class=show_class_label,
                        show_cut_lines=show_cut_lines,
                        school_name=school_name,
                    )
                    st.success(f"✅ PDF 已生成！共 {total_pages} 頁")
                    st.download_button(
                        label="⬇️ 下載 PDF",
                        data=pdf_bytes,
                        file_name="student_qr_codes.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"生成 PDF 時發生錯誤：{e}")
                    st.exception(e)

    with col_b:
        if st.button("🗜️ 下載個別 QR Code（ZIP）", use_container_width=True):
            with st.spinner("正在打包 ZIP…"):
                try:
                    zip_bytes = generate_zip(students_raw)
                    st.success("✅ ZIP 已生成！")
                    st.download_button(
                        label="⬇️ 下載 ZIP",
                        data=zip_bytes,
                        file_name="student_qr_codes.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"生成 ZIP 時發生錯誤：{e}")

else:
    st.info("👆 請先上載 CSV 檔案或在「手動輸入」分頁填寫學生名單。")

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.divider()
st.caption("打印提示：請選擇「實際大小」或「100%」縮放列印，切勿勾選「符合頁面」，以確保 QR Code 尺寸準確。")
