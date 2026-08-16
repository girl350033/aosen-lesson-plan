import streamlit as st
import calendar
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# --- 頁面配置 ---
st.set_page_config(page_title="澳森托嬰中心 月教案輪值排班系統", layout="wide", initial_sidebar_state="collapsed")

# --- Word 排版樣式輔助函式 ---
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=80, bottom=80, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="1F497D", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def generate_docx(branch_name, year_roc, month, teachers_str, table_data):
    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(11.69)  # 橫向 A4
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # 1. 大標題 (分園名稱 + 年度月份)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    r_title = title_p.add_run(f"{branch_name} {year_roc} 年 {month} 月教案輪值表")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    # 2. 副標題 (領域說明 + 主帶老師) - 字體 12
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(8)
    r_sub = sub_p.add_run(f"幼兒發展領域：身體動作、社會情緒、語言溝通、認知探索、生活自理\n主帶托育人員：{teachers_str}")
    r_sub.font.name = "微軟正黑體"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 3. 建立輪值表格 (6 欄)
    table = doc.add_table(rows=len(table_data) + 1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="1F497D", sz="4")

    col_widths = [Inches(1.5), Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8)]
    headers = ["主題", "星期一：繪本", "星期二：小肌肉／認知", "星期三：體能課", "星期四：小肌肉／觸覺", "星期五：藝術創作"]

    # 表頭格式化 - 字體 12
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "1F497D")
        set_cell_margins(cell, top=80, bottom=80, left=60, right=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "微軟正黑體"
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.size = Pt(12)

    # 填入內容 - 嚴格設定微軟正黑體 12pt
    for row_idx, row_content in enumerate(table_data, start=1):
        for col_idx, text in enumerate(row_content):
            cell = table.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            set_cell_margins(cell, top=70, bottom=70, left=60, right=60)
            p = cell.paragraphs[0]
            
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_background(cell, "DCE6F1")
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_cell_background(cell, "FAFAFA" if row_idx % 2 == 1 else "FFFFFF")

            r = p.add_run(text)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(12)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            elif any(k in text for k in ["特約醫師", "教案暫停", "體能暫停", "國定連假", "中秋節", "教師節"]):
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- Web 介面佈局 ---
st.title("🏫 托嬰中心月教案輪值排班系統")
st.markdown("主任可在此選擇分園、設定月份與主帶老師名單，系統將自動產生當月工作日（週一至週五）排班表，並支援**一鍵生成 12pt Word 檔**。")

# 1. 基本設定區
st.subheader("壹、 基本設定")
c1, c2, c3 = st.columns([1.5, 1, 1])
with c1:
    branch = st.selectbox("園所分園：", ["澳森", "澳森文德"])
with c2:
    year_roc = st.number_input("民國年份：", min_value=114, max_value=125, value=115)
with c3:
    month = st.selectbox("月份：", list(range(1, 13)), index=8)  # 預設 9 月

default_teachers = "綺綺, Panda, 秋馨, Candy, 均宜, 樺樺, 小安, 悅熏, 知容, 政勳, 閔興"
teachers_input = st.text_input("主帶老師名單（以逗號分隔）：", default_teachers)

# 整理下拉選項：主任 + 主帶老師名單 + 特殊狀態
raw_teachers = [t.strip() for t in teachers_input.split(",") if t.strip()]
dropdown_options = ["主任"] + raw_teachers + ["教案暫停", "特約醫師", "國定連假"]

st.divider()

# 2. 月曆排班工作區（不含假日，僅抓週一至週五）
st.subheader(f"貳、 {branch} {year_roc} 年 {month} 月 課程輪值排班（週一至週五）")

year_ad = year_roc + 1911
cal = calendar.monthcalendar(year_ad, month)

# 篩選出含工作日（週一到週五）的週次
work_weeks = []
for w in cal:
    mon_to_fri = w[0:5]
    if any(d > 0 for d in mon_to_fri):
        work_weeks.append(mon_to_fri)

table_data = []
default_themes = ["顏色", "長大", "長大", "中秋節", "教師節"]
default_books = ["小藍與小黃", "顏色妖怪", "爸爸，我要月亮", "小星的大月餅", "老師，我們愛你"]

for idx, w in enumerate(work_weeks, start=1):
    with st.expander(f"📌 第 {idx} 週排班設定", expanded=True):
        col_theme, col_m, col_t, col_w, col_th, col_f = st.columns([1.2, 2.2, 1.3, 1.3, 1.3, 1.3])
        
        d_mon = f"{month}/{w[0]}" if w[0] > 0 else ""
        d_tue = f"{month}/{w[1]}" if w[1] > 0 else ""
        d_wed = f"{month}/{w[2]}" if w[2] > 0 else ""
        d_thu = f"{month}/{w[3]}" if w[3] > 0 else ""
        d_fri = f"{month}/{w[4]}" if w[4] > 0 else ""

        with col_theme:
            th_val = default_themes[idx-1] if idx <= len(default_themes) else "生活常規"
            theme = st.text_input(f"週主題", key=f"th_{idx}", value=th_val)

        with col_m:
            st.markdown(f"**週一：繪本 ({d_mon})**")
            m_lead = st.selectbox("導讀人", dropdown_options, key=f"ml_{idx}", index=0)
            bk_val = default_books[idx-1] if idx <= len(default_books) else "主題繪本導讀"
            book_name = st.text_input("繪本名稱", key=f"bn_{idx}", value=bk_val)
            mon_str = f"{d_mon} {m_lead}\n{book_name}" if d_mon else ""

        with col_t:
            st.markdown(f"**週二：認知 ({d_tue})**")
            tue_teacher = st.selectbox("負責老師", dropdown_options, key=f"t_{idx}", index=min(idx, len(dropdown_options)-1))
            tue_str = f"{d_tue} {tue_teacher}" if d_tue else ""

        with col_w:
            st.markdown(f"**週三：體能 ({d_wed})**")
            wed_teacher = st.selectbox("負責老師", dropdown_options, key=f"w_{idx}", index=min(3, len(dropdown_options)-1))
            wed_str = f"{d_wed} {wed_teacher}" if d_wed else ""

        with col_th:
            st.markdown(f"**週四：觸覺 ({d_thu})**")
            thu_teacher = st.selectbox("負責老師", dropdown_options, key=f"th_t_{idx}", index=min(idx+2, len(dropdown_options)-1))
            thu_str = f"{d_thu} {thu_teacher}" if d_thu else ""

        with col_f:
            st.markdown(f"**週五：藝術 ({d_fri})**")
            fri_teacher = st.selectbox("負責老師", dropdown_options, key=f"f_{idx}", index=min(idx+1, len(dropdown_options)-1))
            fri_str = f"{d_fri} {fri_teacher}" if d_fri else ""

        table_data.append([theme, mon_str, tue_str, wed_str, thu_str, fri_str])

# 3. 成果匯出區
st.divider()
st.subheader("參、 成果匯出")

doc_bytes = generate_docx(branch, year_roc, month, teachers_input, table_data)

st.download_button(
    label=f"📥 一鍵生成並下載【{branch} {year_roc}年{month}月教案輪值表】(Word 檔 / 12pt)",
    data=doc_bytes,
    file_name=f"{branch}_{year_roc}年{month}月教案輪值表.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
