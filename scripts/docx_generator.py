"""
Samsung VD Marketing Expense Resolution (.docx) Generator
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    """표 셀 배경색 지정"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """표 셀 내부 여백 설정 (기존 품의서의 조밀하고 깔끔한 패딩 반영)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_memo_docx(doc_data, output_path="비용_메모품의서.docx"):
    doc = Document()
    
    # 1. 페이지 여백 설정 (상하 20mm, 좌우 20mm) - 기존 문서 스타일
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # 기본 스타일 폰트 지정 (바탕체 적용)
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Batang'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 2. 문서 제목 (진하게, 바탕체, 15pt)
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_after = Pt(14)
    run_title = title_p.add_run(doc_data.get('title', "[메모품의] 비용 집행의 件"))
    run_title.font.name = 'Batang'
    run_title.font.size = Pt(15)
    run_title.font.bold = True

    # 3. 문서 기본 정보 테이블 (기안부서 / 기안일자)
    info_table = doc.add_table(rows=1, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_headers = ["기안부서", doc_data.get('dept', "VD 마케팅팀"), "기안일자", doc_data.get('date', "2026.09.21")]
    hdr_cells = info_table.rows[0].cells
    for i, text in enumerate(info_headers):
        hdr_cells[i].text = text
        set_cell_margins(hdr_cells[i], top=80, bottom=80)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i % 2 == 0:
            set_cell_background(hdr_cells[i], "EEF2F8") # 표준 파란색 배경
            p.runs[0].font.bold = True
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 4. 1. 추진 배경 및 목적 (개조식)
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(10)
    h1.paragraph_format.space_after = Pt(4)
    h1_run = h1.add_run("1. 추진 배경 및 목적")
    h1_run.font.bold = True
    h1_run.font.size = Pt(11)

    for item in doc_data.get('background', []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.2)
        p.add_run(f"- {item}")

    # 5. 2. 소요 예산 (테이블 자동 계산 적용 - 크레딧 절감 핵심)
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    h2_run = h2.add_run(f"2. 소요 예산 ({doc_data.get('currency_unit', '단위: 천원, VAT 별도')})")
    h2_run.font.bold = True
    h2_run.font.size = Pt(11)

    cost_items = doc_data.get('cost_items', [])
    cols = ["구분", "전년 소요액(A)", "금년 소요액(B)", "증감(B-A)", "증감률(%)", "비고"]
    table = doc.add_table(rows=len(cost_items) + 1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 헤더 렌더링
    for i, col_name in enumerate(cols):
        cell = table.cell(0, i)
        cell.text = col_name
        set_cell_background(cell, "EEF2F8")
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True

    # 데이터 및 계산 오프라인 처리 (API 크레딧 절약)
    for r_idx, item in enumerate(cost_items, start=1):
        prev = item.get('prev_cost', 0)
        curr = item.get('curr_cost', 0)
        diff = curr - prev
        rate = ((diff / prev) * 100) if prev != 0 else 0.0

        row_data = [
            item.get('name', ''),
            f"{prev:,}",
            f"{curr:,}",
            f"{diff:+,}",
            f"{rate:+.1f}%",
            item.get('note', '')
        ]
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            cell.text = val
            set_cell_margins(cell)
            p = cell.paragraphs[0]
            if c_idx in [1, 2, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 6. 3. 향후 일정
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(4)
    h3_run = h3.add_run("3. 추진 일정")
    h3_run.font.bold = True
    h3_run.font.size = Pt(11)

    for schedule in doc_data.get('schedules', []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.2)
        p.add_run(f"- {schedule}")

    doc.save(output_path)
    return output_path
