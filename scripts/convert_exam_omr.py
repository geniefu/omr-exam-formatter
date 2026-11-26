"""
Reusable OMR Exam Paper Formatter (Plan A: Page 1 Header OMR + Page 2 Footer OMR).
Stdlib-only Python script for transforming Taiwan B4/A4 two-column Word exam papers (.docx)
into B&W mimeograph-friendly, auto-gradable OMR exam papers.
"""

import argparse
import os
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

HF_NS_DECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)

NO_BORDERS = {"top": "none", "left": "none", "bottom": "none", "right": "none"}
GRID_BORDERS = {
    "top": ("single", 4, "000000"),
    "left": ("single", 4, "000000"),
    "bottom": ("single", 4, "000000"),
    "right": ("single", 4, "000000"),
}
GRID_HEADER_BORDERS = {
    "top": ("single", 10, "000000"),
    "left": ("single", 4, "000000"),
    "bottom": ("single", 10, "000000"),
    "right": ("single", 4, "000000"),
}
GRID_BOTTOM_BORDERS = {
    "top": ("single", 4, "000000"),
    "left": ("single", 4, "000000"),
    "bottom": ("single", 10, "000000"),
    "right": ("single", 4, "000000"),
}


def cell_xml(
    text: str,
    width: int,
    *,
    bold: bool = False,
    sz: int = 18,
    color: str = "000000",
    fill: str | None = None,
    font: str = "微軟正黑體",
    borders: dict | None = None,
    grid_span: int = 1,
    v_merge: str | None = None,
    align: str = "center",
    pad_lr: int = 20,
    pad_tb: int = 10,
) -> str:
    tcPr = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if grid_span > 1:
        tcPr.append(f'<w:gridSpan w:val="{grid_span}"/>')
    if v_merge == "restart":
        tcPr.append('<w:vMerge w:val="restart"/>')
    elif v_merge == "continue":
        tcPr.append('<w:vMerge/>')
    if borders is not None:
        b_parts = []
        for side in ("top", "left", "bottom", "right"):
            spec = borders.get(side)
            if spec is None or spec == "none":
                b_parts.append(f'<w:{side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
            else:
                val, bsz, bcol = spec
                b_parts.append(f'<w:{side} w:val="{val}" w:sz="{bsz}" w:space="0" w:color="{bcol}"/>')
        tcPr.append(f'<w:tcBorders>{"".join(b_parts)}</w:tcBorders>')
    if fill:
        tcPr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>')
    tcPr.append(
        f'<w:tcMar>'
        f'<w:top w:w="{pad_tb}" w:type="dxa"/>'
        f'<w:bottom w:w="{pad_tb}" w:type="dxa"/>'
        f'<w:left w:w="{pad_lr}" w:type="dxa"/>'
        f'<w:right w:w="{pad_lr}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append('<w:vAlign w:val="center"/>')

    rPr = [
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/>',
        f'<w:color w:val="{color}"/>',
        f'<w:sz w:val="{sz}"/>',
        f'<w:szCs w:val="{sz}"/>',
    ]
    if bold:
        rPr.append("<w:b/><w:bCs/>")

    run_xml = (
        f'<w:r><w:rPr>{"".join(rPr)}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
        if text
        else ""
    )
    p_xml = (
        f'<w:p>'
        f'<w:pPr>'
        f'<w:spacing w:before="0" w:after="0" w:line="220" w:lineRule="auto"/>'
        f'<w:jc w:val="{align}"/>'
        f'</w:pPr>'
        f'{run_xml}'
        f'</w:p>'
    )
    return f'<w:tc><w:tcPr>{"".join(tcPr)}</w:tcPr>{p_xml}</w:tc>'


def build_omr_table_xml(
    page_num: int,
    q_labels: list[str],
    include_seat: bool,
    total_w: int,
    max_seat: int = 30,
) -> str:
    """Pure B&W OMR bubble strip (no shading, crisp black lines, 4 corner fiducials).
    Seat layout adapts to max_seat (default 30 -> 01~30) using a unified 6-col structure (1 label + 5 digits):
      - max_seat <= 49 (default 30):
          十位 (rows 1-2 merged): 橫向由左至右 ⓪ ① ② ③ 空格 (for <=39) 或 ⓪ ① ② ③ ④ (for 40~49)
          個位 (rows 3-4):        第 3 列 ⓪ ① ② ③ ④、第 4 列 ⑤ ⑥ ⑦ ⑧ ⑨
      - max_seat >= 50:
          十位 (rows 1-2):        第 1 列 ⓪ ① ② ③ ④、第 2 列 ⑤ ⑥ ⑦ ⑧ ⑨
          個位 (rows 3-4):        第 3 列 ⓪ ① ② ③ ④、第 4 列 ⑤ ⑥ ⑦ ⑧ ⑨
    """
    fid_w = 240
    gap_w = 110

    col_widths = [fid_w, gap_w]
    seat_label_w = 360
    seat_digit_w = 350
    seat_gap_w = 140
    single_row_tens = max_seat <= 49

    if include_seat:
        col_widths.extend([seat_label_w] + [seat_digit_w] * 5 + [seat_gap_w])
        used = fid_w * 2 + gap_w * 2 + seat_label_w + seat_digit_w * 5 + seat_gap_w
    else:
        tag_w = 600
        col_widths.extend([tag_w, seat_gap_w])
        used = fid_w * 2 + gap_w * 2 + tag_w + seat_gap_w

    n_q = max(1, len(q_labels))
    rem_w = total_w - used
    q_col_w = rem_w // n_q
    q_widths = [q_col_w] * n_q
    q_widths[-1] += rem_w - q_col_w * n_q
    col_widths.extend(q_widths)
    col_widths.extend([gap_w, fid_w])

    grid_xml = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    row_heights = [280, 330, 330, 330, 330]
    bubbles = ["Ⓐ", "Ⓑ", "Ⓒ", "Ⓓ"]
    digits_top = ["⓪", "①", "②", "③", "④"]
    digits_bot = ["⑤", "⑥", "⑦", "⑧", "⑨"]
    max_tens = min(9, max_seat // 10)
    seat_range_str = f"01~{max_seat:02d}"

    rows_xml = []
    for r_idx in range(5):
        cells = []
        is_header = (r_idx == 0)
        is_bottom = (r_idx == 4)

        # Left fiducial column
        if is_header or is_bottom:
            cells.append(cell_xml("", fid_w, fill="000000", borders=NO_BORDERS, pad_lr=0, pad_tb=0))
        elif r_idx == 2:
            cells.append(cell_xml(f"P{page_num}", fid_w, bold=True, sz=15, borders=NO_BORDERS, pad_lr=0, pad_tb=0))
        else:
            cells.append(cell_xml("", fid_w, borders=NO_BORDERS, pad_lr=0, pad_tb=0))

        cells.append(cell_xml("", gap_w, borders=NO_BORDERS, pad_lr=0, pad_tb=0))

        if include_seat:
            if is_header:
                b = dict(GRID_HEADER_BORDERS)
                b["left"] = ("single", 10, "000000")
                b["right"] = ("single", 10, "000000")
                cells.append(
                    cell_xml(
                        f"座號畫卡 ({seat_range_str})",
                        seat_label_w + seat_digit_w * 5,
                        bold=True,
                        sz=16,
                        borders=b,
                        grid_span=6,
                    )
                )
            else:
                lb = dict(GRID_BORDERS)
                lb["left"] = ("single", 10, "000000")
                if r_idx == 1:
                    cells.append(cell_xml("十位", seat_label_w, bold=True, sz=15, borders=lb, v_merge="restart"))
                elif r_idx == 2:
                    lb["bottom"] = ("single", 10, "000000")
                    cells.append(cell_xml("", seat_label_w, borders=lb, v_merge="continue"))
                elif r_idx == 3:
                    cells.append(cell_xml("個位", seat_label_w, bold=True, sz=15, borders=lb, v_merge="restart"))
                elif r_idx == 4:
                    lb["bottom"] = ("single", 10, "000000")
                    cells.append(cell_xml("", seat_label_w, borders=lb, v_merge="continue"))

                if single_row_tens and r_idx in (1, 2):
                    # Horizontal tens row (merged rows 1-2): ⓪ ① ② ③ (blank) for <=39, or ⓪ ① ② ③ ④ for 40~49
                    for di in range(5):
                        db = dict(GRID_BOTTOM_BORDERS if r_idx == 2 else GRID_BORDERS)
                        if di == 4:
                            db["right"] = ("single", 10, "000000")
                        if r_idx == 1:
                            dch = digits_top[di] if di <= max_tens else ""
                            cells.append(
                                cell_xml(dch, seat_digit_w, sz=24, font="Segoe UI Symbol", borders=db, v_merge="restart", pad_lr=0, pad_tb=0)
                            )
                        else:
                            cells.append(
                                cell_xml("", seat_digit_w, borders=db, v_merge="continue", pad_lr=0, pad_tb=0)
                            )
                else:
                    digs = digits_top if r_idx in (1, 3) else digits_bot
                    for di, dch in enumerate(digs):
                        db = dict(GRID_BOTTOM_BORDERS if r_idx in (2, 4) else GRID_BORDERS)
                        if di == 4:
                            db["right"] = ("single", 10, "000000")
                        cells.append(
                            cell_xml(dch, seat_digit_w, sz=24, font="Segoe UI Symbol", borders=db, pad_lr=0, pad_tb=0)
                        )
            cells.append(cell_xml("", seat_gap_w, borders=NO_BORDERS, pad_lr=0, pad_tb=0))
        else:
            tag_w = 600
            tb = dict(GRID_HEADER_BORDERS if is_header else (GRID_BOTTOM_BORDERS if is_bottom else GRID_BORDERS))
            tb["left"] = ("single", 10, "000000")
            tb["right"] = ("single", 10, "000000")
            if is_header:
                cells.append(cell_xml("背面", tag_w, bold=True, sz=16, borders=tb))
            elif r_idx == 1:
                cells.append(cell_xml("第2面", tag_w, bold=True, sz=16, borders=tb, v_merge="restart"))
            else:
                cells.append(cell_xml("", tag_w, borders=tb, v_merge="continue"))
            cells.append(cell_xml("", seat_gap_w, borders=NO_BORDERS, pad_lr=0, pad_tb=0))

        for qi, (qlab, qw) in enumerate(zip(q_labels, q_widths)):
            qb = dict(GRID_HEADER_BORDERS if is_header else (GRID_BOTTOM_BORDERS if is_bottom else GRID_BORDERS))
            if qi == 0:
                qb["left"] = ("single", 10, "000000")
            if qi == n_q - 1:
                qb["right"] = ("single", 10, "000000")
            if is_header:
                cells.append(cell_xml(qlab, qw, bold=True, sz=18, borders=qb, pad_lr=10, pad_tb=10))
            else:
                cells.append(cell_xml(bubbles[r_idx - 1], qw, sz=26, font="Segoe UI Symbol", borders=qb, pad_lr=0, pad_tb=0))

        cells.append(cell_xml("", gap_w, borders=NO_BORDERS, pad_lr=0, pad_tb=0))

        if is_header or is_bottom:
            cells.append(cell_xml("", fid_w, fill="000000", borders=NO_BORDERS, pad_lr=0, pad_tb=0))
        elif r_idx == 2 and page_num == 2:
            cells.append(cell_xml("■", fid_w, sz=12, borders=NO_BORDERS, pad_lr=0, pad_tb=0))
        else:
            cells.append(cell_xml("", fid_w, borders=NO_BORDERS, pad_lr=0, pad_tb=0))

        rows_xml.append(
            f'<w:tr><w:trPr><w:trHeight w:val="{row_heights[r_idx]}" w:hRule="exact"/></w:trPr>{"".join(cells)}</w:tr>'
        )

    return (
        f'<w:tbl>'
        f'<w:tblPr><w:tblW w:w="{total_w}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        f'<w:tblBorders>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders></w:tblPr>'
        f'<w:tblGrid>{grid_xml}</w:tblGrid>{"".join(rows_xml)}</w:tbl>'
    )


def build_page1_header_xml(title: str, grade: str, q_labels: list[str], total_w: int, max_seat: int = 30) -> str:
    grade_part = f"　　{grade} 年 ______ 班" if grade else "　　____ 年 ______ 班"
    info_p = (
        '<w:p>'
        '<w:pPr><w:spacing w:before="0" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="微軟正黑體" w:eastAsia="微軟正黑體" w:hAnsi="微軟正黑體"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f'<w:t xml:space="preserve">{title}{grade_part}　座號：______　姓名：____________</w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="微軟正黑體" w:eastAsia="微軟正黑體" w:hAnsi="微軟正黑體"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>'
        '<w:t xml:space="preserve">　　【第 1 面畫卡區：請用 2B 鉛筆、黑筆或藍筆將座號與選項圓圈塗滿】</w:t></w:r>'
        '</w:p>'
    )
    tbl_xml = build_omr_table_xml(1, q_labels, include_seat=True, total_w=total_w, max_seat=max_seat)
    spacer_p = (
        '<w:p><w:pPr>'
        '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'
        '<w:spacing w:before="40" w:after="60" w:line="60" w:lineRule="exact"/>'
        '</w:pPr></w:p>'
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr {HF_NS_DECL}>{info_p}{tbl_xml}{spacer_p}</w:hdr>'


def build_page2_empty_header_xml() -> str:
    """Empty Page 2 header: no repeated title/class/seat/name (Page 1 already has them)."""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr {HF_NS_DECL}>'
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>'
        f'</w:hdr>'
    )


def build_page2_footer_xml(q_labels: list[str], total_w: int, note: str = "") -> str:
    note_txt = f"（{note}）" if note else ""
    top_sep = (
        '<w:p><w:pPr>'
        '<w:pBdr><w:top w:val="single" w:sz="6" w:space="2" w:color="000000"/></w:pBdr>'
        '<w:spacing w:before="40" w:after="50" w:line="220" w:lineRule="auto"/>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="微軟正黑體" w:eastAsia="微軟正黑體" w:hAnsi="微軟正黑體"/><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '<w:t xml:space="preserve">【第 2 面畫卡區】</w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="微軟正黑體" w:eastAsia="微軟正黑體" w:hAnsi="微軟正黑體"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>'
        f'<w:t xml:space="preserve">請核對題號，用 2B 鉛筆、黑筆或藍筆將選項圓圈塗滿{note_txt}</w:t></w:r>'
        '</w:p>'
    )
    tbl_xml = build_omr_table_xml(2, q_labels, include_seat=False, total_w=total_w)
    page_num_p = (
        '<w:p><w:pPr><w:spacing w:before="50" w:after="0" w:line="200" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '<w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>2</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr {HF_NS_DECL}>{top_sep}{tbl_xml}{page_num_p}</w:ftr>'


def inspect_docx(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        doc_root = ET.fromstring(z.read("word/document.xml"))
        sect = doc_root.find(".//w:sectPr", NS)
        pgSz = sect.find("w:pgSz", NS) if sect is not None else None
        pgMar = sect.find("w:pgMar", NS) if sect is not None else None
        if pgSz is not None:
            w = int(pgSz.attrib.get(f"{{{W_NS}}}w", 0))
            h = int(pgSz.attrib.get(f"{{{W_NS}}}h", 0))
            print(f"Page size: {w} x {h} dxa ({w/567:.2f} x {h/567:.2f} cm)")
        if pgMar is not None:
            m = {k.split("}")[-1]: int(v) for k, v in pgMar.attrib.items()}
            print("Margins (dxa):", m)
        for name in z.namelist():
            if "header" in name or "footer" in name:
                r = ET.fromstring(z.read(name))
                print(f"{name}: {''.join(r.itertext()).strip()}")
        print("\nParagraphs (index | paraId | numId | preview):")
        for idx, p in enumerate(doc_root.findall(".//w:body/w:p", NS)):
            txt = "".join(p.itertext()).strip()
            if not txt:
                continue
            pid = p.attrib.get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId", "-")
            numId = "-"
            n_el = p.find(".//w:numPr/w:numId", NS)
            if n_el is not None:
                numId = n_el.attrib.get(f"{{{W_NS}}}val", "-")
            print(f"  P{idx:02d} | {pid} | num={numId} | {txt[:70]}")


def export_pdf_word_com(docx_path: str, pdf_path: str) -> int:
    ps = (
        "$word = New-Object -ComObject Word.Application; "
        "$word.Visible = $false; "
        "try { "
        f"$doc = $word.Documents.Open('{os.path.abspath(docx_path)}', $false, $true); "
        "$pages = $doc.ComputeStatistics(2); "
        "Write-Host \"PAGES:$pages\"; "
        f"$doc.ExportAsFixedFormat('{os.path.abspath(pdf_path)}', 17); "
        "$doc.Close($false); "
        "} finally { $word.Quit() }"
    )
    res = subprocess.run(["powershell", "-Command", ps], capture_output=True, text=True, check=True)
    for line in res.stdout.splitlines():
        if line.startswith("PAGES:"):
            return int(line.split(":")[1].strip())
    return -1


def convert_exam(
    src_docx: str,
    out_docx: str,
    title: str,
    grade: str,
    p1_end: int,
    mc_count: int,
    total_q: int,
    max_seat: int = 30,
    break_para_id: str | None = None,
    mc_left_ind: int = 420,
    sub_left_ind: int = 780,
    sub_hanging: int = 300,
    export_pdf: bool = False,
) -> None:
    with zipfile.ZipFile(src_docx, "r") as zin:
        doc_xml_orig = zin.read("word/document.xml").decode("utf-8")
        rels_orig = zin.read("word/_rels/document.xml.rels").decode("utf-8")

        # Determine printable width from pgSz and pgMar
        doc_root = ET.fromstring(doc_xml_orig)
        sect = doc_root.find(".//w:sectPr", NS)
        pg_w = int(sect.find("w:pgSz", NS).attrib.get(f"{{{W_NS}}}w", "14572"))
        pg_mar = sect.find("w:pgMar", NS)
        m_left = int(pg_mar.attrib.get(f"{{{W_NS}}}left", "720"))
        m_right = int(pg_mar.attrib.get(f"{{{W_NS}}}right", "720"))
        total_w = pg_w - m_left - m_right

        # Find existing rIds to avoid collisions
        existing_rids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_orig)]
        next_rid_num = max(existing_rids, default=40) + 1

        # Locate first-page header, default header (if any), & default footer targets from sectPr
        hdr1_target = "word/header1.xml"
        hdr_default_target = None
        ftr_default_target = "word/footer1.xml"
        rid_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_orig))
        for ref in sect.findall("w:headerReference", NS):
            t = ref.attrib.get(f"{{{W_NS}}}type")
            rid = ref.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if t == "first" and rid in rid_map:
                hdr1_target = "word/" + rid_map[rid].lstrip("/")
            elif t == "default" and rid in rid_map:
                hdr_default_target = "word/" + rid_map[rid].lstrip("/")
        for ref in sect.findall("w:footerReference", NS):
            t = ref.attrib.get(f"{{{W_NS}}}type")
            rid = ref.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if t == "default" and rid in rid_map:
                ftr_default_target = "word/" + rid_map[rid].lstrip("/")

        # Transform document.xml
        txt = doc_xml_orig
        # 1. Remove MC brackets '（　）'
        removed_mc = txt.count(">（\u3000）<")
        txt = txt.replace(">（\u3000）<", "><")

        # 2. Align MC wrapped lines under 1st char (remove hanging on wide 1219/737 indents)
        txt = re.sub(
            r'<w:ind w:left="\d+" w:hanging="\d+"/>',
            f'<w:ind w:left="{mc_left_ind}"/>',
            txt,
        )

        # 3. Split sub-questions (　　)(1).. into independent hanging-indented paragraphs
        q_num = mc_count + 1
        single_run_pre = r'<w:r\b[^>]*>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:t[^>]*>'
        first_run_open = r'<w:r\b[^>]*>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?(?:<w:br/>)?<w:t[^>]*>\(</w:t></w:r>'
        pat_5 = (
            first_run_open
            + single_run_pre + r'\u3000\u3000</w:t></w:r>'
            + single_run_pre + r'\)\s*\(</w:t></w:r>'
            + single_run_pre + r'\d+</w:t></w:r>'
            + single_run_pre + r'\)([^<]*)</w:t></w:r>'
        )
        pat_4 = (
            first_run_open
            + single_run_pre + r'\u3000\u3000</w:t></w:r>'
            + single_run_pre + r'\)\s*</w:t></w:r>'
            + single_run_pre + r'\(\d+\)([^<]*)</w:t></w:r>'
        )
        pat_3 = (
            first_run_open
            + single_run_pre + r'\u3000\u3000</w:t></w:r>'
            + single_run_pre + r'\)\s*\(\d+\)([^<]*)</w:t></w:r>'
        )
        combined = re.compile(f"(?:{pat_5})|(?:{pat_4})|(?:{pat_3})", re.DOTALL)

        sub_p_open = (
            '</w:p><w:p><w:pPr>'
            '<w:pStyle w:val="1"/>'
            '<w:adjustRightInd w:val="0"/>'
            '<w:snapToGrid w:val="0"/>'
            f'<w:tabs><w:tab w:val="left" w:pos="{sub_left_ind}"/></w:tabs>'
            '<w:spacing w:before="20" w:after="0"/>'
            f'<w:ind w:left="{sub_left_ind}" w:hanging="{sub_hanging}"/>'
            '<w:rPr><w:rStyle w:val="char"/></w:rPr>'
            '</w:pPr>'
        )
        num_rpr = '<w:rPr><w:rStyle w:val="char"/><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/></w:rPr>'

        def repl(m: re.Match) -> str:
            nonlocal q_num
            tail = next((g for g in m.groups() if g is not None), "")
            tail_run = f'<w:r>{num_rpr}<w:t xml:space="preserve">{tail}</w:t></w:r>' if tail else ""
            out = f'{sub_p_open}<w:r>{num_rpr}<w:t>{q_num}.</w:t><w:tab/></w:r>{tail_run}'
            q_num += 1
            return out

        txt = combined.sub(repl, txt)
        txt = re.sub(
            r'<w:r\b[^>]*>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:br/></w:r></w:p>',
            '</w:p>',
            txt,
            flags=re.DOTALL,
        )
        sub_count = q_num - (mc_count + 1)
        if total_q <= 0:
            total_q = mc_count + sub_count

        # 3b. Renumber group stem paragraphs (after 題組 heading) with full-width digits １., ２., ３., ...
        #     so they are visually distinct from the continuous sub-question numbers (21., 22., ...).
        group_hdr_pos = txt.find("題組")
        if group_hdr_pos != -1:
            fw_digits = str.maketrans("0123456789", "０１２３４５６７８９")
            stem_idx = 0
            stem_rpr = '<w:rPr><w:rStyle w:val="char"/><w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體"/></w:rPr>'

            def _replace_stem_p(m: re.Match) -> str:
                nonlocal stem_idx
                p_open, ppr_inner = m.group(1), m.group(2)
                if "<w:numPr>" not in ppr_inner:
                    return m.group(0)
                stem_idx += 1
                fw_num = str(stem_idx).translate(fw_digits)
                new_ppr = (
                    '<w:pStyle w:val="1"/>'
                    '<w:adjustRightInd w:val="0"/>'
                    '<w:snapToGrid w:val="0"/>'
                    '<w:tabs><w:tab w:val="left" w:pos="480"/></w:tabs>'
                    '<w:spacing w:before="40" w:after="0"/>'
                    '<w:ind w:left="480" w:hanging="400"/>'
                    '<w:rPr><w:rStyle w:val="char"/></w:rPr>'
                )
                return f"{p_open}<w:pPr>{new_ppr}</w:pPr><w:r>{stem_rpr}<w:t>{fw_num}.</w:t><w:tab/></w:r>"

            tail_part = re.sub(
                r"(<w:p\b[^>]*>)<w:pPr>(.*?)</w:pPr>",
                _replace_stem_p,
                txt[group_hdr_pos:],
                flags=re.DOTALL,
            )
            txt = txt[:group_hdr_pos] + tail_part

        # 4. Ensure titlePg is enabled (Page 1 has header; Page 2 has NO header)
        if "<w:titlePg/>" not in txt:
            txt = txt.replace("</w:sectPr>", "<w:titlePg/></w:sectPr>")

        # Tighten top/bottom/header/footer margins cleanly
        txt = re.sub(
            r'w:top="\d+" w:right="(\d+)" w:bottom="\d+" w:left="(\d+)" w:header="\d+" w:footer="\d+"',
            r'w:top="600" w:right="\1" w:bottom="567" w:left="\2" w:header="300" w:footer="340"',
            txt,
            count=1,
        )

        # 5. Optional clean page break before a specific paragraph
        if break_para_id:
            marker = f'w14:paraId="{break_para_id}"'
            pos = txt.find(marker)
            if pos != -1:
                ppr_start = txt.find("<w:pPr>", pos)
                if ppr_start != -1 and ppr_start - pos < 400:
                    insert_at = ppr_start + len("<w:pPr>")
                    txt = txt[:insert_at] + "<w:pageBreakBefore/>" + txt[insert_at:]

        p1_labels = [str(i) for i in range(1, p1_end + 1)]
        p2_labels = [str(i) for i in range(p1_end + 1, total_q + 1)]
        note = f"{p1_end+1}～{mc_count} 為選擇題，{mc_count+1}～{total_q} 為題組" if p1_end < mc_count < total_q else ""

        hdr1_xml = build_page1_header_xml(title, grade, p1_labels, total_w, max_seat=max_seat)
        hdr2_empty_xml = build_page2_empty_header_xml()
        ftr_xml = build_page2_footer_xml(p2_labels, total_w, note=note)

        os.makedirs(os.path.dirname(os.path.abspath(out_docx)), exist_ok=True)
        with zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == hdr1_target:
                    zout.writestr(item, hdr1_xml.encode("utf-8"))
                elif hdr_default_target and item.filename == hdr_default_target:
                    zout.writestr(item, hdr2_empty_xml.encode("utf-8"))
                elif item.filename == ftr_default_target:
                    zout.writestr(item, ftr_xml.encode("utf-8"))
                elif item.filename == "word/document.xml":
                    zout.writestr(item, txt.encode("utf-8"))
                else:
                    zout.writestr(item, data)

    print(f"Saved: {out_docx}")
    print(f"  Removed MC brackets: {removed_mc}, Split sub-questions: {sub_count}")
    print(f"  Page 1 Header OMR: 1..{p1_end} | Page 2 Footer OMR: {p1_end+1}..{total_q}")

    if export_pdf:
        pdf_path = os.path.splitext(out_docx)[0] + ".pdf"
        pages = export_pdf_word_com(out_docx, pdf_path)
        print(f"Exported PDF: {pdf_path} (pages={pages})")


def main() -> None:
    ap = argparse.ArgumentParser(description="Convert Word exam (.docx) to Plan-A OMR exam paper.")
    ap.add_argument("--inspect", action="store_true", help="Inspect document structure and paraIds.")
    ap.add_argument("--input", required=True, help="Source .docx exam path.")
    ap.add_argument("--output", help="Output .docx path.")
    ap.add_argument("--title", default="範圍：小考", help="Exam title prefix, e.g. '範圍：CH4'.")
    ap.add_argument("--grade", default="8", help="Grade number, e.g. '8'.")
    ap.add_argument("--mc-count", type=int, default=20, help="Number of multiple-choice questions before groups.")
    ap.add_argument("--total", type=int, default=0, help="Total question count (0 = auto-detect).")
    ap.add_argument("--p1-end", type=int, default=18, help="Last question number on Page 1.")
    ap.add_argument("--max-seat", type=int, default=30, help="Maximum seat number (default: 30 -> 01~30).")
    ap.add_argument("--break-para-id", default=None, help="Optional w14:paraId to insert pageBreakBefore.")
    ap.add_argument("--export-pdf", action="store_true", help="Export companion PDF via Word COM.")
    args = ap.parse_args()

    if args.inspect:
        inspect_docx(args.input)
        return
    if not args.output:
        ap.error("--output is required unless --inspect is used.")

    convert_exam(
        src_docx=args.input,
        out_docx=args.output,
        title=args.title,
        grade=args.grade,
        p1_end=args.p1_end,
        mc_count=args.mc_count,
        total_q=args.total,
        max_seat=args.max_seat,
        break_para_id=args.break_para_id,
        export_pdf=args.export_pdf,
    )


if __name__ == "__main__":
    main()
