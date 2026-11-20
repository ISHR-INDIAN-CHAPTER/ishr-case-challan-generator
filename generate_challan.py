import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_challan(output_path="ISHR_Indian_Chapter_Case_Challan_14x22cm.docx"):
    doc = docx.Document()

    # Page setup: Exact 14 cm (width) x 22 cm (height)
    section = doc.sections[0]
    section.page_width = Cm(14.0)
    section.page_height = Cm(22.0)
    section.top_margin = Cm(0.75)
    section.bottom_margin = Cm(0.75)
    section.left_margin = Cm(0.85)
    section.right_margin = Cm(0.85)

    # Double-line formal border around the page
    pgBorders = parse_xml(
        f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
        f'<w:top w:val="double" w:sz="10" w:space="12" w:color="000000"/>'
        f'<w:left w:val="double" w:sz="10" w:space="12" w:color="000000"/>'
        f'<w:bottom w:val="double" w:sz="10" w:space="12" w:color="000000"/>'
        f'<w:right w:val="double" w:sz="10" w:space="12" w:color="000000"/>'
        f'</w:pgBorders>'
    )
    section._sectPr.append(pgBorders)

    def set_cell_margins(cell, top=32, bottom=32, left=65, right=65):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def set_table_borders(table, color="000000", sz="6", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def set_cell_shading(cell, fill_hex="F2F2F2"):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    BLACK = RGBColor(0, 0, 0)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(8.5)
    style.font.color.rgb = BLACK

    # Top Form Code
    p_form = doc.add_paragraph()
    p_form.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_form.paragraph_format.space_before = Pt(0)
    p_form.paragraph_format.space_after = Pt(1)
    r_form = p_form.add_run("FORM – ISHR/CH-01  |  ORIGINAL / OFFICE COPY")
    r_form.bold = True
    r_form.font.size = Pt(6.8)

    # Header Block
    p_org = doc.add_paragraph()
    p_org.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_org.paragraph_format.space_before = Pt(0)
    p_org.paragraph_format.space_after = Pt(1)
    r_org = p_org.add_run("INTERNATIONAL SOCIETY FOR HUMAN RIGHTS (ISHR) – INDIAN CHAPTER")
    r_org.bold = True
    r_org.font.size = Pt(9.8)

    p_sub1 = doc.add_paragraph()
    p_sub1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub1.paragraph_format.space_before = Pt(0)
    p_sub1.paragraph_format.space_after = Pt(1)
    r_sub1 = p_sub1.add_run("(Independent International Human Rights Body)")
    r_sub1.bold = True
    r_sub1.italic = True
    r_sub1.font.size = Pt(8)

    p_sub2 = doc.add_paragraph()
    p_sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub2.paragraph_format.space_before = Pt(0)
    p_sub2.paragraph_format.space_after = Pt(1)
    r_sub2 = p_sub2.add_run("Under UDHR, ICCPR & The Protection of Human Rights Act, 1993")
    r_sub2.font.size = Pt(7.5)

    p_sub3 = doc.add_paragraph()
    p_sub3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub3.paragraph_format.space_before = Pt(0)
    p_sub3.paragraph_format.space_after = Pt(3)
    r_sub3 = p_sub3.add_run("NSO: Tirupati – 517501, A.P., India | E-mail: ishr.india.secretariat@gmail.com | ishr.indian.info@gmail.com")
    r_sub3.font.size = Pt(7)

    pPr_h = p_sub3._p.get_or_add_pPr()
    pBdr_h = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="10" w:space="3" w:color="000000"/></w:pBdr>')
    pPr_h.append(pBdr_h)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("OFFICIAL CASE CHALLAN & ACKNOWLEDGMENT RECEIPT")
    r_title.bold = True
    r_title.underline = True
    r_title.font.size = Pt(9.5)

    # Metadata Grid
    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    set_table_borders(meta_table, color="000000", sz="6")

    col_widths_meta = [Cm(2.45), Cm(4.25), Cm(2.25), Cm(3.35)]
    meta_rows = [
        [("Challan No.:", True), ("ISHR/IND/20___/ [ _ _ _ _ ]", False), ("Issue Date:", True), ("___ / ___ / 20___", False)],
        [("Case Ref. No.:", True), ("_______________________", False), ("Office:", True), ("NSO, Tirupati – 517501", False)]
    ]

    for r_idx, row in enumerate(meta_rows):
        for c_idx, (text, is_label) in enumerate(row):
            cell = meta_table.cell(r_idx, c_idx)
            cell.width = col_widths_meta[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=32, bottom=32, left=55, right=55)
            if is_label:
                set_cell_shading(cell, "EFEFEF")
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.bold = is_label
            run.font.size = Pt(7.8)

    def add_gov_part_header(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(8.5)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>')
        pPr.append(pBdr)

    # Part I
    add_gov_part_header("PART – I : PARTICULARS OF THE COMPLAINANT")
    comp_table = doc.add_table(rows=5, cols=3)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_table.autofit = False
    set_table_borders(comp_table, color="000000", sz="6")

    comp_data = [
        ("1.", "Complainant Name\n(Block Letters)", "__________________________________________________________"),
        ("2.", "Father's / Spouse's Name", "__________________________________________________________"),
        ("3.", "Phone Number", "+91 – ______________________  (Alt: ________________________)"),
        ("4.", "Aadhaar Number", "[ ____ – ____ – ____ ]   (Copy enclosed: Yes / No)"),
        ("5.", "Place & Pincode", "Place: __________________________________  Pin: ____________")
    ]

    col_widths_comp = [Cm(0.65), Cm(3.35), Cm(8.30)]
    for r_idx, (sl, label, field) in enumerate(comp_data):
        for c_idx, txt in enumerate([sl, label, field]):
            cell = comp_table.cell(r_idx, c_idx)
            cell.width = col_widths_comp[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=34, bottom=34, left=60, right=60)
            if c_idx in (0, 1):
                set_cell_shading(cell, "F5F5F5")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.08
            r = p.add_run(txt)
            r.bold = (c_idx in (0, 1))
            r.font.size = Pt(7.8)

    # Part II
    add_gov_part_header("PART – II : CASE DETAILS & VOLUNTARY PAYABLE PARTICULARS")
    case_table = doc.add_table(rows=4, cols=3)
    case_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    case_table.autofit = False
    set_table_borders(case_table, color="000000", sz="6")

    case_data = [
        ("1.", "Nature of Complaint",
         "__________________________________________________________\n__________________________________________________________"),
        ("2.", "Voluntary Payable\n(in Figures)",
         "Rs. ___________________ /-  (Strictly Voluntary Administrative Support)"),
        ("3.", "Amount in Words",
         "Rupees _______________________________________________ Only"),
        ("4.", "Payment Mode & Ref.",
         "[ ] Cash   [ ] UPI/Online   [ ] DD/Cheque   Ref: ___________________")
    ]

    for r_idx, (sl, label, field) in enumerate(case_data):
        for c_idx, txt in enumerate([sl, label, field]):
            cell = case_table.cell(r_idx, c_idx)
            cell.width = col_widths_comp[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=34, bottom=34, left=60, right=60)
            if c_idx in (0, 1):
                set_cell_shading(cell, "F5F5F5")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.08
            r = p.add_run(txt)
            r.bold = (c_idx in (0, 1))
            r.font.size = Pt(7.8)

    # Part III
    add_gov_part_header("PART – III : SOLEMN DECLARATION BY THE COMPLAINANT")
    decl_table = doc.add_table(rows=1, cols=1)
    decl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    decl_table.autofit = False
    decl_table.columns[0].width = Cm(12.30)
    set_table_borders(decl_table, color="000000", sz="8")

    d_cell = decl_table.cell(0, 0)
    set_cell_margins(d_cell, top=44, bottom=44, left=75, right=75)

    dp = d_cell.paragraphs[0]
    dp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    dp.paragraph_format.space_before = Pt(0)
    dp.paragraph_format.space_after = Pt(0)
    dp.paragraph_format.line_spacing = 1.1

    dp.add_run("I, ").font.size = Pt(7.6)
    r_d2 = dp.add_run("________________________________________")
    r_d2.bold = True
    r_d2.font.size = Pt(7.6)
    dp.add_run(", S/o, D/o, W/o ______________________________, solemnly declare that the particulars furnished above are true and correct to the best of my knowledge and belief. I further confirm that any sum remitted as ").font.size = Pt(7.6)
    r_d4 = dp.add_run("Voluntary Payable")
    r_d4.bold = True
    r_d4.font.size = Pt(7.6)
    dp.add_run(" towards documentation or case-processing support is paid entirely out of my own free will without coercion, and is non-refundable.").font.size = Pt(7.6)

    # Signatures & Seal
    p_place_date = doc.add_paragraph()
    p_place_date.paragraph_format.space_before = Pt(5)
    p_place_date.paragraph_format.space_after = Pt(3)
    r_pd = p_place_date.add_run("Place: Tirupati (517501)                                                       Date: ___ / ___ / 20___")
    r_pd.bold = True
    r_pd.font.size = Pt(7.8)

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    set_table_borders(sig_table, color="000000", sz="6")

    sc0, sc1 = sig_table.cell(0, 0), sig_table.cell(0, 1)
    sc0.width, sc1.width = Cm(6.15), Cm(6.15)
    sc0.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    sc1.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    set_cell_margins(sc0, top=210, bottom=40, left=60, right=60)
    set_cell_margins(sc1, top=210, bottom=40, left=60, right=60)

    sp0 = sc0.paragraphs[0]
    sp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp0.paragraph_format.space_before = Pt(0)
    sp0.paragraph_format.space_after = Pt(0)
    sp0.paragraph_format.line_spacing = 1.08
    sr0 = sp0.add_run("_______________________________\n")
    sr0.bold = True
    sr0.font.size = Pt(7.8)
    sr0_lbl = sp0.add_run("Signature of Complainant\n")
    sr0_lbl.bold = True
    sr0_lbl.font.size = Pt(7.8)
    sp0.add_run("(Name: _________________________)").font.size = Pt(7.2)

    sp1 = sc1.paragraphs[0]
    sp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp1.paragraph_format.space_before = Pt(0)
    sp1.paragraph_format.space_after = Pt(0)
    sp1.paragraph_format.line_spacing = 1.08
    sr1 = sp1.add_run("_______________________________\n")
    sr1.bold = True
    sr1.font.size = Pt(7.8)
    sr1_lbl = sp1.add_run("Seal & Authorized Signatory\n")
    sr1_lbl.bold = True
    sr1_lbl.font.size = Pt(7.8)
    sp1.add_run("ISHR – Indian Chapter, Tirupati (517501)").font.size = Pt(7.2)

    # Footer
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(4)
    p_foot.paragraph_format.space_after = Pt(0)
    r_foot = p_foot.add_run("NOTE: Quote the 4-digit Challan Number in all future correspondence with the Secretariat.")
    r_foot.bold = True
    r_foot.font.size = Pt(7)

    doc.save(output_path)
    print(f"Generated: {output_path}")

if __name__ == "__main__":
    create_challan()
  
