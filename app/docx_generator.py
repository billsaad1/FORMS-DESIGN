import io
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing shading if any
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=144, bottom=144, left=180, right=180):
    # cell margins in dxa (1 pt = 20 dxa, 1/144 inch = 10 dxa. 144 dxa is about 7.2 pt padding)
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing tcMar if any
    for child in list(tcPr):
        if child.tag.endswith('tcMar'):
            tcPr.remove(child)
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell, color="D3D3D3", sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing borders if any
    for child in list(tcPr):
        if child.tag.endswith('tcBorders'):
            tcPr.remove(child)
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    # Check if bidiVisual already exists
    bidi_exists = False
    for child in list(tblPr):
        if child.tag.endswith('bidiVisual'):
            bidi_exists = True
    if not bidi_exists:
        bidiVisual = OxmlElement('w:bidiVisual')
        tblPr.append(bidiVisual)

def set_table_fixed_layout(table):
    tblPr = table._tbl.tblPr
    # check if layout already exists
    layout_exists = False
    for child in list(tblPr):
        if child.tag.endswith('tblLayout'):
            child.set(qn('w:type'), 'fixed')
            layout_exists = True
    if not layout_exists:
        tblLayout = OxmlElement('w:tblLayout')
        tblLayout.set(qn('w:type'), 'fixed')
        tblPr.append(tblLayout)
    set_table_rtl(table)

def set_cell_width_and_properties(cell, width_in_inches):
    cell.width = Inches(width_in_inches)
    tcPr = cell._tc.get_or_add_tcPr()
    # check for existing tcW
    for child in list(tcPr):
        if child.tag.endswith('tcW'):
            tcPr.remove(child)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_in_inches * 1440))) # 1440 dxa per inch
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def add_header_table(doc, title_ar, title_en):
    # Set document properties for A4 with clean 0.5 inch margins
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # 1 row, 2 cols table for top logo & organization text
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_fixed_layout(table)

    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)

    # 5.27 inches for left, 2.0 inches for right = 7.27 inches total printable area on A4 with 0.5 margins
    set_cell_width_and_properties(cell_left, 5.27)
    set_cell_width_and_properties(cell_right, 2.0)

    # Left organization text (bilingual, cleanly stacked)
    p = cell_left.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT # Right align for bilingual consistency
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15

    run_ar = p.add_run("جمعية جيل البناء للتنمية الإنسانية\n")
    run_ar.bold = True
    run_ar.font.size = Pt(11.5)
    run_ar.font.name = 'Calibri'
    run_ar.font.color.rgb = RGBColor(0, 51, 102)

    run_en = p.add_run("Jeel Al Bena Association for Humanitarian Development\n")
    run_en.bold = True
    run_en.font.size = Pt(10)
    run_en.font.name = 'Calibri'
    run_en.font.color.rgb = RGBColor(0, 51, 102)

    run_dept = p.add_run("قسم الحماية / Protection Department")
    run_dept.italic = True
    run_dept.bold = True
    run_dept.font.size = Pt(9.5)
    run_dept.font.name = 'Calibri'
    run_dept.font.color.rgb = RGBColor(120, 120, 120)

    # Right logo image (under RTL table visual, cell_right is rendered on the left)
    p_logo = cell_right.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(0)

    logo_path = 'app/static/images/logo_emblem.png'
    if not os.path.exists(logo_path):
         logo_path = 'app/static/images/emblem.png'

    if os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Inches(0.65))
    else:
        p_logo.add_run("[Logo]")

    # Remove borders from header table
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if child.tag.endswith('tcBorders'):
                    tcPr.remove(child)
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
            tcPr.append(tcBorders)

    # Double bottom line colored separator
    p_sep = doc.add_paragraph()
    p_sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sep.paragraph_format.space_before = Pt(4)
    p_sep.paragraph_format.space_after = Pt(10)

    run_sep = p_sep.add_run("―" * 68)
    run_sep.bold = True
    run_sep.font.color.rgb = RGBColor(0, 51, 102)
    run_sep.font.size = Pt(10)

    # Big beautiful main title (Bilingual)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(12)
    p_title.paragraph_format.line_spacing = 1.15

    run_t_ar = p_title.add_run(title_ar + "\n")
    run_t_ar.bold = True
    run_t_ar.font.size = Pt(14)
    run_t_ar.font.color.rgb = RGBColor(0, 51, 102)
    run_t_ar.font.name = 'Calibri'

    run_t_en = p_title.add_run(title_en)
    run_t_en.bold = True
    run_t_en.font.size = Pt(12.5)
    run_t_en.font.color.rgb = RGBColor(0, 51, 102)
    run_t_en.font.name = 'Calibri'

def add_section_header(doc, ar_text, en_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True

    run = p.add_run(f"■ {ar_text} / {en_text}")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 51, 102)
    run.font.name = 'Calibri'

def create_row_style(row, col_widths, col_shades, col_alignments, col_texts, is_bold=False, is_header=False):
    # Helper to populate and style a table row cleanly
    for c_idx, cell in enumerate(row.cells):
        set_cell_width_and_properties(cell, col_widths[c_idx])
        set_cell_borders(cell, color="D3D3D3")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

        if col_shades[c_idx]:
            set_cell_background(cell, col_shades[c_idx])

        cell.text = "" # Clear default
        p = cell.paragraphs[0]
        p.alignment = col_alignments[c_idx]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15

        text = col_texts[c_idx]
        if isinstance(text, list):
            # Multiple items to be added with custom spacing
            for part in text:
                r = p.add_run(part)
                r.font.name = 'Calibri'
                r.font.size = Pt(9.5)
                if is_header:
                    r.bold = True
                    r.font.color.rgb = RGBColor(255, 255, 255)
                elif is_bold or "[ " in part or "☐" in part or "☑" in part:
                    r.bold = True
        else:
            r = p.add_run(text)
            r.font.name = 'Calibri'
            r.font.size = Pt(9.5)
            if is_header:
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            elif is_bold:
                r.bold = True

        # For RTL Arabic parts, enable RTL text flow
        p.paragraph_format.right_to_left = (col_alignments[c_idx] == WD_ALIGN_PARAGRAPH.RIGHT)

def generate_form1_docx():
    doc = Document()
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    # PAGE 1: Basic Information
    add_section_header(doc, "المعلومات الأساسية", "Basic Information")

    basic_table = doc.add_table(rows=6, cols=4)
    basic_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    basic_table.autofit = False
    set_table_fixed_layout(basic_table)

    col_widths = [1.8, 1.83, 1.8, 1.84]
    col_shades = ["F2F2F2", None, "F2F2F2", None]
    col_aligns = [WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]

    rows_data = [
        ("كود الحالة / Case Code", "........................................", "تاريخ التقييم / Assessment Date", "2025 / ______ / ______"),
        ("اسم الموظف / Staff Name", "........................................", "المنظمة / Organization", "........................................"),
        ("المحافظة / Province", "........................................", "المديرية / District", "........................................"),
        ("المركز / Center", "........................................", "الموقع / Location", "........................................"),
        ("اسم المستفيد / Beneficiary Name", "........................................", "رقم الهاتف / Phone Number", "........................................"),
        ("العمر / Age", ".................. (years / سنة)", "الجنس / Gender", "☐ ذكر / Male      ☐ أنثى / Female")
    ]

    for r_idx, row_data in enumerate(rows_data):
        row = basic_table.rows[r_idx]
        texts = [row_data[0], row_data[1], row_data[2], row_data[3]]
        create_row_style(row, col_widths, col_shades, col_aligns, texts, is_bold=False)
        # Bold labels specifically
        for c in [0, 2]:
            for run in row.cells[c].paragraphs[0].runs:
                run.bold = True

    # PAGE 1: Section 1 Demographics
    add_section_header(doc, "القسم الأول: البيانات الديموغرافية", "Section 1: Demographics")
    demo_table = doc.add_table(rows=4, cols=4)
    demo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    demo_table.autofit = False
    set_table_fixed_layout(demo_table)

    demo_rows = [
        ("الحالة الاجتماعية / Marital Status", "☐ أعزب / Single   ☐ متزوج / Married\n☐ أرمل / Widowed   ☐ مطلق / Divorced", "عدد أفراد الأسرة / Family Size", ".................. (members / أفراد)"),
        ("حالة النزوح / Displacement Status", "☐ نازح / IDP      ☐ عائد / Returnee\n☐ مجتمع مضيف / Host Community", "تاريخ النزوح / Date of Displacement", "______ / ______ / 2025"),
        ("العنوان الحالي / Current Address", "................................................................................", "العنوان الأصلي / Original Address", "................................................................................"),
        ("حالة التسجيل / Registration Status", "☐ مسجل / Registered   ☐ غير مسجل / Unregistered", "الاحتياجات الخاصة / Specific Needs", "☐ نعم / Yes   ☐ لا / No")
    ]

    for r_idx, row_data in enumerate(demo_rows):
        row = demo_table.rows[r_idx]
        texts = [row_data[0], row_data[1], row_data[2], row_data[3]]
        create_row_style(row, col_widths, col_shades, col_aligns, texts, is_bold=False)
        # Bold labels
        for c in [0, 2]:
            for run in row.cells[c].paragraphs[0].runs:
                run.bold = True

    # PAGE 1: Section 2 Specific Vulnerabilities
    add_section_header(doc, "القسم الثاني: نقاط الضعف الخاصة", "Section 2: Specific Vulnerabilities")

    vuln_table = doc.add_table(rows=5, cols=2)
    vuln_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    vuln_table.autofit = False
    set_table_fixed_layout(vuln_table)

    vuln_col_widths = [3.63, 3.64]
    vuln_col_shades = [None, None]
    vuln_col_aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]

    vuln_rows = [
        ("☐ طفل معرض للخطر / Child at risk", "☐ عائل وحيد / Single parent"),
        ("☐ مسن معرض للخطر / Older person at risk", "☐ حالة طبية حرجة / Serious medical condition"),
        ("☐ إعاقة جسدية / Physical disability", "☐ إعاقة ذهنية / Mental disability"),
        ("☐ امرأة حامل أو مرضعة / Pregnant or lactating", "☐ امرأة معرضة للخطر / Woman at risk"),
        ("☐ ناجٍ من العنف / Survivor of violence", "☐ غير ذلك (حدد) / Other (Specify): ........................................")
    ]

    for r_idx, row_data in enumerate(vuln_rows):
        row = vuln_table.rows[r_idx]
        create_row_style(row, vuln_col_widths, vuln_col_shades, vuln_col_aligns, [row_data[0], row_data[1]], is_bold=False)

    doc.add_page_break()

    # PAGE 2: Assessment Details & Displacement History
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")
    add_section_header(doc, "القسم الثالث: تفاصيل التقييم وتاريخ النزوح", "Section 3: Assessment Details & Displacement History")

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(4)
    p_desc.paragraph_format.space_after = Pt(4)
    run_desc = p_desc.add_run("أسباب النزوح والوضع الحالي ومصادر العيش الرئيسية للأسرة / Reasons for displacement, current situation & main livelihood sources:")
    run_desc.bold = True
    run_desc.font.name = 'Calibri'
    run_desc.font.size = Pt(10)

    desc_box = doc.add_table(rows=1, cols=1)
    desc_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    desc_box.autofit = False
    set_table_fixed_layout(desc_box)

    cell_box = desc_box.rows[0].cells[0]
    set_cell_width_and_properties(cell_box, 7.27)
    set_cell_borders(cell_box)
    set_cell_margins(cell_box, top=600, bottom=600, left=150, right=150) # Spacious drawing area
    cell_box.text = ""
    p_box = cell_box.paragraphs[0]
    p_box.paragraph_format.space_before = Pt(0)
    p_box.paragraph_format.space_after = Pt(0)
    r_empty = p_box.add_run("\n\n\n\n\n\n\n\n\n\n") # Generous height block
    r_empty.font.size = Pt(10)

    add_section_header(doc, "الظروف المعيشية والسكن", "Living Conditions & Accommodation")

    living_table = doc.add_table(rows=3, cols=2)
    living_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    living_table.autofit = False
    set_table_fixed_layout(living_table)

    living_col_widths = [2.2, 5.07]
    living_col_shades = ["F2F2F2", None]
    living_col_aligns = [WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]

    living_rows = [
        ("نوع السكن / Accommodation Type", "☐ إيجار / Rented     ☐ مع عائلة مضيفة / Host Family     ☐ خيمة / Tent     ☐ منزل متضرر / Damaged house"),
        ("مصدر الدخل / Source of Income", "................................................................................"),
        ("الدخل الشهري / Monthly Income", "........................................ dollars / دولار")
    ]

    for r_idx, row_data in enumerate(living_rows):
        row = living_table.rows[r_idx]
        create_row_style(row, living_col_widths, living_col_shades, living_col_aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    doc.add_page_break()

    # PAGE 3: Risk Analysis & Protection Concerns
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")
    add_section_header(doc, "القسم الرابع: تحليل المخاطر ومخاوف الحماية", "Section 4: Risk Analysis & Protection Concerns")

    p_risk_lbl = doc.add_paragraph()
    p_risk_lbl.paragraph_format.space_before = Pt(4)
    p_risk_lbl.paragraph_format.space_after = Pt(4)
    run_risk = p_risk_lbl.add_run("يرجى تقديم وصف تفصيلي لمخاطر الحماية التي يواجهها المستفيد وأسرته / Detailed protection risks description:")
    run_risk.bold = True
    run_risk.font.name = 'Calibri'
    run_risk.font.size = Pt(10)

    risk_box = doc.add_table(rows=1, cols=1)
    risk_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    risk_box.autofit = False
    set_table_fixed_layout(risk_box)

    cell_rbox = risk_box.rows[0].cells[0]
    set_cell_width_and_properties(cell_rbox, 7.27)
    set_cell_borders(cell_rbox)
    set_cell_margins(cell_rbox, top=800, bottom=800, left=150, right=150)
    cell_rbox.text = ""
    p_rbox = cell_rbox.paragraphs[0]
    p_rbox.paragraph_format.space_before = Pt(0)
    p_rbox.paragraph_format.space_after = Pt(0)
    r_rempty = p_rbox.add_run("\n\n\n\n\n\n\n\n\n\n\n\n")
    r_rempty.font.size = Pt(10)

    add_section_header(doc, "مخاوف ومخاطر الحماية المحددة", "Specific Protection Risks & Concerns")

    risk_table = doc.add_table(rows=4, cols=2)
    risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    risk_table.autofit = False
    set_table_fixed_layout(risk_table)

    risk_rows = [
        ("☐ تهديدات الأمن والسلامة / Safety & Security Threats", "☐ الضيق النفسي والاجتماعي / Psychosocial distress"),
        ("☐ عمالة الأطفال / Child labor", "☐ نقص الوثائق المدنية / Lack of civil documentation"),
        ("☐ خطر الإخلاء / Risk of eviction", "☐ التمييز والاستبعاد / Discrimination & exclusion"),
        ("☐ العنف القائم على النوع الاجتماعي / Gender-based violence", "☐ مخاطر أخرى / Other risks: ........................................")
    ]

    for r_idx, row_data in enumerate(risk_rows):
        row = risk_table.rows[r_idx]
        create_row_style(row, vuln_col_widths, vuln_col_shades, vuln_col_aligns, [row_data[0], row_data[1]], is_bold=False)

    doc.add_page_break()

    # PAGE 4: Action Plan & Referrals
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")
    add_section_header(doc, "القسم الخامس: خطة العمل والإحالات", "Section 5: Action Plan & Referrals")

    action_table = doc.add_table(rows=4, cols=2)
    action_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    action_table.autofit = False
    set_table_fixed_layout(action_table)

    action_rows = [
        ("المساعدة الموصى بها / Recommended Assistance", "☐ خدمات الحماية / Protection services   ☐ المساعدة النقدية / Cash support\n☐ دعم المأوى / Shelter support          ☐ المساعدة القانونية / Legal support"),
        ("إحالة داخلية / Internal Referral", "إلى قسم: ................................................................................"),
        ("إحالة خارجية / External Referral", "إلى منظمة/جهة: ........................................................................"),
        ("الإجراء المتخذ فورا / Immediate Action Taken", "........................................................................................................")
    ]

    for r_idx, row_data in enumerate(action_rows):
        row = action_table.rows[r_idx]
        create_row_style(row, living_col_widths, living_col_shades, living_col_aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    add_section_header(doc, "التوقيعات والموافقات", "Signatures & Approvals")

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    set_table_fixed_layout(sig_table)

    sig_rows = [
        ("اسم وتوقيع الموظف المقيم / Assessor Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025",
         "اسم وتوقيع مشرف الحماية / Protection Supervisor Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]

    row = sig_table.rows[0]
    create_row_style(row, vuln_col_widths, vuln_col_shades, vuln_col_aligns, [sig_rows[0][0], sig_rows[0][1]], is_bold=False)
    # Style the first lines bold
    for c in range(2):
        cell_p = row.cells[c].paragraphs[0]
        # Reinforce padding for signature block
        set_cell_margins(row.cells[c], top=200, bottom=200, left=150, right=150)
        # Bold the first line
        cell_p.runs[0].bold = True

    # Save to buffer
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()


def generate_form2_docx():
    doc = Document()
    add_header_table(doc, "نموذج دراسة الحالة وإدارة الحالة", "Protection Case Study and Case Management Form")

    # PAGE 1: Case Registration
    add_section_header(doc, "القسم الأول: تسجيل الحالة والمعلومات العامة", "Section 1: Case Registration & General Information")

    reg_table = doc.add_table(rows=5, cols=4)
    reg_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    reg_table.autofit = False
    set_table_fixed_layout(reg_table)

    col_widths = [1.8, 1.83, 1.8, 1.84]
    col_shades = ["F2F2F2", None, "F2F2F2", None]
    col_aligns = [WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]

    reg_rows = [
        ("كود الحالة / Case Code", "........................................", "تاريخ فتح الحالة / Case Opening Date", "______/______/2025"),
        ("مسؤول الحالة / Case Worker", "........................................", "مصدر الإحالة / Referral Source", "........................................"),
        ("الاسم الكامل للمستفيد / Full Name", "................................................................................", "العمر / Age", ".................. years / سنة"),
        ("الجنس / Gender", "☐ ذكر / Male      ☐ أنثى / Female", "رقم الهاتف / Phone Number", "........................................"),
        ("العنوان الحالي / Current Address", "................................................................................", "نوع الوثيقة القانونية / Legal Document", "........................................")
    ]

    for r_idx, row_data in enumerate(reg_rows):
        row = reg_table.rows[r_idx]
        texts = [row_data[0], row_data[1], row_data[2], row_data[3]]
        create_row_style(row, col_widths, col_shades, col_aligns, texts, is_bold=False)
        for c in [0, 2]:
            for run in row.cells[c].paragraphs[0].runs:
                run.bold = True

    # Section 2: Description & Assessment
    add_section_header(doc, "الالقسم الثاني: وصف وتقييم الحالة", "Section 2: Case Description & Assessment")

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(4)
    p_desc.paragraph_format.space_after = Pt(4)
    run_desc = p_desc.add_run("الخلفية التفصيلية للحالة ومخاطر الحماية الفورية والوضع الاجتماعي والاقتصادي / Detailed case background, immediate protection risks & socio-economic status:")
    run_desc.bold = True
    run_desc.font.name = 'Calibri'
    run_desc.font.size = Pt(10)

    desc_box = doc.add_table(rows=1, cols=1)
    desc_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    desc_box.autofit = False
    set_table_fixed_layout(desc_box)

    cell_box = desc_box.rows[0].cells[0]
    set_cell_width_and_properties(cell_box, 7.27)
    set_cell_borders(cell_box)
    set_cell_margins(cell_box, top=800, bottom=800, left=150, right=150)
    cell_box.text = ""
    p_box = cell_box.paragraphs[0]
    r_empty = p_box.add_run("\n\n\n\n\n\n\n\n\n\n\n\n\n\n")
    r_empty.font.size = Pt(10)

    doc.add_page_break()

    # PAGE 2: Action Plan & Follow-Up
    add_header_table(doc, "نموذج دراسة الحالة وإدارة الحالة", "Protection Case Study and Case Management Form")

    add_section_header(doc, "القسم الثالث: خطة عمل الحالة", "Section 3: Case Action Plan")

    # 5-column plan table with Navy backgrounds
    plan_table = doc.add_table(rows=4, cols=5)
    plan_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    plan_table.autofit = False
    set_table_fixed_layout(plan_table)

    plan_widths = [1.45, 1.45, 1.45, 1.45, 1.47]
    plan_headers = [
        "الهدف / Objective",
        "الإجراء المطلوب / Action Required",
        "الشخص المسؤول / Responsible",
        "الإطار الزمني / Timeframe",
        "الحالة / Status"
    ]

    # Header row style
    create_row_style(plan_table.rows[0], plan_widths, ["003366"]*5, [WD_ALIGN_PARAGRAPH.CENTER]*5, plan_headers, is_bold=True, is_header=True)

    # Content rows empty blocks
    for r_idx in range(1, 4):
        row = plan_table.rows[r_idx]
        create_row_style(row, plan_widths, [None]*5, [WD_ALIGN_PARAGRAPH.LEFT]*5, ["\n\n", "\n\n", "\n\n", "\n\n", "\n\n"], is_bold=False)

    add_section_header(doc, "القسم الرابع: المتابعة وإغلاق الحالة", "Section 4: Follow-up & Case Closure")

    close_table = doc.add_table(rows=3, cols=2)
    close_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    close_table.autofit = False
    set_table_fixed_layout(close_table)

    close_widths = [2.2, 5.07]
    close_shades = ["F2F2F2", None]
    close_aligns = [WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]

    close_rows = [
        ("ملخص زيارات المتابعة وتحديثاتها / Follow-up summary", "................................................................................................................................................"),
        ("حالة الحالة / Case Status", "☐ نشطة / Active         ☐ مغلقة / Closed         ☐ معلقة / Suspended"),
        ("أسباب الإغلاق / Reasons for closure", "☐ تم تحقيق الأهداف / Goals met    ☐ اكتمال الإحالة / Referral complete    ☐ انتقال / Relocation")
    ]

    for r_idx, row_data in enumerate(close_rows):
        row = close_table.rows[r_idx]
        create_row_style(row, close_widths, close_shades, close_aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    add_section_header(doc, "التوقيعات والموافقات", "Signatures & Approvals")

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    set_table_fixed_layout(sig_table)

    sig_rows = [
        ("اسم وتوقيع مسؤول الحالة / Case Worker Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025",
         "اسم وتوقيع مشرف إدارة الحالة / Case Management Supervisor Name\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]

    row = sig_table.rows[0]
    create_row_style(row, [3.63, 3.64], [None, None], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT], [sig_rows[0][0], sig_rows[0][1]], is_bold=False)
    for c in range(2):
        set_cell_margins(row.cells[c], top=200, bottom=200, left=150, right=150)
        row.cells[c].paragraphs[0].runs[0].bold = True

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()


def generate_form3_docx():
    doc = Document()
    add_header_table(doc, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    # PAGE 1: Basic Info
    add_section_header(doc, "المعلومات الأساسية", "Basic Information")

    info_table = doc.add_table(rows=4, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False
    set_table_fixed_layout(info_table)

    widths = [2.2, 5.07]
    shades = ["F2F2F2", None]
    aligns = [WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]

    info_rows = [
        ("معرّف الحالة / Case ID", "................................................................................"),
        ("تاريخ التقييم / Evaluation Date", "2025 / ______ / ______"),
        ("تاريخ قرار الموافقة / Approval Decision Date", "2025 / ______ / ______"),
        ("نوع الخطر الحمائي / Protection Risk Type", "................................................................................")
    ]

    for r_idx, row_data in enumerate(info_rows):
        row = info_table.rows[r_idx]
        create_row_style(row, widths, shades, aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    # Section 1: Case Summary
    add_section_header(doc, "القسم الأول: ملخص الحالة", "Section 1: Case Summary")

    p_bdata = doc.add_paragraph()
    p_bdata.paragraph_format.space_before = Pt(4)
    p_bdata.paragraph_format.space_after = Pt(4)
    run_bdata = p_bdata.add_run("1.1 البيانات الأساسية للمستفيد / Beneficiary Basic Data")
    run_bdata.bold = True
    run_bdata.font.name = 'Calibri'
    run_bdata.font.size = Pt(10)

    bdata_table = doc.add_table(rows=4, cols=2)
    bdata_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    bdata_table.autofit = False
    set_table_fixed_layout(bdata_table)

    bdata_rows = [
        ("الاسم أو معرف / Name or ID", "................................................................................"),
        ("العمر / Age", ".................. سنة / years"),
        ("الجنس / Gender", "☐ ذكر / Male         ☐ أنثى / Female"),
        ("عدد أفراد الأسرة / Family Members", ".................. أفراد / members")
    ]

    for r_idx, row_data in enumerate(bdata_rows):
        row = bdata_table.rows[r_idx]
        create_row_style(row, widths, shades, aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    # Risk Summary
    p_risk = doc.add_paragraph()
    p_risk.paragraph_format.space_before = Pt(8)
    p_risk.paragraph_format.space_after = Pt(4)
    run_risk = p_risk.add_run("1.2 ملخص الخطر / Risk Summary")
    run_risk.bold = True
    run_risk.font.name = 'Calibri'
    run_risk.font.size = Pt(10)

    risk_table = doc.add_table(rows=2, cols=2)
    risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    risk_table.autofit = False
    set_table_fixed_layout(risk_table)

    risk_rows = [
        ("الخطر الأساسي / Primary Risk", "\n\n\n"),
        ("الشدة / Severity", "☐ حاد وفوري جداً (24-48 ساعة) / Very acute and immediate (24-48 hours)\n☐ حاد (3-5 أيام) / Acute (3-5 days)\n☐ متوسط (أسبوع - أسبوعين) / Moderate (1-2 weeks)")
    ]

    for r_idx, row_data in enumerate(risk_rows):
        row = risk_table.rows[r_idx]
        create_row_style(row, widths, shades, aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    doc.add_page_break()

    # PAGE 2: Proposed Solution & Decisions
    add_header_table(doc, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    # 1.3 Proposed Solution & Amount
    add_section_header(doc, "1.3 الحل المقترح والمبلغ", "Proposed Solution and Amount")

    sol_table = doc.add_table(rows=3, cols=2)
    sol_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sol_table.autofit = False
    set_table_fixed_layout(sol_table)

    sol_rows = [
        ("كيف سيساعد النقد؟ / How will cash help?", "\n\n"),
        ("المبلغ المقترح / Proposed Amount", "........................................ دولار / dollars"),
        ("السبب / Reason", "\n\n")
    ]

    for r_idx, row_data in enumerate(sol_rows):
        row = sol_table.rows[r_idx]
        create_row_style(row, widths, shades, aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    # Section 2: Approval Reviews
    add_section_header(doc, "القسم الثاني: موافقة قسم الحماية", "Section 2: Protection Department Approval")

    p_rev = doc.add_paragraph()
    p_rev.paragraph_format.space_before = Pt(4)
    p_rev.paragraph_format.space_after = Pt(4)
    run_rev = p_rev.add_run("2.1 مراجعة التقييم / Evaluation Review")
    run_rev.bold = True
    run_rev.font.name = 'Calibri'
    run_rev.font.size = Pt(10)

    rev_table = doc.add_table(rows=4, cols=2)
    rev_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    rev_table.autofit = False
    set_table_fixed_layout(rev_table)

    rev_rows = [
        ("وضوح الخطر / Risk Clarity", "☐ واضح جداً / Very Clear      ☐ واضح / Clear      ☐ غامض / Vague"),
        ("الربط السببي / Causal Link", "☐ منطقي جداً / Very Logical   ☐ منطقي / Logical   ☐ ضعيف / Weak"),
        ("كفاية التقييم / Evaluation Adequacy", "☐ شامل / Comprehensive         ☐ متوسط / Moderate     ☐ ناقص / Incomplete"),
        ("المخاطر الثانوية / Secondary Risks", "☐ منخفضة / Low               ☐ متوسطة / Medium     ☐ عالية / High")
    ]

    for r_idx, row_data in enumerate(rev_rows):
        row = rev_table.rows[r_idx]
        create_row_style(row, widths, shades, aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    # Section 2.2: Protection Officer's Decision
    p_off = doc.add_paragraph()
    p_off.paragraph_format.space_before = Pt(8)
    p_off.paragraph_format.space_after = Pt(4)
    run_off = p_off.add_run("2.2 قرار ضابط الحماية / Protection Officer's Decision")
    run_off.bold = True
    run_off.font.name = 'Calibri'
    run_off.font.size = Pt(10)

    off_table = doc.add_table(rows=1, cols=1)
    off_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    off_table.autofit = False
    set_table_fixed_layout(off_table)

    cell_off = off_table.rows[0].cells[0]
    set_cell_width_and_properties(cell_off, 7.27)
    set_cell_borders(cell_off)
    set_cell_margins(cell_off, top=140, bottom=140, left=150, right=150)
    cell_off.text = ""
    p_off_cell = cell_off.paragraphs[0]
    p_off_cell.paragraph_format.line_spacing = 1.25
    p_off_cell.add_run("هل توصي بالموافقة على صرف النقد؟ / Do you recommend approval for cash disbursement?\n").bold = True
    p_off_cell.add_run("☑ نعم، موافق تماماً / Yes, fully approve   ").bold = True
    p_off_cell.add_run("☐ نعم، مع شروط / Yes, with conditions:\n\n").bold = True
    p_off_cell.add_run("الشروط / Conditions: ............................................................................................................\n\n").italic = True
    p_off_cell.add_run("☐ لا، غير موافق / No, do not approve\n\n").bold = True
    p_off_cell.add_run("السبب / Reason: ................................................................................................................\n\n").italic = True
    p_off_cell.add_run("اسم وتوقيع ضابط الحماية / Protection Officer Name & Signature: .....................................................").bold = True

    # Section 3: Final Decision from Protection Manager
    add_section_header(doc, "القسم الثالث: القرار النهائي من مدير الحماية", "Section 3: Final Decision from Protection Manager")

    mgr_table = doc.add_table(rows=3, cols=2)
    mgr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    mgr_table.autofit = False
    set_table_fixed_layout(mgr_table)

    mgr_rows = [
        ("القرار النهائي للمدير / Final Decision", "☐ موافق على الصرف / Approved       ☐ غير موافق / Not Approved"),
        ("المبلغ وآلية الصرف / Amount & Mechanism", "المبلغ: ........................ دولار    الآلية: ☐ نقد / Cash   ☐ حوالة / Transfer   ☐ دفع مباشر / Direct Payment"),
        ("التوقيع والتاريخ / Signature & Date", "الاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]

    for r_idx, row_data in enumerate(mgr_rows):
        row = mgr_table.rows[r_idx]
        create_row_style(row, widths, shades, aligns, [row_data[0], row_data[1]], is_bold=False)
        for r_run in row.cells[0].paragraphs[0].runs:
            r_run.bold = True

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()
