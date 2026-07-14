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
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell, color="D3D3D3", sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def add_header_table(doc, title_ar, title_en):
    # Set document properties for A4
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # 1 row, 2 cols table for top logo & organization
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Left cell: Text, Right cell: Logo
    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)
    cell_left.width = Inches(5.27)
    cell_right.width = Inches(2.0)

    # Left text
    p = cell_left.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run1 = p.add_run("جمعية جيل البناء للتنمية الإنسانية\n")
    run1.bold = True
    run1.font.size = Pt(11)
    run1.font.name = 'Arial'

    run2 = p.add_run("Jeel Al Bena Association for Humanitarian Development\n")
    run2.font.size = Pt(10)
    run2.font.name = 'Arial'

    run3 = p.add_run("قسم الحماية / Protection Department")
    run3.italic = True
    run3.font.size = Pt(9.5)
    run3.font.name = 'Arial'

    # Right logo
    p_logo = cell_right.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    logo_path = 'app/static/images/emblem.png'
    if os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Inches(0.65))
    else:
        p_logo.add_run("[Logo / شعار]")

    # Clear borders of header table
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
            tcPr.append(tcBorders)

    # Divider line
    div_p = doc.add_paragraph()
    div_p.paragraph_format.space_before = Pt(4)
    div_p.paragraph_format.space_after = Pt(8)
    run_div = div_p.add_run("_________________________________________________________________________________")
    run_div.font.color.rgb = RGBColor(0, 51, 102)
    run_div.font.size = Pt(10)

    # Big main title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(12)

    run_t_ar = title_p.add_run(title_ar + "\n")
    run_t_ar.bold = True
    run_t_ar.font.size = Pt(14)
    run_t_ar.font.color.rgb = RGBColor(0, 51, 102)
    run_t_ar.font.name = 'Arial'

    run_t_en = title_p.add_run(title_en)
    run_t_en.bold = True
    run_t_en.font.size = Pt(13)
    run_t_en.font.color.rgb = RGBColor(0, 51, 102)
    run_t_en.font.name = 'Arial'

def add_section_header(doc, ar_text, en_text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True

    # Style it with deep blue background or left accent bar
    run = p.add_run(f"■ {ar_text} / {en_text}")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 51, 102)
    run.font.name = 'Arial'

def generate_form1_docx():
    doc = Document()
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    # PAGE 1
    add_section_header(doc, "المعلومات الأساسية", "Basic Information")

    table = doc.add_table(rows=6, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    headers = [
        ("كود الحالة / Case Code", "........................................", "تاريخ التقييم / Assessment Date", "2025 / ______ / ______"),
        ("اسم الموظف / Staff Name", "........................................", "المنظمة / Organization", "........................................"),
        ("المحافظة / Province", "........................................", "المديرية / District", "........................................"),
        ("المركز / Center", "........................................", "الموقع / Location", "........................................"),
        ("اسم المستفيد / Beneficiary Name", "........................................", "رقم الهاتف / Phone Number", "........................................"),
        ("العمر / Age", ".................. (years / سنة)", "الجنس / Gender", "[  ] ذكر / Male      [  ] أنثى / Female")
    ]

    col_widths = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8)]
    for r_idx, row_data in enumerate(headers):
        row = table.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx % 2 == 0 else WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)
                if c_idx % 2 == 0:
                    run.bold = True
                    set_cell_background(cell, "F2F2F2")

    add_section_header(doc, "القسم الأول: البيانات الديموغرافية", "Section 1: Demographics")
    table_demo = doc.add_table(rows=4, cols=4)
    table_demo.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_demo.autofit = False

    demo_data = [
        ("الحالة الاجتماعية / Marital Status", "[  ] أعزب / Single   [  ] متزوج / Married   [  ] أرمل / Widowed   [  ] مطلق / Divorced", "عدد أفراد الأسرة / Family Size", ".................. (members / أفراد)"),
        ("حالة النزوح / Displacement Status", "[  ] نازح / IDP   [  ] عائد / Returnee   [  ] مجتمع مضيف / Host", "تاريخ النزوح / Date of Displacement", "______ / ______ / 2025"),
        ("العنوان الحالي / Current Address", "................................................................................", "العنوان الأصلي / Original Address", "................................................................................"),
        ("حالة التسجيل / Registration Status", "[  ] مسجل / Registered   [  ] غير مسجل / Unregistered", "الاحتياجات الخاصة / Specific Needs", "[  ] نعم / Yes   [  ] لا / No")
    ]
    for r_idx, row_data in enumerate(demo_data):
        row = table_demo.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx % 2 == 0 else WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)
                if c_idx % 2 == 0:
                    run.bold = True
                    set_cell_background(cell, "F2F2F2")

    add_section_header(doc, "القسم الثاني: نقاط الضعف الخاصة", "Section 2: Specific Vulnerabilities")

    vuln_table = doc.add_table(rows=5, cols=2)
    vuln_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    vuln_table.autofit = False

    vuln_data = [
        ("[  ] طفل معرض للخطر / Child at risk", "[  ] عائل وحيد / Single parent"),
        ("[  ] مسن معرض للخطر / Older person at risk", "[  ] حالة طبية حرجة / Serious medical condition"),
        ("[  ] إعاقة جسدية / Physical disability", "[  ] إعاقة ذهنية / Mental disability"),
        ("[  ] امرأة حامل أو مرضعة / Pregnant or lactating", "[  ] امرأة معرضة للخطر / Woman at risk"),
        ("[  ] ناجٍ من العنف / Survivor of violence", "[  ] غير ذلك (حدد) / Other (Specify): ........................................")
    ]
    for r_idx, row_data in enumerate(vuln_data):
        row = vuln_table.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Inches(3.6)
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)

    doc.add_page_break()

    # PAGE 2
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")
    add_section_header(doc, "القسم الثالث: تفاصيل التقييم وتاريخ النزوح", "Section 3: Assessment Details & Displacement History")

    p_desc = doc.add_paragraph()
    run_desc = p_desc.add_run("أسباب النزوح والوضع الحالي / Reasons for displacement and current situation:")
    run_desc.bold = True
    run_desc.font.name = 'Arial'
    run_desc.font.size = Pt(10)

    desc_box = doc.add_table(rows=1, cols=1)
    desc_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    desc_box.rows[0].cells[0].width = Inches(7.2)
    set_cell_borders(desc_box.rows[0].cells[0])
    set_cell_margins(desc_box.rows[0].cells[0], top=400, bottom=400, left=150, right=150)
    p_box = desc_box.rows[0].cells[0].paragraphs[0]
    p_box.text = "\n\n\n\n"

    add_section_header(doc, "الظروف المعيشية والسكن", "Living Conditions & Accommodation")

    living_table = doc.add_table(rows=3, cols=2)
    living_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    living_table.autofit = False

    living_data = [
        ("نوع السكن / Accommodation Type", "[  ] إيجار / Rented     [  ] مع عائلة مضيفة / Host Family     [  ] خيمة / Tent     [  ] منزل متضرر / Damaged house"),
        ("مصدر الدخل / Source of Income", "................................................................................"),
        ("الدخل الشهري / Monthly Income", "........................................ dollars / دولار")
    ]
    for r_idx, row_data in enumerate(living_data):
        row = living_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    doc.add_page_break()

    # PAGE 3
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")
    add_section_header(doc, "القسم الرابع: تحليل المخاطر ومخاوف الحماية", "Section 4: Risk Analysis & Protection Concerns")

    p_risk_lbl = doc.add_paragraph()
    run_risk = p_risk_lbl.add_run("يرجى تقديم وصف تفصيلي لمخاطر الحماية التي يواجهها المستفيد وأسرته / Please provide a detailed description of the protection risks faced by the beneficiary and their family:")
    run_risk.bold = True
    run_risk.font.name = 'Arial'
    run_risk.font.size = Pt(10)

    risk_box = doc.add_table(rows=1, cols=1)
    risk_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    risk_box.rows[0].cells[0].width = Inches(7.2)
    set_cell_borders(risk_box.rows[0].cells[0])
    set_cell_margins(risk_box.rows[0].cells[0], top=600, bottom=600, left=150, right=150)
    p_rbox = risk_box.rows[0].cells[0].paragraphs[0]
    p_rbox.text = "\n\n\n\n\n\n"

    add_section_header(doc, "مخاوف ومخاطر الحماية المحددة", "Specific Protection Risks & Concerns")

    risk_table = doc.add_table(rows=4, cols=2)
    risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    risk_table.autofit = False

    risk_chk_data = [
        ("[  ] تهديدات الأمن والسلامة / Safety & Security Threats", "[  ] الضيق النفسي والاجتماعي / Psychosocial distress"),
        ("[  ] عمالة الأطفال / Child labor", "[  ] نقص الوثائق المدنية / Lack of civil documentation"),
        ("[  ] خطر الإخلاء / Risk of eviction", "[  ] التمييز والاستبعاد / Discrimination & exclusion"),
        ("[  ] العنف القائم على النوع الاجتماعي / Gender-based violence", "[  ] مخاطر أخرى / Other risks: ........................................")
    ]
    for r_idx, row_data in enumerate(risk_chk_data):
        row = risk_table.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Inches(3.6)
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)

    doc.add_page_break()

    # PAGE 4
    add_header_table(doc, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")
    add_section_header(doc, "القسم الخامس: خطة العمل والإحالات", "Section 5: Action Plan & Referrals")

    action_table = doc.add_table(rows=4, cols=2)
    action_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    action_table.autofit = False

    action_data = [
        ("المساعدة الموصى بها / Recommended Assistance", "[  ] خدمات الحماية / Protection services   [  ] المساعدة النقدية / Cash support\n[  ] دعم المأوى / Shelter support          [  ] المساعدة القانونية / Legal support"),
        ("إحالة داخلية / Internal Referral", "إلى قسم: ................................................................................"),
        ("إحالة خارجية / External Referral", "إلى منظمة/جهة: ........................................................................"),
        ("الإجراء المتخذ فورا / Immediate Action Taken", "........................................................................................................")
    ]
    for r_idx, row_data in enumerate(action_data):
        row = action_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    add_section_header(doc, "التوقيعات والموافقات", "Signatures & Approvals")

    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False

    sig_data = [
        ("اسم وتوقيع الموظف المقيم / Assessor Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025",
         "اسم وتوقيع مشرف الحماية / Protection Supervisor Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]
    for r_idx, row_data in enumerate(sig_data):
        row = sig_table.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Inches(3.6)
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=150, bottom=150, left=150, right=150)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)

    # Save to dynamic buffer
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()

def generate_form2_docx():
    doc = Document()
    add_header_table(doc, "نموذج دراسة الحالة وإدارة الحالة", "Protection Case Study and Case Management Form")

    # PAGE 1
    add_section_header(doc, "القسم الأول: تسجيل الحالة والمعلومات العامة", "Section 1: Case Registration & General Information")

    table_reg = doc.add_table(rows=5, cols=4)
    table_reg.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_reg.autofit = False

    reg_data = [
        ("كود الحالة / Case Code", "........................................", "تاريخ فتح الحالة / Case Opening Date", "______/______/2025"),
        ("مسؤول الحالة / Case Worker", "........................................", "مصدر الإحالة / Referral Source", "........................................"),
        ("الاسم الكامل للمستفيد / Full Name", "................................................................................", "العمر / Age", ".................. years / سنة"),
        ("الجنس / Gender", "[  ] ذكر / Male      [  ] أنثى / Female", "رقم الهاتف / Phone Number", "........................................"),
        ("العنوان الحالي / Current Address", "................................................................................", "نوع الوثيقة القانونية / Legal Document", "........................................")
    ]
    col_widths = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8)]

    for r_idx, row_data in enumerate(reg_data):
        row = table_reg.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx % 2 == 0 else WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)
                if c_idx % 2 == 0:
                    run.bold = True
                    set_cell_background(cell, "F2F2F2")

    add_section_header(doc, "القسم الثاني: وصف وتقييم الحالة", "Section 2: Case Description & Assessment")

    p_desc_lbl = doc.add_paragraph()
    run_desc_lbl = p_desc_lbl.add_run("الخلفية التفصيلية للحالة ومخاطر الحماية الفورية / Detailed background of the case and immediate protection risks:")
    run_desc_lbl.bold = True
    run_desc_lbl.font.name = 'Arial'
    run_desc_lbl.font.size = Pt(10)

    case_desc_box = doc.add_table(rows=1, cols=1)
    case_desc_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    case_desc_box.rows[0].cells[0].width = Inches(7.2)
    set_cell_borders(case_desc_box.rows[0].cells[0])
    set_cell_margins(case_desc_box.rows[0].cells[0], top=600, bottom=600, left=150, right=150)
    p_cdbox = case_desc_box.rows[0].cells[0].paragraphs[0]
    p_cdbox.text = "\n\n\n\n\n\n"

    doc.add_page_break()

    # PAGE 2
    add_header_table(doc, "نموذج دراسة الحالة وإدارة الحالة", "Protection Case Study and Case Management Form")
    add_section_header(doc, "القسم الثالث: خطة عمل الحالة", "Section 3: Case Action Plan")

    plan_table = doc.add_table(rows=4, cols=5)
    plan_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    plan_table.autofit = False

    headers = [
        "الهدف / Objective",
        "الإجراء المطلوب / Action Required",
        "الشخص المسؤول / Responsible",
        "الإطار الزمني / Timeframe",
        "الحالة / Status"
    ]
    # Header row
    for c_idx, h_text in enumerate(headers):
        cell = plan_table.rows[0].cells[c_idx]
        cell.width = Inches(1.44)
        cell.text = h_text
        set_cell_borders(cell)
        set_cell_background(cell, "003366")
        set_cell_margins(cell, top=100, bottom=100, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)
            r.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

    # Empty rows
    for r_idx in range(1, 4):
        row = plan_table.rows[r_idx]
        for c_idx in range(5):
            cell = row.cells[c_idx]
            cell.width = Inches(1.44)
            cell.text = ""
            set_cell_borders(cell)
            set_cell_margins(cell, top=300, bottom=300, left=80, right=80)

    add_section_header(doc, "القسم الرابع: المتابعة وإغلاق الحالة", "Section 4: Follow-up & Case Closure")

    close_table = doc.add_table(rows=3, cols=2)
    close_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    close_table.autofit = False

    close_data = [
        ("ملخص زيارات المتابعة / Summary of follow-up visits", "................................................................................................................................................"),
        ("حالة الحالة / Case Status", "[  ] نشطة / Active         [  ] مغلقة / Closed         [  ] معلقة / Suspended"),
        ("أسباب الإغلاق / Reasons for closure", "[  ] تم تحقيق الأهداف / Goals met    [  ] اكتمال الإحالة / Referral complete    [  ] انتقال / Relocation")
    ]
    for r_idx, row_data in enumerate(close_data):
        row = close_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    add_section_header(doc, "التوقيعات والموافقات", "Signatures & Approvals")

    sig_table2 = doc.add_table(rows=2, cols=2)
    sig_table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table2.autofit = False

    sig_data2 = [
        ("اسم وتوقيع مسؤول الحالة / Case Worker Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025",
         "اسم وتوقيع مشرف إدارة الحالة / Supervisor Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]
    for r_idx, row_data in enumerate(sig_data2):
        row = sig_table2.rows[r_idx]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Inches(3.6)
            cell.text = text
            set_cell_borders(cell)
            set_cell_margins(cell, top=150, bottom=150, left=150, right=150)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()

def generate_form3_docx():
    doc = Document()
    add_header_table(doc, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    # PAGE 1
    add_section_header(doc, "المعلومات الأساسية", "Basic Information")

    info_table = doc.add_table(rows=4, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False

    info_data = [
        ("معرّف الحالة / Case ID", "................................................................................"),
        ("تاريخ التقييم / Evaluation Date", "2025 / ______ / ______"),
        ("تاريخ قرار الموافقة / Approval Decision Date", "2025 / ______ / ______"),
        ("نوع الخطر الحمائي / Protection Risk Type", "................................................................................")
    ]
    for r_idx, row_data in enumerate(info_data):
        row = info_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    add_section_header(doc, "القسم الأول: ملخص الحالة", "Section 1: Case Summary")

    p_bdata = doc.add_paragraph()
    p_bdata.add_run("1.1 البيانات الأساسية للمستفيد / Beneficiary Basic Data").bold = True
    p_bdata.runs[0].font.name = 'Arial'
    p_bdata.runs[0].font.size = Pt(10)

    bdata_table = doc.add_table(rows=4, cols=2)
    bdata_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    bdata_table.autofit = False

    bdata_rows = [
        ("الاسم أو معرف / Name or ID", "................................................................................"),
        ("العمر / Age", ".................. سنة / years"),
        ("الجنس / Gender", "[  ] ذكر / Male         [  ] أنثى / Female"),
        ("عدد أفراد الأسرة / Family Members", ".................. أفراد / members")
    ]
    for r_idx, row_data in enumerate(bdata_rows):
        row = bdata_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    # Risk Summary
    p_rsum = doc.add_paragraph()
    p_rsum.paragraph_format.space_before = Pt(8)
    p_rsum.add_run("1.2 ملخص الخطر / Risk Summary").bold = True
    p_rsum.runs[0].font.name = 'Arial'
    p_rsum.runs[0].font.size = Pt(10)

    risk_table = doc.add_table(rows=2, cols=2)
    risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    risk_table.autofit = False

    risk_rows = [
        ("الخطر الأساسي / Primary Risk", "\n\n\n"),
        ("الشدة / Severity", "[  ] حاد وفوري جداً (24-48 ساعة) / Very acute and immediate (24-48 hours)\n[  ] حاد (3-5 أيام) / Acute (3-5 days)\n[  ] متوسط (أسبوع - أسبوعين) / Moderate (1-2 weeks)")
    ]
    for r_idx, row_data in enumerate(risk_rows):
        row = risk_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    doc.add_page_break()

    # PAGE 2
    add_header_table(doc, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    # 1.3 Proposed Solution and Amount
    add_section_header(doc, "1.3 الحل المقترح والمبلغ", "Proposed Solution and Amount")

    sol_table = doc.add_table(rows=3, cols=2)
    sol_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sol_table.autofit = False

    sol_rows = [
        ("كيف سيساعد النقد؟ / How will cash help?", "\n\n"),
        ("المبلغ المقترح / Proposed Amount", "........................................ دولار / dollars"),
        ("السبب / Reason", "\n\n")
    ]
    for r_idx, row_data in enumerate(sol_rows):
        row = sol_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    add_section_header(doc, "القسم الثاني: موافقة قسم الحماية", "Section 2: Protection Department Approval")

    p_review = doc.add_paragraph()
    p_review.add_run("2.1 مراجعة التقييم / Evaluation Review").bold = True
    p_review.runs[0].font.name = 'Arial'
    p_review.runs[0].font.size = Pt(10)

    review_table = doc.add_table(rows=4, cols=2)
    review_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    review_table.autofit = False

    review_rows = [
        ("وضوح الخطر / Risk Clarity", "[  ] واضح جداً / Very Clear      [  ] واضح / Clear      [  ] غامض / Vague"),
        ("الربط السببي / Causal Link", "[  ] منطقي جداً / Very Logical   [  ] منطقي / Logical   [  ] ضعيف / Weak"),
        ("كفاية التقييم / Evaluation Adequacy", "[  ] شامل / Comprehensive         [  ] متوسط / Moderate     [  ] ناقص / Incomplete"),
        ("المخاطر الثانوية / Secondary Risks", "[  ] منخفضة / Low               [  ] متوسطة / Medium     [  ] عالية / High")
    ]
    for r_idx, row_data in enumerate(review_rows):
        row = review_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    add_section_header(doc, "القسم الثالث: القرار النهائي من مدير الحماية", "Section 3: Final Decision from Protection Manager")

    mgr_table = doc.add_table(rows=3, cols=2)
    mgr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    mgr_table.autofit = False

    mgr_rows = [
        ("القرار النهائي للمدير / Final Decision", "[  ] موافق على الصرف / Approved       [  ] غير موافق / Not Approved"),
        ("المبلغ وآلية الصرف / Amount & Mechanism", "المبلغ: ........................ دولار    الآلية: [  ] نقد / Cash   [  ] حوالة / Transfer   [  ] دفع مباشر"),
        ("التاريخ والتوقيع / Date & Signature", "الاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]
    for r_idx, row_data in enumerate(mgr_rows):
        row = mgr_table.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(5.0)

        cell_lbl.text = row_data[0]
        cell_val.text = row_data[1]

        for c in [cell_lbl, cell_val]:
            set_cell_borders(c)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_background(cell_lbl, "F2F2F2")
        for r in p_lbl.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p_val.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()
