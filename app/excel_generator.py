import io
import os
import xlsxwriter

def create_excel_styles(workbook):
    styles = {}

    # Base configuration
    styles['title_ar'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#003366'
    })
    styles['title_en'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#003366'
    })
    styles['header_org_ar'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 11,
        'align': 'right',
        'valign': 'vcenter',
        'font_color': '#003366'
    })
    styles['header_org_en'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 9.5,
        'align': 'right',
        'valign': 'vcenter',
        'font_color': '#003366'
    })
    styles['header_org_dept'] = workbook.add_format({
        'italic': True,
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 9,
        'align': 'right',
        'valign': 'vcenter',
        'font_color': '#787878'
    })
    styles['section'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 11,
        'bg_color': '#E6F0FA',
        'font_color': '#003366',
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#CCCCCC'
    })
    styles['label'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 9.5,
        'bg_color': '#F2F2F2',
        'align': 'right',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3',
        'text_wrap': True
    })
    styles['label_left'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 9.5,
        'bg_color': '#F2F2F2',
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3',
        'text_wrap': True
    })
    styles['value'] = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 9.5,
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3',
        'text_wrap': True
    })
    styles['value_bold'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 9.5,
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3',
        'text_wrap': True
    })
    styles['table_header'] = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 9.5,
        'bg_color': '#003366',
        'font_color': '#FFFFFF',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#CCCCCC'
    })
    return styles

def apply_page_setup(worksheet):
    worksheet.set_paper(9) # A4
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    worksheet.hide_gridlines(0) # Ensure gridlines are printed and visible
    worksheet.fit_to_pages(1, 0)
    worksheet.right_to_left()  # RTL configuration to ensure labels are on the right

def add_excel_header(workbook, worksheet, styles, title_ar, title_en):
    # Set default row heights for header block
    worksheet.set_row(0, 18)
    worksheet.set_row(1, 15)
    worksheet.set_row(2, 15)
    worksheet.set_row(4, 22)
    worksheet.set_row(5, 18)

    # Left organization text stacked on the right side of Columns A-D
    worksheet.merge_range('A1:C1', "جمعية جيل البناء للتنمية الإنسانية", styles['header_org_ar'])
    worksheet.merge_range('A2:C2', "Jeel Al Bena Association for Humanitarian Development", styles['header_org_en'])
    worksheet.merge_range('A3:C3', "قسم الحماية / Protection Department", styles['header_org_dept'])

    # Right logo place (embed picture if exists)
    logo_path = 'app/static/images/logo_emblem.png'
    if not os.path.exists(logo_path):
        logo_path = 'app/static/images/emblem.png'

    if os.path.exists(logo_path):
        worksheet.insert_image('D1', logo_path, {
            'x_scale': 0.45,
            'y_scale': 0.45,
            'x_offset': 40,
            'y_offset': 3
        })
    else:
        worksheet.merge_range('D1:D3', "[ Logo / شعار ]", styles['title_en'])

    # Blue underline separator
    border_format = workbook.add_format({'bottom': 6, 'bottom_color': '#003366'})
    worksheet.write('A4', '', border_format)
    worksheet.write('B4', '', border_format)
    worksheet.write('C4', '', border_format)
    worksheet.write('D4', '', border_format)

    # Main titles centered
    worksheet.merge_range('A5:D5', title_ar, styles['title_ar'])
    worksheet.merge_range('A6:D6', title_en, styles['title_en'])

def generate_form1_xlsx():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    styles = create_excel_styles(workbook)

    # Form 1 Page
    ws = workbook.add_worksheet("Individual Assessment Form")
    apply_page_setup(ws)

    # Column Widths
    ws.set_column('A:A', 28)
    ws.set_column('B:B', 24)
    ws.set_column('C:C', 28)
    ws.set_column('D:D', 24)

    add_excel_header(workbook, ws, styles, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    row = 7
    # Section: Basic Information
    ws.merge_range(row, 0, row, 3, "  المعلومات الأساسية / Basic Information", styles['section'])
    ws.set_row(row, 24)
    row += 1

    basic_data = [
        ("كود الحالة / Case Code", "........................................", "تاريخ التقييم / Assessment Date", "2025 / ______ / ______"),
        ("اسم الموظف / Staff Name", "........................................", "المنظمة / Organization", "........................................"),
        ("المحافظة / Province", "........................................", "المديرية / District", "........................................"),
        ("المركز / Center", "........................................", "الموقع / Location", "........................................"),
        ("اسم المستفيد / Beneficiary Name", "........................................", "رقم الهاتف / Phone Number", "........................................"),
        ("العمر / Age", ".................. (years / سنة)", "الجنس / Gender", "☐ ذكر / Male      ☐ أنثى / Female")
    ]

    for r in basic_data:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.write(row, 1, r[1], styles['value'])
        ws.write(row, 2, r[2], styles['label'])
        ws.write(row, 3, r[3], styles['value'])
        row += 1

    row += 1
    # Section 1: Demographics
    ws.merge_range(row, 0, row, 3, "  القسم الأول: البيانات الديموغرافية / Section 1: Demographics", styles['section'])
    ws.set_row(row, 24)
    row += 1

    demo_data = [
        ("الحالة الاجتماعية / Marital Status", "☐ أعزب / Single   ☐ متزوج / Married\n☐ أرمل / Widowed   ☐ مطلق / Divorced", "عدد أفراد الأسرة / Family Size", ".................. (members / أفراد)"),
        ("حالة النزوح / Displacement Status", "☐ نازح / IDP   ☐ عائد / Returnee\n☐ مجتمع مضيف / Host", "تاريخ النزوح / Date of Displacement", "______ / ______ / 2025"),
        ("العنوان الحالي / Current Address", "................................................................................", "العنوان الأصلي / Original Address", "................................................................................"),
        ("حالة التسجيل / Registration Status", "☐ مسجل / Registered   ☐ غير مسجل / Unregistered", "الاحتياجات الخاصة / Specific Needs", "☐ نعم / Yes   ☐ لا / No")
    ]

    for r in demo_data:
        ws.set_row(row, 28 if "\n" in r[1] else 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.write(row, 1, r[1], styles['value'])
        ws.write(row, 2, r[2], styles['label'])
        ws.write(row, 3, r[3], styles['value'])
        row += 1

    row += 1
    # Section 2: Specific Vulnerabilities
    ws.merge_range(row, 0, row, 3, "  القسم الثاني: نقاط الضعف الخاصة / Section 2: Specific Vulnerabilities", styles['section'])
    ws.set_row(row, 24)
    row += 1

    vuln_data = [
        ("☐ طفل معرض للخطر / Child at risk", "☐ عائل وحيد / Single parent"),
        ("☐ مسن معرض للخطر / Older person at risk", "☐ حالة طبية حرجة / Serious medical condition"),
        ("☐ إعاقة جسدية / Physical disability", "☐ إعاقة ذهنية / Mental disability"),
        ("☐ امرأة حامل أو مرضعة / Pregnant or lactating", "☐ امرأة معرضة للخطر / Woman at risk"),
        ("☐ ناجٍ من العنف / Survivor of violence", "☐ غير ذلك (حدد) / Other (Specify): ........................................")
    ]

    for r in vuln_data:
        ws.set_row(row, 22)
        ws.merge_range(row, 0, row, 1, r[0], styles['value_bold'])
        ws.merge_range(row, 2, row, 3, r[1], styles['value_bold'])
        row += 1

    row += 1
    # Section 3: Assessment Details
    ws.merge_range(row, 0, row, 3, "  القسم الثالث: تفاصيل التقييم وتاريخ النزوح / Section 3: Assessment Details & Displacement History", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row, 3, "أسباب النزوح والوضع الحالي ومصادر العيش الرئيسية للأسرة / Reasons for displacement and current situation:", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    ws.merge_range(row, 0, row + 4, 3, "\n\n\n\n", styles['value'])
    ws.set_row(row, 100)
    row += 5

    # Living Conditions
    ws.merge_range(row, 0, row, 3, "  الظروف المعيشية والسكن / Living Conditions & Accommodation", styles['section'])
    ws.set_row(row, 24)
    row += 1

    living_data = [
        ("نوع السكن / Accommodation Type", "☐ إيجار / Rented     ☐ مع عائلة مضيفة / Host Family     ☐ خيمة / Tent     ☐ منزل متضرر / Damaged house"),
        ("مصدر الدخل / Source of Income", "................................................................................"),
        ("الدخل الشهري / Monthly Income", "........................................ dollars / دولار")
    ]
    for r in living_data:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    # Section 4: Risk Analysis
    ws.merge_range(row, 0, row, 3, "  القسم الرابع: تحليل المخاطر ومخاوف الحماية / Section 4: Risk Analysis & Protection Concerns", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row, 3, "يرجى تقديم وصف تفصيلي لمخاطر الحماية التي يواجهها المستفيد وأسرته / Detailed protection risks description:", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    ws.merge_range(row, 0, row + 4, 3, "\n\n\n\n", styles['value'])
    ws.set_row(row, 100)
    row += 5

    # Protection Concerns
    ws.merge_range(row, 0, row, 3, "مخاوف ومخاطر الحماية المحددة / Specific Protection Risks & Concerns", styles['section'])
    ws.set_row(row, 24)
    row += 1

    risk_chk_data = [
        ("☐ تهديدات الأمن والسلامة / Safety & Security Threats", "☐ الضيق النفسي والاجتماعي / Psychosocial distress"),
        ("☐ عمالة الأطفال / Child labor", "☐ نقص الوثائق المدنية / Lack of civil documentation"),
        ("☐ خطر الإخلاء / Risk of eviction", "☐ التمييز والاستبعاد / Discrimination & exclusion"),
        ("☐ العنف القائم على النوع الاجتماعي / Gender-based violence", "☐ مخاطر أخرى / Other risks: ........................................")
    ]
    for r in risk_chk_data:
        ws.set_row(row, 22)
        ws.merge_range(row, 0, row, 1, r[0], styles['value_bold'])
        ws.merge_range(row, 2, row, 3, r[1], styles['value_bold'])
        row += 1

    row += 1
    # Section 5: Action Plan & Referrals
    ws.merge_range(row, 0, row, 3, "  القسم الخامس: خطة العمل والإحالات / Section 5: Action Plan & Referrals", styles['section'])
    ws.set_row(row, 24)
    row += 1

    action_data = [
        ("المساعدة الموصى بها / Recommended Assistance", "☐ خدمات الحماية / Protection services   ☐ المساعدة النقدية / Cash support\n☐ دعم المأوى / Shelter support          ☐ المساعدة القانونية / Legal support"),
        ("إحالة داخلية / Internal Referral", "إلى قسم: ................................................................................"),
        ("إحالة خارجية / External Referral", "إلى منظمة/جهة: ........................................................................"),
        ("الإجراء المتخذ فورا / Immediate Action Taken", "........................................................................................................")
    ]
    for r in action_data:
        ws.set_row(row, 28 if "\n" in r[1] else 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    # Signatures
    ws.merge_range(row, 0, row, 3, "التوقيعات والموافقات / Signatures & Approvals", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row + 4, 1, "اسم وتوقيع الموظف المقيم / Assessor Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025", styles['value_bold'])
    ws.merge_range(row, 2, row + 4, 3, "اسم وتوقيع مشرف الحماية / Protection Supervisor Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025", styles['value_bold'])
    ws.set_row(row, 100)
    row += 5

    workbook.close()
    output.seek(0)
    return output.read()

def generate_form2_xlsx():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    styles = create_excel_styles(workbook)

    # Form 2 Page
    ws = workbook.add_worksheet("Protection Case Management")
    apply_page_setup(ws)

    # Use 5 Columns to perfectly match the Action Plan table
    ws.set_column('A:A', 21)
    ws.set_column('B:B', 21)
    ws.set_column('C:C', 21)
    ws.set_column('D:D', 21)
    ws.set_column('E:E', 21)

    # Custom Header stack with 5 merged cells
    ws.set_row(0, 18)
    ws.set_row(1, 15)
    ws.set_row(2, 15)
    ws.set_row(4, 22)
    ws.set_row(5, 18)
    ws.merge_range('A1:D1', "جمعية جيل البناء للتنمية الإنسانية", styles['header_org_ar'])
    ws.merge_range('A2:D2', "Jeel Al Bena Association for Humanitarian Development", styles['header_org_en'])
    ws.merge_range('A3:D3', "قسم الحماية / Protection Department", styles['header_org_dept'])

    # Emblem Insert
    logo_path = 'app/static/images/logo_emblem.png'
    if not os.path.exists(logo_path):
        logo_path = 'app/static/images/emblem.png'

    if os.path.exists(logo_path):
        ws.insert_image('E1', logo_path, {
            'x_scale': 0.45,
            'y_scale': 0.45,
            'x_offset': 10,
            'y_offset': 3
        })
    else:
        ws.write('E1', "[ Logo ]", styles['title_en'])

    border_format = workbook.add_format({'bottom': 6, 'bottom_color': '#003366'})
    for c_i in range(5):
        ws.write(3, c_i, '', border_format)

    ws.merge_range('A5:E5', "نموذج دراسة الحالة وإدارة الحالة", styles['title_ar'])
    ws.merge_range('A6:E6', "Protection Case Study and Case Management Form", styles['title_en'])

    row = 7
    # Section 1: Registration
    ws.merge_range(row, 0, row, 4, "  القسم الأول: تسجيل الحالة والمعلومات العامة / Section 1: Case Registration", styles['section'])
    ws.set_row(row, 24)
    row += 1

    reg_data = [
        ("كود الحالة / Case Code", "........................................", "تاريخ فتح الحالة / Case Opening Date", "______/______/2025"),
        ("مسؤول الحالة / Case Worker", "........................................", "مصدر الإحالة / Referral Source", "........................................"),
        ("الاسم الكامل للمستفيد / Full Name", "................................................................................", "العمر / Age", ".................. years / سنة"),
        ("الجنس / Gender", "☐ ذكر / Male      ☐ أنثى / Female", "رقم الهاتف / Phone Number", "........................................"),
        ("العنوان الحالي / Current Address", "................................................................................", "نوع الوثيقة القانونية / Legal Document", "........................................")
    ]

    for r in reg_data:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.write(row, 1, r[1], styles['value'])
        ws.write(row, 2, r[2], styles['label'])
        ws.merge_range(row, 3, row, 4, r[3], styles['value'])
        row += 1

    row += 1
    # Section 2: Case Description
    ws.merge_range(row, 0, row, 4, "  القسم الثاني: وصف وتقييم الحالة / Section 2: Case Description & Assessment", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row, 4, "الخلفية التفصيلية للحالة ومخاطر الحماية الفورية والوضع الاجتماعي والاقتصادي / Detailed background of the case:", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    ws.merge_range(row, 0, row + 5, 4, "\n\n\n\n\n\n", styles['value'])
    ws.set_row(row, 120)
    row += 6

    row += 1
    # Section 3: Action Plan
    ws.merge_range(row, 0, row, 4, "  القسم الثالث: خطة عمل الحالة / Section 3: Case Action Plan", styles['section'])
    ws.set_row(row, 24)
    row += 1

    # 5 Column Table headers
    ws.set_row(row, 22)
    ws.write(row, 0, "الهدف / Objective", styles['table_header'])
    ws.write(row, 1, "الإجراء المطلوب / Action Required", styles['table_header'])
    ws.write(row, 2, "الشخص المسؤول / Responsible", styles['table_header'])
    ws.write(row, 3, "الإطار الزمني / Timeframe", styles['table_header'])
    ws.write(row, 4, "الحالة / Status", styles['table_header'])
    row += 1

    # Empty Rows
    for _ in range(3):
        ws.set_row(row, 30)
        for c_idx in range(5):
            ws.write(row, c_idx, "", styles['value'])
        row += 1

    row += 1
    # Section 4: Follow-up & Case Closure
    ws.merge_range(row, 0, row, 4, "  القسم الرابع: المتابعة وإغلاق الحالة / Section 4: Follow-up & Case Closure", styles['section'])
    ws.set_row(row, 24)
    row += 1

    close_data = [
        ("ملخص زيارات المتابعة وتحديثاتها / Follow-up summary", "................................................................................................................................................"),
        ("حالة الحالة / Case Status", "☐ نشطة / Active         ☐ مغلقة / Closed         ☐ معلقة / Suspended"),
        ("أسباب الإغلاق / Reasons for closure", "☐ تم تحقيق الأهداف / Goals met    ☐ اكتمال الإحالة / Referral complete    ☐ انتقال / Relocation")
    ]

    for r in close_data:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 4, r[1], styles['value'])
        row += 1

    row += 1
    # Section 5: Signatures & Approvals
    ws.merge_range(row, 0, row, 4, "التوقيعات والموافقات / Signatures & Approvals", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row + 4, 1, "اسم وتوقيع مسؤول الحالة / Case Worker Name & Signature\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025", styles['value_bold'])
    ws.merge_range(row, 2, row + 4, 4, "اسم وتوقيع مشرف إدارة الحالة / Case Management Supervisor\n\nالاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025", styles['value_bold'])
    ws.set_row(row, 100)
    row += 5

    workbook.close()
    output.seek(0)
    return output.read()

def generate_form3_xlsx():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    styles = create_excel_styles(workbook)

    # Form 3 Page
    ws = workbook.add_worksheet("Cash for Protection")
    apply_page_setup(ws)

    # Columns Widths
    ws.set_column('A:A', 28)
    ws.set_column('B:B', 24)
    ws.set_column('C:C', 28)
    ws.set_column('D:D', 24)

    add_excel_header(workbook, ws, styles, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    row = 7
    # Basic Information
    ws.merge_range(row, 0, row, 3, "  المعلومات الأساسية / Basic Information", styles['section'])
    ws.set_row(row, 24)
    row += 1

    info_data = [
        ("معرّف الحالة / Case ID", "................................................................................"),
        ("تاريخ التقييم / Evaluation Date", "2025 / ______ / ______"),
        ("تاريخ قرار الموافقة / Approval Decision Date", "2025 / ______ / ______"),
        ("نوع الخطر الحمائي / Protection Risk Type", "................................................................................")
    ]
    for r in info_data:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    # Section 1: Case Summary
    ws.merge_range(row, 0, row, 3, "  القسم الأول: ملخص الحالة / Section 1: Case Summary", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row, 3, "1.1 البيانات الأساسية للمستفيد / Beneficiary Basic Data", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    bdata_rows = [
        ("الاسم أو معرف / Name or ID", "................................................................................"),
        ("العمر / Age", ".................. سنة / years"),
        ("الجنس / Gender", "☐ ذكر / Male         ☐ أنثى / Female"),
        ("عدد أفراد الأسرة / Family Members", ".................. أفراد / members")
    ]
    for r in bdata_rows:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    ws.merge_range(row, 0, row, 3, "1.2 ملخص الخطر / Risk Summary", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    risk_rows = [
        ("الخطر الأساسي / Primary Risk", "\n\n\n"),
        ("الشدة / Severity", "☐ حاد وفوري جداً (24-48 ساعة) / Very acute and immediate (24-48 hours)\n☐ حاد (3-5 أيام) / Acute (3-5 days)\n☐ متوسط (أسبوع - أسبوعين) / Moderate (1-2 weeks)")
    ]
    for r in risk_rows:
        ws.set_row(row, 40 if "\n" in r[1] else 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    ws.merge_range(row, 0, row, 3, "1.3 الحل المقترح والمبلغ / Proposed Solution and Amount", styles['section'])
    ws.set_row(row, 24)
    row += 1

    sol_rows = [
        ("كيف سيساعد النقد؟ / How will cash help?", "\n\n"),
        ("المبلغ المقترح / Proposed Amount", "........................................ دولار / dollars"),
        ("السبب / Reason", "\n\n")
    ]
    for r in sol_rows:
        ws.set_row(row, 40 if "\n" in r[1] else 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    # Section 2: Protection Approval
    ws.merge_range(row, 0, row, 3, "  القسم الثاني: موافقة قسم الحماية / Section 2: Protection Department Approval", styles['section'])
    ws.set_row(row, 24)
    row += 1

    ws.merge_range(row, 0, row, 3, "2.1 مراجعة التقييم / Evaluation Review", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    review_rows = [
        ("وضوح الخطر / Risk Clarity", "☐ واضح جداً / Very Clear      ☐ واضح / Clear      ☐ غامض / Vague"),
        ("الربط السببي / Causal Link", "☐ منطقي جداً / Very Logical   ☐ منطقي / Logical   ☐ ضعيف / Weak"),
        ("كفاية التقييم / Evaluation Adequacy", "☐ شامل / Comprehensive         ☐ متوسط / Moderate     ☐ ناقص / Incomplete"),
        ("المخاطر الثانوية / Secondary Risks", "☐ منخفضة / Low               ☐ متوسطة / Medium     ☐ عالية / High")
    ]
    for r in review_rows:
        ws.set_row(row, 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    row += 1
    ws.merge_range(row, 0, row, 3, "2.2 قرار ضابط الحماية / Protection Officer's Decision", styles['label_left'])
    ws.set_row(row, 22)
    row += 1

    # Complex detailed cell with checkboxes in the officer section
    ws.merge_range(row, 0, row + 5, 3, "هل توصي بالموافقة على صرف النقد؟ / Do you recommend approval for cash disbursement?\n" +
                   "☑ نعم، موافق تماماً / Yes, fully approve     " +
                   "☐ نعم، مع شروط / Yes, with conditions:\n" +
                   "الشروط / Conditions: ............................................................................................................\n" +
                   "☐ لا، غير موافق / No, do not approve\n" +
                   "السبب / Reason: ................................................................................................................\n" +
                   "اسم وتوقيع ضابط الحماية / Protection Officer: ..................................................................................", styles['value_bold'])
    ws.set_row(row, 120)
    row += 6

    row += 1
    # Section 3: Final Decision
    ws.merge_range(row, 0, row, 3, "  القسم الثالث: القرار النهائي من مدير الحماية / Section 3: Final Decision", styles['section'])
    ws.set_row(row, 24)
    row += 1

    mgr_rows = [
        ("القرار النهائي للمدير / Final Decision", "☐ موافق على الصرف / Approved       ☐ غير موافق / Not Approved"),
        ("المبلغ وآلية الصرف / Amount & Mechanism", "المبلغ: ........................ دولار    الآلية: ☐ نقد / Cash   ☐ حوالة / Transfer   ☐ دفع مباشر / Direct Payment"),
        ("التوقيع والتاريخ / Signature & Date", "الاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]
    for r in mgr_rows:
        ws.set_row(row, 40 if "\n" in r[0] or "\n" in r[1] else 22)
        ws.write(row, 0, r[0], styles['label'])
        ws.merge_range(row, 1, row, 3, r[1], styles['value'])
        row += 1

    workbook.close()
    output.seek(0)
    return output.read()
