import io
import xlsxwriter

def create_excel_styles(workbook):
    styles = {}
    styles['title_ar'] = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#003366'
    })
    styles['title_en'] = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#003366'
    })
    styles['header_org_ar'] = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 10,
        'align': 'left',
        'valign': 'vcenter'
    })
    styles['header_org_en'] = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 9,
        'align': 'left',
        'valign': 'vcenter'
    })
    styles['section'] = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 11,
        'bg_color': '#E6F0FA',
        'font_color': '#003366',
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3'
    })
    styles['label'] = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 9.5,
        'bg_color': '#F2F2F2',
        'align': 'right',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3',
        'text_wrap': True
    })
    styles['value'] = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 9.5,
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3',
        'text_wrap': True
    })
    styles['table_header'] = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 9.5,
        'bg_color': '#003366',
        'font_color': '#FFFFFF',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#D3D3D3'
    })
    return styles

def apply_page_setup(worksheet):
    worksheet.set_paper(9) # A4
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    worksheet.fit_to_pages(1, 0)

def add_excel_header(worksheet, styles, title_ar, title_en):
    worksheet.set_row(0, 18)
    worksheet.set_row(1, 15)
    worksheet.set_row(2, 15)
    worksheet.set_row(4, 20)
    worksheet.set_row(5, 18)

    # Left organization text
    worksheet.merge_range('A1:C1', "جمعية جيل البناء للتنمية الإنسانية", styles['header_org_ar'])
    worksheet.merge_range('A2:C2', "Jeel Al Bena Association for Humanitarian Development", styles['header_org_en'])
    worksheet.merge_range('A3:C3', "قسم الحماية / Protection Department", styles['header_org_en'])

    # Right logo place (we can write "LOGO" or embed picture)
    worksheet.merge_range('D1:E3', "[ Emblem / شعار ]", styles['title_en'])

    # Main title
    worksheet.merge_range('A5:E5', title_ar, styles['title_ar'])
    worksheet.merge_range('A6:E6', title_en, styles['title_en'])

def generate_form1_xlsx():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    styles = create_excel_styles(workbook)

    # Form 1 page 1
    ws = workbook.add_worksheet("Form 1 - Page 1")
    apply_page_setup(ws)
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 22)
    ws.set_column('D:D', 22)
    ws.set_column('E:E', 12)

    add_excel_header(ws, styles, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    # Basic info
    ws.merge_range('A8:E8', "  المعلومات الأساسية / Basic Information", styles['section'])
    ws.set_row(7, 24)

    basic_data = [
        ("كود الحالة / Case Code", "........................................", "تاريخ التقييم / Assessment Date", "2025 / ______ / ______"),
        ("اسم الموظف / Staff Name", "........................................", "المنظمة / Organization", "........................................"),
        ("المحافظة / Province", "........................................", "المديرية / District", "........................................"),
        ("المركز / Center", "........................................", "الموقع / Location", "........................................"),
        ("اسم المستفيد / Beneficiary Name", "........................................", "رقم الهاتف / Phone Number", "........................................"),
        ("العمر / Age", ".................. (years / سنة)", "الجنس / Gender", "[  ] ذكر / Male      [  ] أنثى / Female")
    ]

    row_num = 8
    for r in basic_data:
        ws.set_row(row_num, 20)
        ws.write(row_num, 0, r[0], styles['label'])
        ws.write(row_num, 1, r[1], styles['value'])
        ws.write(row_num, 2, r[2], styles['label'])
        ws.merge_range(row_num, 3, row_num, 4, r[3], styles['value'])
        row_num += 1

    row_num += 1
    # Section 1: Demographics
    ws.merge_range(f'A{row_num+1}:E{row_num+1}', "  القسم الأول: البيانات الديموغرافية / Section 1: Demographics", styles['section'])
    ws.set_row(row_num, 24)
    row_num += 1

    demo_data = [
        ("الحالة الاجتماعية / Marital Status", "[  ] أعزب / Single   [  ] متزوج / Married   [  ] أرمل / Widowed   [  ] مطلق / Divorced", "عدد أفراد الأسرة / Family Size", ".................. (members / أفراد)"),
        ("حالة النزوح / Displacement Status", "[  ] نازح / IDP   [  ] عائد / Returnee   [  ] مجتمع مضيف / Host", "تاريخ النزوح / Date of Displacement", "______ / ______ / 2025"),
        ("العنوان الحالي / Current Address", "................................................................................", "العنوان الأصلي / Original Address", "................................................................................"),
        ("حالة التسجيل / Registration Status", "[  ] مسجل / Registered   [  ] غير مسجل / Unregistered", "الاحتياجات الخاصة / Specific Needs", "[  ] نعم / Yes   [  ] لا / No")
    ]
    for r in demo_data:
        ws.set_row(row_num, 20)
        ws.write(row_num, 0, r[0], styles['label'])
        ws.write(row_num, 1, r[1], styles['value'])
        ws.write(row_num, 2, r[2], styles['label'])
        ws.merge_range(row_num, 3, row_num, 4, r[3], styles['value'])
        row_num += 1

    row_num += 1
    # Section 2: Specific Vulnerabilities
    ws.merge_range(f'A{row_num+1}:E{row_num+1}', "  القسم الثاني: نقاط الضعف الخاصة / Section 2: Specific Vulnerabilities", styles['section'])
    ws.set_row(row_num, 24)
    row_num += 1

    vuln_data = [
        ("[  ] طفل معرض للخطر / Child at risk", "[  ] عائل وحيد / Single parent"),
        ("[  ] مسن معرض للخطر / Older person at risk", "[  ] حالة طبية حرجة / Serious medical condition"),
        ("[  ] إعاقة جسدية / Physical disability", "[  ] إعاقة ذهنية / Mental disability"),
        ("[  ] امرأة حامل أو مرضعة / Pregnant or lactating", "[  ] امرأة معرضة للخطر / Woman at risk"),
        ("[  ] ناجٍ من العنف / Survivor of violence", "[  ] غير ذلك (حدد) / Other (Specify): ........................................")
    ]
    for r in vuln_data:
        ws.set_row(row_num, 20)
        ws.merge_range(row_num, 0, row_num, 1, r[0], styles['value'])
        ws.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    # Form 1 page 2
    ws2 = workbook.add_worksheet("Form 1 - Page 2")
    apply_page_setup(ws2)
    ws2.set_column('A:A', 22)
    ws2.set_column('B:B', 22)
    ws2.set_column('C:C', 22)
    ws2.set_column('D:D', 22)
    ws2.set_column('E:E', 12)
    add_excel_header(ws2, styles, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    ws2.merge_range('A8:E8', "  القسم الثالث: تفاصيل التقييم وتاريخ النزوح / Section 3: Assessment Details & Displacement History", styles['section'])
    ws2.set_row(7, 24)

    ws2.merge_range('A10:E10', "أسباب النزوح والوضع الحالي / Reasons for displacement and current situation:", styles['label'])
    ws2.merge_range('A11:E15', "\n\n\n\n", styles['value'])
    ws2.set_row(10, 80)

    ws2.merge_range('A17:E17', "  الظروف المعيشية والسكن / Living Conditions & Accommodation", styles['section'])
    ws2.set_row(16, 24)

    ws2.set_row(18, 20)
    ws2.merge_range('A19:B19', "نوع السكن / Accommodation Type", styles['label'])
    ws2.merge_range('C19:E19', "[  ] إيجار / Rented   [  ] مع عائلة مضيفة / Host Family   [  ] خيمة / Tent   [  ] منزل متضرر / Damaged house", styles['value'])

    ws2.set_row(19, 20)
    ws2.merge_range('A20:B20', "مصدر الدخل / Source of Income", styles['label'])
    ws2.merge_range('C20:E20', "................................................................................", styles['value'])

    ws2.set_row(20, 20)
    ws2.merge_range('A21:B21', "الدخل الشهري / Monthly Income", styles['label'])
    ws2.merge_range('C21:E21', "........................................ dollars / دولار", styles['value'])

    # Form 1 page 3
    ws3 = workbook.add_worksheet("Form 1 - Page 3")
    apply_page_setup(ws3)
    ws3.set_column('A:A', 22)
    ws3.set_column('B:B', 22)
    ws3.set_column('C:C', 22)
    ws3.set_column('D:D', 22)
    ws3.set_column('E:E', 12)
    add_excel_header(ws3, styles, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    ws3.merge_range('A8:E8', "  القسم الرابع: تحليل المخاطر ومخاوف الحماية / Section 4: Risk Analysis & Protection Concerns", styles['section'])
    ws3.set_row(7, 24)

    ws3.merge_range('A10:E10', "يرجى تقديم وصف تفصيلي لمخاطر الحماية التي يواجهها المستفيد وأسرته / Detailed protection risks:", styles['label'])
    ws3.merge_range('A11:E15', "\n\n\n\n", styles['value'])
    ws3.set_row(10, 80)

    ws3.merge_range('A17:E17', "مخاوف ومخاطر الحماية المحددة / Specific Protection Risks & Concerns", styles['section'])
    ws3.set_row(16, 24)

    risk_chk_data = [
        ("[  ] تهديدات الأمن والسلامة / Safety & Security Threats", "[  ] الضيق النفسي والاجتماعي / Psychosocial distress"),
        ("[  ] عمالة الأطفال / Child labor", "[  ] نقص الوثائق المدنية / Lack of civil documentation"),
        ("[  ] خطر الإخلاء / Risk of eviction", "[  ] التمييز والاستبعاد / Discrimination & exclusion"),
        ("[  ] العنف القائم على النوع الاجتماعي / Gender-based violence", "[  ] مخاطر أخرى / Other risks: ........................................")
    ]
    r_idx = 18
    for r in risk_chk_data:
        ws3.set_row(r_idx, 20)
        ws3.merge_range(r_idx, 0, r_idx, 1, r[0], styles['value'])
        ws3.merge_range(r_idx, 2, r_idx, 4, r[1], styles['value'])
        r_idx += 1

    # Form 1 page 4
    ws4 = workbook.add_worksheet("Form 1 - Page 4")
    apply_page_setup(ws4)
    ws4.set_column('A:A', 22)
    ws4.set_column('B:B', 22)
    ws4.set_column('C:C', 22)
    ws4.set_column('D:D', 22)
    ws4.set_column('E:E', 12)
    add_excel_header(ws4, styles, "استمارة تقييم حالة فردية", "Individual Assessment Form (PIAF)")

    ws4.merge_range('A8:E8', "  القسم الخامس: خطة العمل والإحالات / Section 5: Action Plan & Referrals", styles['section'])
    ws4.set_row(7, 24)

    ws4.set_row(9, 30)
    ws4.merge_range('A10:B10', "المساعدة الموصى بها / Recommended Assistance", styles['label'])
    ws4.merge_range('C10:E10', "[  ] خدمات الحماية / Protection services   [  ] المساعدة النقدية / Cash support\n[  ] دعم المأوى / Shelter support          [  ] المساعدة القانونية / Legal support", styles['value'])

    ws4.set_row(10, 20)
    ws4.merge_range('A11:B11', "إحالة داخلية / Internal Referral", styles['label'])
    ws4.merge_range('C11:E11', "إلى قسم: ................................................................................", styles['value'])

    ws4.set_row(11, 20)
    ws4.merge_range('A12:B12', "إحالة خارجية / External Referral", styles['label'])
    ws4.merge_range('C12:E12', "إلى منظمة/جهة: ........................................................................", styles['value'])

    ws4.set_row(12, 20)
    ws4.merge_range('A13:B13', "الإجراء المتخذ فورا / Immediate Action Taken", styles['label'])
    ws4.merge_range('C13:E13', "........................................................................................................", styles['value'])

    ws4.merge_range('A15:E15', "التوقيعات والموافقات / Signatures & Approvals", styles['section'])
    ws4.set_row(14, 24)

    ws4.merge_range('A17:B20', "اسم وتوقيع الموظف المقيم / Assessor Signature\n\nالاسم:\nالتوقيع:\nالتاريخ / Date: ______/______/2025", styles['value'])
    ws4.merge_range('C17:E20', "اسم وتوقيع مشرف الحماية / Supervisor Signature\n\nالاسم:\nالتوقيع:\nالتاريخ / Date: ______/______/2025", styles['value'])

    workbook.close()
    output.seek(0)
    return output.read()

def generate_form2_xlsx():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    styles = create_excel_styles(workbook)

    # Form 2 page 1
    ws = workbook.add_worksheet("Form 2 - Page 1")
    apply_page_setup(ws)
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 22)
    ws.set_column('D:D', 22)
    ws.set_column('E:E', 12)

    add_excel_header(ws, styles, "نموذج دراسة الحالة وإدارة الحالة", "Protection Case Study and Case Management Form")

    ws.merge_range('A8:E8', "  القسم الأول: تسجيل الحالة والمعلومات العامة / Section 1: Case Registration", styles['section'])
    ws.set_row(7, 24)

    reg_data = [
        ("كود الحالة / Case Code", "........................................", "تاريخ فتح الحالة / Case Opening Date", "______/______/2025"),
        ("مسؤول الحالة / Case Worker", "........................................", "مصدر الإحالة / Referral Source", "........................................"),
        ("الاسم الكامل للمستفيد / Full Name", "................................................................................", "العمر / Age", ".................. years / سنة"),
        ("الجنس / Gender", "[  ] ذكر / Male      [  ] أنثى / Female", "رقم الهاتف / Phone Number", "........................................"),
        ("العنوان الحالي / Current Address", "................................................................................", "نوع الوثيقة القانونية / Legal Document", "........................................")
    ]

    row_num = 8
    for r in reg_data:
        ws.set_row(row_num, 20)
        ws.write(row_num, 0, r[0], styles['label'])
        ws.write(row_num, 1, r[1], styles['value'])
        ws.write(row_num, 2, r[2], styles['label'])
        ws.merge_range(row_num, 3, row_num, 4, r[3], styles['value'])
        row_num += 1

    row_num += 1
    ws.merge_range(f'A{row_num+1}:E{row_num+1}', "  القسم الثاني: وصف وتقييم الحالة / Section 2: Case Description & Assessment", styles['section'])
    ws.set_row(row_num, 24)
    row_num += 1

    ws.merge_range(f'A{row_num+2}:E{row_num+2}', "الخلفية التفصيلية للحالة ومخاطر الحماية الفورية / Detailed background of the case:", styles['label'])
    ws.merge_range(f'A{row_num+3}:E{row_num+8}', "\n\n\n\n\n\n", styles['value'])
    ws.set_row(row_num+1, 100)

    # Form 2 page 2
    ws2 = workbook.add_worksheet("Form 2 - Page 2")
    apply_page_setup(ws2)
    ws2.set_column('A:A', 22)
    ws2.set_column('B:B', 22)
    ws2.set_column('C:C', 22)
    ws2.set_column('D:D', 22)
    ws2.set_column('E:E', 12)
    add_excel_header(ws2, styles, "نموذج دراسة الحالة وإدارة الحالة", "Protection Case Study and Case Management Form")

    ws2.merge_range('A8:E8', "  القسم الثالث: خطة عمل الحالة / Section 3: Case Action Plan", styles['section'])
    ws2.set_row(7, 24)

    # Table headers
    ws2.set_row(9, 20)
    ws2.write(9, 0, "الهدف / Objective", styles['table_header'])
    ws2.write(9, 1, "الإجراء المطلوب / Action", styles['table_header'])
    ws2.write(9, 2, "الشخص المسؤول / Resp.", styles['table_header'])
    ws2.write(9, 3, "الإطار الزمني / Timeframe", styles['table_header'])
    ws2.write(9, 4, "الحالة / Status", styles['table_header'])

    for r_idx in range(10, 13):
        ws2.set_row(r_idx, 25)
        for c_idx in range(5):
            ws2.write(r_idx, c_idx, "", styles['value'])

    ws2.merge_range('A14:E14', "  القسم الرابع: المتابعة وإغلاق الحالة / Section 4: Follow-up & Case Closure", styles['section'])
    ws2.set_row(13, 24)

    ws2.set_row(15, 20)
    ws2.merge_range('A16:B16', "ملخص زيارات المتابعة / Follow-up summary", styles['label'])
    ws2.merge_range('C16:E16', "................................................................................", styles['value'])

    ws2.set_row(16, 20)
    ws2.merge_range('A17:B17', "حالة الحالة / Case Status", styles['label'])
    ws2.merge_range('C17:E17', "[  ] نشطة / Active         [  ] مغلقة / Closed         [  ] معلقة / Suspended", styles['value'])

    ws2.set_row(17, 20)
    ws2.merge_range('A18:B18', "أسباب الإغلاق / Closure reasons", styles['label'])
    ws2.merge_range('C18:E18', "[  ] تم تحقيق الأهداف / Goals met    [  ] اكتمال الإحالة / Referral complete", styles['value'])

    ws2.merge_range('A20:E20', "التوقيعات والموافقات / Signatures & Approvals", styles['section'])
    ws2.set_row(19, 24)

    ws2.merge_range('A22:B25', "اسم وتوقيع مسؤول الحالة / Case Worker\n\nالاسم:\nالتوقيع:\nالتاريخ / Date: ______/______/2025", styles['value'])
    ws2.merge_range('C22:E25', "اسم وتوقيع مشرف إدارة الحالة / Supervisor\n\nالاسم:\nالتوقيع:\nالتاريخ / Date: ______/______/2025", styles['value'])

    workbook.close()
    output.seek(0)
    return output.read()

def generate_form3_xlsx():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    styles = create_excel_styles(workbook)

    # Form 3 page 1
    ws = workbook.add_worksheet("Form 3 - Page 1")
    apply_page_setup(ws)
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 22)
    ws.set_column('D:D', 22)
    ws.set_column('E:E', 12)

    add_excel_header(ws, styles, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    ws.merge_range('A8:E8', "  المعلومات الأساسية / Basic Information", styles['section'])
    ws.set_row(7, 24)

    info_data = [
        ("معرّف الحالة / Case ID", "................................................................................"),
        ("تاريخ التقييم / Evaluation Date", "2025 / ______ / ______"),
        ("تاريخ قرار الموافقة / Approval Decision Date", "2025 / ______ / ______"),
        ("نوع الخطر الحمائي / Protection Risk Type", "................................................................................")
    ]
    row_num = 8
    for r in info_data:
        ws.set_row(row_num, 20)
        ws.merge_range(row_num, 0, row_num, 1, r[0], styles['label'])
        ws.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    row_num += 1
    ws.merge_range(f'A{row_num+1}:E{row_num+1}', "  القسم الأول: ملخص الحالة / Section 1: Case Summary", styles['section'])
    ws.set_row(row_num, 24)
    row_num += 1

    ws.merge_range(f'A{row_num+1}:E{row_num+1}', "1.1 البيانات الأساسية للمستفيد / Beneficiary Basic Data", styles['label'])
    row_num += 1

    bdata_rows = [
        ("الاسم أو معرف / Name or ID", "................................................................................"),
        ("العمر / Age", ".................. سنة / years"),
        ("الجنس / Gender", "[  ] ذكر / Male         [  ] أنثى / Female"),
        ("عدد أفراد الأسرة / Family Members", ".................. أفراد / members")
    ]
    for r in bdata_rows:
        ws.set_row(row_num, 20)
        ws.merge_range(row_num, 0, row_num, 1, r[0], styles['label'])
        ws.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    row_num += 1
    ws.merge_range(f'A{row_num+1}:E{row_num+1}', "1.2 ملخص الخطر / Risk Summary", styles['label'])
    row_num += 1

    risk_rows = [
        ("الخطر الأساسي / Primary Risk", "\n\n\n"),
        ("الشدة / Severity", "[  ] حاد وفوري جداً (24-48 ساعة) / Very acute and immediate (24-48 hours)\n[  ] حاد (3-5 أيام) / Acute (3-5 days)\n[  ] متوسط (أسبوع - أسبوعين) / Moderate (1-2 weeks)")
    ]
    for r in risk_rows:
        ws.set_row(row_num, 20)
        ws.merge_range(row_num, 0, row_num, 1, r[0], styles['label'])
        ws.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    # Form 3 page 2
    ws2 = workbook.add_worksheet("Form 3 - Page 2")
    apply_page_setup(ws2)
    ws2.set_column('A:A', 22)
    ws2.set_column('B:B', 22)
    ws2.set_column('C:C', 22)
    ws2.set_column('D:D', 22)
    ws2.set_column('E:E', 12)
    add_excel_header(ws2, styles, "نموذج قرار النقد مقابل الحماية", "Cash for Protection Approval Form")

    ws2.merge_range('A8:E8', "  1.3 الحل المقترح والمبلغ / Proposed Solution and Amount", styles['section'])
    ws2.set_row(7, 24)

    sol_rows = [
        ("كيف سيساعد النقد؟ / How will cash help?", "\n\n"),
        ("المبلغ المقترح / Proposed Amount", "........................................ دولار / dollars"),
        ("السبب / Reason", "\n\n")
    ]
    row_num = 8
    for r in sol_rows:
        ws2.set_row(row_num, 20)
        ws2.merge_range(row_num, 0, row_num, 1, r[0], styles['label'])
        ws2.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    row_num += 1
    ws2.merge_range(f'A{row_num+1}:E{row_num+1}', "  القسم الثاني: موافقة قسم الحماية / Section 2: Protection Department Approval", styles['section'])
    ws2.set_row(row_num, 24)
    row_num += 1

    ws2.merge_range(f'A{row_num+1}:E{row_num+1}', "2.1 مراجعة التقييم / Evaluation Review", styles['label'])
    row_num += 1

    review_rows = [
        ("وضوح الخطر / Risk Clarity", "[  ] واضح جداً / Very Clear      [  ] واضح / Clear      [  ] غامض / Vague"),
        ("الربط السببي / Causal Link", "[  ] منطقي جداً / Very Logical   [  ] منطقي / Logical   [  ] ضعيف / Weak"),
        ("كفاية التقييم / Evaluation Adequacy", "[  ] شامل / Comprehensive         [  ] متوسط / Moderate     [  ] ناقص / Incomplete"),
        ("المخاطر الثانوية / Secondary Risks", "[  ] منخفضة / Low               [  ] متوسطة / Medium     [  ] عالية / High")
    ]
    for r in review_rows:
        ws2.set_row(row_num, 20)
        ws2.merge_range(row_num, 0, row_num, 1, r[0], styles['label'])
        ws2.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    row_num += 1
    ws2.merge_range(f'A{row_num+1}:E{row_num+1}', "  القسم الثالث: القرار النهائي من مدير الحماية / Section 3: Final Decision", styles['section'])
    ws2.set_row(row_num, 24)
    row_num += 1

    mgr_rows = [
        ("القرار النهائي للمدير / Final Decision", "[  ] موافق على الصرف / Approved       [  ] غير موافق / Not Approved"),
        ("المبلغ وآلية الصرف / Amount & Mechanism", "المبلغ: ........................ دولار    الآلية: [  ] نقد / Cash   [  ] حوالة / Transfer"),
        ("التاريخ والتوقيع / Date & Signature", "الاسم: ........................................................\nالتوقيع: ....................................................\nالتاريخ / Date: ______/______/2025")
    ]
    for r in mgr_rows:
        ws2.set_row(row_num, 20)
        ws2.merge_range(row_num, 0, row_num, 1, r[0], styles['label'])
        ws2.merge_range(row_num, 2, row_num, 4, r[1], styles['value'])
        row_num += 1

    workbook.close()
    output.seek(0)
    return output.read()
