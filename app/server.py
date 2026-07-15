import os
from flask import Flask, render_template, send_file, Response
from docx_generator import generate_form1_docx, generate_form2_docx, generate_form3_docx
from excel_generator import generate_form1_xlsx, generate_form2_xlsx, generate_form3_xlsx

app = Flask(__name__, static_folder='static', template_folder='templates')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/export/word/<int:form_id>')
def export_word(form_id):
    if form_id == 1:
        data = generate_form1_docx()
        filename = "Individual_Assessment_Form_PIAF.docx"
    elif form_id == 2:
        data = generate_form2_docx()
        filename = "Protection_Case_Study_and_Case_Management_Form.docx"
    elif form_id == 3:
        data = generate_form3_docx()
        filename = "Cash_for_Protection_Approval_Form.docx"
    else:
        return "Invalid Form ID", 400

    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-disposition": f"attachment; filename={filename}"}
    )

@app.route('/export/excel/<int:form_id>')
def export_excel(form_id):
    if form_id == 1:
        data = generate_form1_xlsx()
        filename = "Individual_Assessment_Form_PIAF.xlsx"
    elif form_id == 2:
        data = generate_form2_xlsx()
        filename = "Protection_Case_Study_and_Case_Management_Form.xlsx"
    elif form_id == 3:
        data = generate_form3_xlsx()
        filename = "Cash_for_Protection_Approval_Form.xlsx"
    else:
        return "Invalid Form ID", 400

    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-disposition": f"attachment; filename={filename}"}
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
