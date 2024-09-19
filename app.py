from flask import Flask, request, jsonify, send_file
import os
import pdfkit
import pandas as pd

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# 1
def get_excel_info(file_path):
    excel_data = pd.ExcelFile(file_path)
    num_sheets = len(excel_data.sheet_names)
    return {'file_path': file_path, 'num_sheets': num_sheets}



def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
@app.route('/upload-excel', methods=['POST'])
def upload_excel():
    data = request.get_json()

    file_path = data.get('file_path')
    # sheets_data = data.get('sheets')
    if not file_path:
        return jsonify({'error': 'No file part'})

    file = data.get('file')
    if file.filename == '':
        return jsonify({'error': 'No selected file'})

    if file and not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file. Only Excel files allowed'})

    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        excel_info = get_excel_info(file_path)
        return jsonify(excel_info)

# 2
def process_excel_data(file_path, sheets_data):
    report = {}

    for sheet_info in sheets_data:
        sheet_name = sheet_info.get('sheet_name')
        operation = sheet_info.get('operation')
        columns = sheet_info.get('columns')

        # Read the selected sheet from the Excel file
        sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)

        if operation == 'sum':
            result = sheet_data[columns].sum().to_dict()
        elif operation == 'average':
            result = sheet_data[columns].mean().to_dict()
        else:
            result = {'error': 'Unsupported operation'}

        report[sheet_name] = result

    return report


@app.route('/process-excel', methods=['POST'])
def process_excel():
    data = request.get_json()
    file_path = data.get('file_path')
    sheets_data = data.get('sheets')

    report = process_excel_data(file_path, sheets_data)

    return jsonify(report)


# 3

@app.route('/download-pdf', methods=['POST'])
def download_pdf():
    report_data = request.get_json()

    html_content = "<html><head><title>Report</title></head><body>"

    for sheet_name, sheet_result in report_data.items():
        html_content += f"<h2>{sheet_name}</h2>"
        html_content += "<ul>"
        for column, value in sheet_result.items():
            html_content += f"<li>{column}: {value}</li>"
        html_content += "</ul>"

    html_content += "</body></html>"

    pdfkit.from_string(html_content, 'report.pdf')

    return send_file('report.pdf', as_attachment=True)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'

    file = request.files['file']
    if file.filename == '':
        return 'No selected file'

    file.save('uploaded_file.pdf')  # Save the uploaded file

    return 'File uploaded successfully'


if __name__ == '__main__':
    app.run(debug=True)