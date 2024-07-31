from flask import Flask, request, render_template, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    sheet_name = request.form['sheet_name']
    user_data = request.form['user_data']

    # Read the Excel file
    xls = pd.ExcelFile(file)
    if sheet_name not in xls.sheet_names:
        return f"Sheet {sheet_name} not found in the uploaded file."

    df = pd.read_excel(file, sheet_name=sheet_name)

    # Assuming user_data is a comma-separated string of values to be appended as a new row
    user_data_list = user_data.split(',')
    df.loc[len(df)] = user_data_list

    # Save to a new Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)

    return send_file(output, attachment_filename='updated_file.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)


