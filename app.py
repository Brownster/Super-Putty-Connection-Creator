from flask import Flask, request, send_file, render_template, redirect, url_for
import os
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp/'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx'}

@app.route('/', methods=['GET'])
def index():
    # Render the HTML page for file upload
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']
        group_name = request.form['group_name']
        column_name = request.form['column_name']
        match_value = request.form['match_value']

        if file.filename == '' or not allowed_file(file.filename):
            return redirect(request.url)

        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        df = pd.read_excel(filename, engine='openpyxl')

        if column_name not in df.columns:
            return f"The column '{column_name}' does not exist in the Excel file.", 400

        matched_df = df[df[column_name].astype(str).str.strip() == match_value.strip()]
        putty_xml_content = generate_putty_sessions_xml(matched_df, group_name)

        processed_filename = f"processed_{file.filename.rsplit('.', 1)[0]}.xml"
        processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)

        with open(processed_filepath, 'w') as putty_file:
            putty_file.write(putty_xml_content)

        return redirect(url_for('download_file', filename=processed_filename))

    return 'File upload error'

def prettify_xml(element):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(element, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def generate_putty_sessions_xml(df, group_name):
    root = ET.Element('ArrayOfSessionData')
    root.set('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema')
    root.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')

    for index, row in df.iterrows():
        try:
            session_id = f"{group_name}/{str(row['Country'])}/{str(row['Location'])}/{str(row['Hostnames'])}"
            host = f"{str(row['ssh_username'])}@{str(row['IP Address'])}"

            session_data = ET.SubElement(root, 'SessionData')
            session_data.set('SessionId', session_id)
            session_data.set('SessionName', str(row['Hostnames']))
            session_data.set('ImageKey', 'computer')
            session_data.set('Host', host)
            session_data.set('Port', str(row.get('OS-Listen-Port', '22')))
            session_data.set('Proto', 'SSH')
            session_data.set('PuttySession', 'Default Settings')
            session_data.set('Username', str(row['ssh_username']))
            # Add any other necessary attributes or handle them being optional
        except KeyError as e:
            print(f"Missing key in DataFrame row: {e}")
        except Exception as e:
            print(f"An error occurred: {e}")

    return prettify_xml(root)

@app.route('/downloads/<filename>', methods=['GET'])
def download_file(filename):
    download_folder = app.config['UPLOAD_FOLDER']
    file_path = os.path.join(download_folder, filename)
    
    if not os.path.isfile(file_path):
        return "File not found.", 404

    response = send_file(file_path, as_attachment=True, download_name=filename)
    os.remove(file_path)
    
    return response

if __name__ == '__main__':
    app.run(debug=True)

