from flask import Flask, request, send_file, render_template, redirect, url_for, flash
import os
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp/'
app.secret_key = 'your_secret_key'  # Required for flash messages


def allowed_file(filename):
    """Check if the uploaded file is an allowed type."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'csv'}


@app.route('/', methods=['GET'])
def index():
    """Render the HTML page for file upload."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing."""
    try:
        if 'file' not in request.files:
            flash('No file part in the request.')
            return redirect(request.url)

        file = request.files['file']
        group_name = request.form.get('group_name', '').strip()
        column_name = request.form.get('column_name', '').strip()
        match_value = request.form.get('match_value', '').strip()

        if file.filename == '' or not allowed_file(file.filename):
            flash('Invalid file type or no file selected. Please upload an Excel or CSV file.')
            return redirect(request.url)

        if not group_name:
            flash('Group name is required.')
            return redirect(request.url)

        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        file_ext = file.filename.rsplit('.', 1)[1].lower()

        if file_ext == 'xlsx':
            df = load_excel(filename)
        elif file_ext == 'csv':
            df = pd.read_csv(filename)
        else:
            flash('Unsupported file format.')
            return redirect(request.url)

        if column_name not in df.columns:
            flash(f"The column '{column_name}' does not exist in the file.")
            return redirect(request.url)

        matched_df = df[df[column_name].astype(str).str.strip() == match_value.strip()]
        if matched_df.empty:
            flash('No matching entries found.')
            return redirect(request.url)

        putty_xml_content = generate_putty_sessions_xml(matched_df, group_name, match_value)

        processed_filename = f"processed_{file.filename.rsplit('.', 1)[0]}.xml"
        processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)

        with open(processed_filepath, 'w', encoding='utf-8') as putty_file:
            putty_file.write(putty_xml_content)

        return redirect(url_for('download_file', filename=processed_filename))

    except Exception as e:
        app.logger.error('Error processing file: %s', e)
        flash('An error occurred while processing the file. Please try again.')
        return redirect(request.url)


def load_excel(filename):
    """Load data from an Excel file starting from sheet 2, row 7."""
    try:
        wb = pd.ExcelFile(filename)
        if len(wb.sheet_names) < 2:
            raise ValueError("The Excel file does not contain a second sheet.")
        df = pd.read_excel(wb, sheet_name=wb.sheet_names[1], header=6)
        return df
    except Exception as e:
        app.logger.error('Error loading Excel file: %s', e)
        raise


def prettify_xml(element):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(element, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


def generate_putty_sessions_xml(df, group_name, match_value):
    """Generate the XML content for Putty sessions."""
    root = ET.Element('ArrayOfSessionData')
    root.set('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema')
    root.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')

    folder_mapping = {
        'exporter_linux': 'Linux Server',
        'exporter_gateway': 'Media Gateway',
        'exporter_windows': 'Windows Server',
        'exporter_verint': 'Verint Server',
    }

    subfolder = folder_mapping.get(match_value, 'Other')

    for _, row in df.iterrows():
        session_data = ET.SubElement(root, 'SessionData')
        session_id = (
            f"{group_name}/{subfolder}/"
            f"{row['Country']}/{row['Location']}/{row['Hostnames']}"
        )
        session_data.set('SessionId', session_id)
        session_data.set('SessionName', row['Hostnames'])
        session_data.set('Host', row['IP Address'])

        if match_value in ['exporter_windows', 'exporter_verint']:
            session_data.set('ImageKey', match_value.split('_')[1])
            session_data.set('Port', '3389')
            session_data.set('Proto', 'RDP')
        else:
            session_data.set('ImageKey', 'tux')
            session_data.set('Port', '22')
            session_data.set('Proto', 'SSH')
            session_data.set('PuttySession', 'Default Settings')
            if pd.notna(row.get('ssh_username', None)) and str(
                row['ssh_username']
            ).strip():
                session_data.set('Username', str(row['ssh_username']))

        secret_server_url = row.get('Secret Server', None)
        if secret_server_url:
            ET.SubElement(session_data, 'SPSLFileName').text = secret_server_url

    return prettify_xml(root)


@app.route('/downloads/<filename>', methods=['GET'])
def download_file(filename):
    """Handle file download after processing."""
    try:
        download_folder = app.config['UPLOAD_FOLDER']
        file_path = os.path.join(download_folder, filename)

        if not os.path.isfile(file_path):
            flash('File not found.')
            return redirect(url_for('index'))

        response = send_file(file_path, as_attachment=True, download_name=filename)
        os.remove(file_path)

        return response
    except Exception as e:
        app.logger.error('Error downloading file: %s', e)
        flash('An error occurred while downloading the file.')
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
