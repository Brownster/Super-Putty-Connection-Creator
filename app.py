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
        app.logger.debug('Received request to upload file.')

        if 'file' not in request.files:
            app.logger.error('No file part in the request.')
            flash('No file part in the request.')
            return redirect(url_for('index'))

        file = request.files['file']
        group_name = request.form.get('group_name', '').strip()

        app.logger.debug(f'File received: {file.filename}')
        app.logger.debug(f'Group name received: {group_name}')

        if file.filename == '' or not allowed_file(file.filename):
            app.logger.error('Invalid file type or no file selected.')
            flash('Invalid file type or no file selected. Please upload an Excel or CSV file.')
            return redirect(url_for('index'))

        if not group_name:
            app.logger.error('Group name is required.')
            flash('Group name is required.')
            return redirect(url_for('index'))

        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        file_ext = file.filename.rsplit('.', 1)[1].lower()

        if file_ext == 'xlsx':
            df = load_excel(filename)
        elif file_ext == 'csv':
            df = load_csv(filename)
        else:
            app.logger.error('Unsupported file format.')
            flash('Unsupported file format.')
            return redirect(url_for('index'))

        putty_xml_content = generate_putty_sessions_xml(df, group_name)

        processed_filename = f"processed_{file.filename.rsplit('.', 1)[0]}.xml"
        processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)

        with open(processed_filepath, 'w', encoding='utf-8') as putty_file:
            putty_file.write(putty_xml_content)

        app.logger.debug(f'File processed and saved as {processed_filename}')

        return redirect(url_for('download_file', filename=processed_filename))

    except Exception as e:
        app.logger.error('Error processing file: %s', e)
        flash('An error occurred while processing the file. Please try again.')
        return redirect(url_for('index'))

def load_excel(filename):
    """Load data from an Excel file starting from sheet 2, row 7."""
    try:
        wb = pd.ExcelFile(filename)
        if len(wb.sheet_names) < 2:
            raise ValueError("The Excel file does not contain a second sheet.")
        df = pd.read_excel(wb, sheet_name=wb.sheet_names[1], header=6)
        
        # Define only the columns you actually need for processing
        required_columns = [
            'Country', 'Location', 'Hostnames', 'IP Address',
            'Exporter_name_os', 'Exporter_name_app', 'ssh_username'
        ]

        # Check for missing columns
        missing_cols = set(required_columns) - set(df.columns)
        if missing_cols:
            raise ValueError(f"Missing columns in Excel: {missing_cols}")

        return df
    except Exception as e:
        app.logger.error('Error loading Excel file: %s', e)
        raise

def load_csv(filename):
    """Load data from a CSV file starting from the header row."""
    try:
        # Read the CSV, starting from the actual header row, usually row 7 (zero-indexed in Pandas)
        df = pd.read_csv(filename, delimiter=',', skiprows=6, skipinitialspace=True)

        # Define only the columns you actually need for processing
        required_columns = [
            'Country', 'Location', 'Hostnames', 'IP Address',
            'Exporter_name_os', 'Exporter_name_app', 'ssh_username'
        ]

        # Check for missing columns
        missing_cols = set(required_columns) - set(df.columns)
        if missing_cols:
            raise ValueError(f"Missing columns in CSV: {missing_cols}")

        return df

    except pd.errors.ParserError as e:
        app.logger.error('CSV parsing error: %s', e)
        raise
    except Exception as e:
        app.logger.error('Error loading CSV file: %s', e)
        raise

def prettify_xml(element):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(element, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def generate_putty_sessions_xml(df, group_name):
    """Generate the XML content for Putty sessions."""
    root = ET.Element('ArrayOfSessionData')
    root.set('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema')
    root.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')

    folder_mapping = {
        'exporter_linux': 'Linux Server',
        'exporter_gateway': 'Media Gateway',
        'exporter_windows': 'Windows Server',
        'exporter_verint': 'Verint Server',
        'exporter_vmware': 'VMware Server',  # Separate folder for VMware
        'exporter_aacc': 'AACC Server',  # Separate folder for AACC
    }

    processed_sessions = set()  # Set to track processed sessions

    for _, row in df.iterrows():
        # Determine the exporter type, preferring OS over application to avoid duplicates
        exporter_type_os = str(row.get('Exporter_name_os', 'other')).lower()
        exporter_type_app = str(row.get('Exporter_name_app', 'other')).lower()
        
        # Choose the correct exporter type based on precedence and avoid duplicates
        exporter_type = exporter_type_os if exporter_type_os != 'other' else exporter_type_app

        # Skip if session already processed
        session_key = (row['Country'], row['Location'], row['Hostnames'], exporter_type)
        if session_key in processed_sessions:
            continue
        
        # Mark this session as processed
        processed_sessions.add(session_key)

        subfolder = folder_mapping.get(exporter_type, 'Other')

        session_data = ET.SubElement(root, 'SessionData')
        session_id = (
            f"{group_name}/{subfolder}/"
            f"{row['Country']}/{row['Location']}/{row['Hostnames']}"
        )
        session_data.set('SessionId', session_id)
        session_data.set('SessionName', str(row['Hostnames']))  # Ensure string conversion
        session_data.set('Host', str(row['IP Address']))  # Ensure string conversion

        # Handle different types of exporters for connection types
        if exporter_type in ['exporter_windows', 'exporter_verint', 'exporter_aacc']:
            session_data.set('ImageKey', exporter_type.split('_')[1] if '_' in exporter_type else exporter_type)
            session_data.set('Port', '3389')
            session_data.set('Proto', 'RDP')
        else:
            session_data.set('ImageKey', 'tux')
            session_data.set('Port', '22')
            session_data.set('Proto', 'SSH')
            session_data.set('PuttySession', 'Default Settings')
            if pd.notna(row.get('ssh_username')) and str(row['ssh_username']).strip():
                session_data.set('Username', str(row['ssh_username']))

        secret_server_url = row.get('SS URL', None)
        if secret_server_url:
            ET.SubElement(session_data, 'SPSLFileName').text = str(secret_server_url)  # Ensure string conversion

    return prettify_xml(root)

@app.route('/downloads/<filename>', methods=['GET'])
def download_file(filename):
    """Handle file download after processing."""
    try:
        download_folder = app.config['UPLOAD_FOLDER']
        file_path = os.path.join(download_folder, filename)

        if not os.path.isfile(file_path):
            app.logger.error('File not found for download: %s', file_path)
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
