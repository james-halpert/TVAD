from Cryptodome.Hash import MD4
import sys
import os
import tempfile
from flask import Flask, render_template, request, send_file, Response, stream_with_context
import pandas as pd
import ldap3
import getpass
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor
import subprocess


# Check and install missing dependencies
required = ["flask", "pandas", "ldap3", "pycryptodome"]
for package in required:
    try:
        __import__(package)
    except ImportError:
        print(f"Installing missing package: {package}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])


# Determine base directory for PyInstaller bundling
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"), static_folder=os.path.join(BASE_DIR, "static"))

LDAP_ATTRIBUTES = ['displayName', 'physicalDeliveryOfficeName', 'department', 'title']
TEMP_DIR = tempfile.gettempdir()  # Directory for temp files
EMAILS_FILE = os.path.join(TEMP_DIR, "emails_to_process.txt")  # Temp file path
EXCEL_FILE = os.path.join(TEMP_DIR, "ldap_results.xlsx")  # Temp Excel file path

def get_ldap_details():
    """Auto-detects LDAP server and determines the search base."""
    try:
        result = os.popen("nslookup -type=SRV _ldap._tcp").read()
        ldap_server = ""
        search_base = ""
        for line in result.split("\n"):
            if "svr hostname" in line.lower():
                ldap_server = line.split()[-1].strip()
                break
        if ldap_server and "." in ldap_server:
            domain_parts = ldap_server.split(".")[1:]
            search_base = ",".join([f"DC={part}" for part in domain_parts])
        return ldap_server, search_base
    except Exception:
        return "", ""

def get_current_user():
    """Gets the current domain and username."""
    domain = os.environ.get('USERDOMAIN', '')
    username = getpass.getuser()
    return f"{domain}\\{username}" if domain else username

def search_ad(email, ldap_server, ldap_user, ldap_password, ldap_search_base):
    """Performs an LDAP search for a given email, checking both primary email and aliases."""
    try:
        server = ldap3.Server(ldap_server)
        conn = ldap3.Connection(
            server,
            user=ldap_user,
            password=ldap_password,
            authentication=ldap3.NTLM,
            auto_bind=True
        )

        search_filter = f"(|(mail={email})(proxyAddresses=smtp:{email}))"  # Check primary mail and aliases
        conn.search(ldap_search_base, search_filter, attributes=LDAP_ATTRIBUTES + ['proxyAddresses'])

        if conn.entries:
            entry = conn.entries[0]
            return {
                'Email': email,
                'Name': entry.displayName.value if entry.displayName else 'Not Found',
                'Office': entry.physicalDeliveryOfficeName.value if entry.physicalDeliveryOfficeName else '',
                'Department': entry.department.value if entry.department else '',
                'Title': entry.title.value if entry.title else '',
                'Aliases': ", ".join([str(a) for a in entry.proxyAddresses]) if entry.proxyAddresses else ''
            }

        return {'Email': email, 'Name': 'Not Found', 'Office': '', 'Department': '', 'Title': '', 'Aliases': ''}

    except Exception as e:
        return {'Email': email, 'Name': f'Error: {str(e)}', 'Office': '', 'Department': '', 'Title': '', 'Aliases': ''}

def generate_progress(ldap_server, ldap_user, ldap_password, ldap_search_base):
    """Reads emails from file, processes them with LDAP lookups, and updates the UI in real-time."""
    if not os.path.exists(EMAILS_FILE):
        yield "data: ERROR - No email file found\n\n"
        return

    with open(EMAILS_FILE, "r") as f:
        emails = [line.strip() for line in f.readlines() if line.strip()]

    results = []
    total = len(emails)

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(search_ad, email, ldap_server, ldap_user, ldap_password, ldap_search_base): email for email in emails}

        for index, future in enumerate(futures, start=1):
            results.append(future.result())
            yield f"data: Processing {index}/{total}\n\n"

    df = pd.DataFrame(results)
    df.to_excel(EXCEL_FILE, index=False)

    yield "data: COMPLETE\n\n"

@app.route('/progress')
def progress():
    """Handles the real-time streaming of progress updates."""
    return Response(stream_with_context(generate_progress(ldap_server, ldap_user, ldap_password, ldap_search_base)), mimetype='text/event-stream')

@app.route('/download')
def download():
    """Allows users to download the generated Excel file."""
    if not os.path.exists(EXCEL_FILE):
        return "No file available for download", 400
    return send_file(EXCEL_FILE, as_attachment=True, download_name='TeamViewer_AD_Check.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/', methods=['GET', 'POST'])
def index():
    """Handles the file upload and form submission."""
    global ldap_server, ldap_user, ldap_password, ldap_search_base
    detected_ldap_server, detected_search_base = get_ldap_details()
    current_user = get_current_user()

    if request.method == 'POST':
        ldap_server = request.form['ldap_server']
        ldap_user = request.form['ldap_user']
        ldap_password = request.form['ldap_password']
        ldap_search_base = request.form['ldap_search_base']

        file = request.files['file']
        if file:
            with open(EMAILS_FILE, "w") as f:
                for line in file.readlines():
                    f.write(line.decode('utf-8').strip() + "\n")  # Save each email to the file

            return render_template('progress.html')

    return render_template('index.html', detected_ldap_server=detected_ldap_server, detected_search_base=detected_search_base, current_user=current_user)

if __name__ == '__main__':
    app.run(debug=True, port=5500)
