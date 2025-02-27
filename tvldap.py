from flask import Flask, render_template, request, send_file, Response, stream_with_context
import pandas as pd
import ldap3
import os
import socket
import getpass
from io import BytesIO
import time

app = Flask(__name__)

# LDAP Attributes to Retrieve
LDAP_ATTRIBUTES = ['displayName', 'physicalDeliveryOfficeName', 'department', 'title']

# Function to auto-detect LDAP server and determine search base
def get_ldap_details():
    try:
        result = os.popen("nslookup -type=SRV _ldap._tcp").read()
        lines = result.split('\n')
        for line in lines:
            if 'svr hostname' in line.lower():
                ldap_server = line.split()[-1].strip()
                domain_parts = [part for part in ldap_server.split('.') if "dc" not in part.lower()]
                if len(domain_parts) >= 2:  # Ensure we have at least a domain name structure
                    search_base = ','.join([f"DC={part}" for part in domain_parts[-2:]])
                    return ldap_server, search_base
    except Exception as e:
        print(f"Error detecting LDAP server: {e}")
    return "", ""


# Function to search AD
def search_ad(email, ldap_server, ldap_user, ldap_password, ldap_search_base):
    server = ldap3.Server(ldap_server)
    conn = ldap3.Connection(
        server,
        user=ldap_user,
        password=ldap_password,
        authentication=ldap3.NTLM,
        auto_bind=True
    )

    search_filter = f'(mail={email})'
    conn.search(ldap_search_base, search_filter, attributes=LDAP_ATTRIBUTES)
    
    if conn.entries:
        entry = conn.entries[0]
        return {
            'Name': entry.displayName.value if entry.displayName else '',
            'Office': entry.physicalDeliveryOfficeName.value if entry.physicalDeliveryOfficeName else '',
            'Department': entry.department.value if entry.department else '',
            'Title': entry.title.value if entry.title else ''
        }
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    detected_ldap_server, detected_search_base = get_ldap_details()
    current_user = getpass.getuser()
    
    if request.method == 'POST':
        ldap_server = request.form['ldap_server']
        ldap_user = request.form['ldap_user']
        ldap_password = request.form['ldap_password']
        ldap_search_base = request.form['ldap_search_base']
        
        file = request.files['file']
        if file:
            emails = [line.strip().decode('utf-8') for line in file.readlines() if line.strip()]
            results = []
            for email in emails:
                user_info = search_ad(email, ldap_server, ldap_user, ldap_password, ldap_search_base)
                results.append({
                    'Email': email,
                    'Name': user_info.get('Name', 'Not Found') if user_info else 'Not Found',
                    'Office': user_info.get('Office', '') if user_info else '',
                    'Department': user_info.get('Department', '') if user_info else '',
                    'Title': user_info.get('Title', '') if user_info else ''
                })
            df = pd.DataFrame(results)
            excel_output = BytesIO()
            df.to_excel(excel_output, index=False)
            excel_output.seek(0)
            return send_file(excel_output, as_attachment=True, download_name='TeamViewer_AD_Check.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    return render_template('index.html', detected_ldap_server=detected_ldap_server, detected_search_base=detected_search_base, current_user=current_user)

if __name__ == '__main__':
    app.run(debug=True, port=5500)
