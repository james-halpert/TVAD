![image](https://github.com/user-attachments/assets/65e61735-52d1-43b5-90ae-0b60fc2fdfc6)

# TeamViewer AD Check

## Overview
This project is a **Flask-based web application** that allows users to upload a list of email addresses and check their existence in Active Directory (AD). If an email exists, it retrieves the user's **name, office, department, and title** from AD and exports the results into a downloadable Excel file.

## Features
- **Auto-detects the LDAP server** and pre-fills the form.
- **Auto-populates the LDAP search base** based on the detected domain.
- **Auto-fills the username in DOMAIN\username format**.
- **Real-time progress updates** while processing AD queries.
- **Audio notification** when processing starts.
- **Downloadable Excel file** containing AD lookup results.
- **Bootstrap-based UI** for a clean and responsive design.

## Requirements
- Python 3.x
- Flask
- ldap3
- pandas
- openpyxl
- Bootstrap (included via CDN)
- pycryptodome

## Installation
1. Clone this repository:
   ```sh
   git clone https://github.com/james-halpert/TVAD.git
   cd teamviewer-ad-check
   ```
2. Install dependencies:
   ```sh
   pip install flask pandas ldap3 pycryptodome
   ```
3. Run the application:
   ```sh
   python app.py
   ```
4. Open your browser and go to:
   ```
   http://127.0.0.1:5500
   ```

## Usage
1. **Enter LDAP Credentials:**
   - The **LDAP Server** auto-populates based on `nslookup`.
   - The **LDAP User** auto-populates with `DOMAIN\username`.
   - The **LDAP Search Base** auto-fills based on the detected domain.
2. **Upload a TXT file** containing email addresses (one per line).
3. Click **Upload & Check** to start the AD lookup process.
4. **Progress updates** dynamically display the number of records being processed.
5. Once complete, a **download button** appears to get the results in Excel format.

## File Structure
```
teamviewer-ad-check/
│-- static/
│   ├── placeholder-logo.png  # Placeholder for company logo
│-- templates/
│   ├── index.html             # Main UI template
│   ├── progress.html          # Progress page for real-time updates
│-- app.py                      # Flask backend
│-- README.md                   # Project documentation
```

## Notes
- Ensure the **LDAP user has read access** to Active Directory.
- If the **LDAP search base does not auto-populate**, verify that `nslookup -type=SRV _ldap._tcp` correctly returns the domain information.
- If you encounter **MD4 hash errors**, your system might require NTLMv2 enforcement.
