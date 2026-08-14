# GDCK Admission

Python desktop application for the admission process of:

- BDS
- MDS
- DORA
- Dental Mechanic

## Run

```powershell
python -m pip install -r requirements.txt
python main.py
```

## Install

From the project folder:

```powershell
python -m pip install .
gdck-admission
```

On Windows, you can also run the installer script. It checks for Python, installs
Python with winget when missing, then installs this application and its
dependencies:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\install_windows.ps1
```

For development, install in editable mode:

```powershell
python -m pip install -e .
python -m gdck_admission
```

The packaged app includes the Windows OCR helper and default settings. Live settings,
Google Drive login tokens, and generated receipts are stored in the user's application
data folder instead of inside the installed package.

The application opens a main panel with two buttons:

- Academic Section
- Account Section

Clicking either button opens its own panel.

The Account Section also has:

- Course dropdown: choose BDS, MDS, DORA, or DCDM.
- Category dropdown: choose General, SC, ST, or OEC.
- Settings: manually set fee values for each course and connect a Google Sheet/CSV source for student names.
- Load Names: reads student names from the configured Google Sheet source.
- Generate Receipt: creates a PDF fee acknowledgement with student name, course, category, and fee details.
- PDF Preview: shows generated receipts or selected TR 12A PDFs in the right-side panel.
- Print: prints the PDF currently shown in the preview panel.
- Select Folder: choose the default folder where saved BIMS TR 12A PDF sheets are stored.
- Generate: reads the PDFs in that folder, renames each file using the parsed Full Name when available, and displays parsed details when the folder contains only one PDF.

Searchable PDFs are parsed directly. Scanned/image PDFs are parsed through the Windows OCR engine when available.

## Google Drive Login

1. Create or open a project in Google Cloud Console.
2. Enable the Google Drive API and Google Sheets API.
3. Configure the OAuth consent screen and add the Google account as a test user when required.
4. Create an OAuth client with application type **Desktop app** and download its JSON file.
5. Open **Settings > Drive Settings**, select the JSON under **OAuth Client JSON**, then click **Login to Google Drive**.

The browser handles the Google sign-in. The local login is saved in `google_drive_token.json`; this file is excluded from source control. After login, **Load Names** can read the configured private Google Sheet URL.
