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
dependencies, and creates a **GDCK Admission** desktop shortcut:

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

## Publish to GitHub

Create an empty GitHub repository, then connect and push this local project:

```powershell
git remote add origin https://github.com/YOUR-USERNAME/GDCK_ADMISSION.git
git push -u origin main
```

After publishing, open **Settings > Updates** in the app and set **GitHub Repository**
to `YOUR-USERNAME/GDCK_ADMISSION`. On startup, the app checks the latest GitHub
release first, then the latest tag. If that version is newer than the installed
version, it shows an update window and opens the GitHub update page when the user
clicks **Yes**.

The application opens a main panel with two buttons:

- Academic Section
- Account Section

Clicking either button opens its own panel.

The Academic Section shows:

- Registration Summary: registered, completed, incomplete, admitted, and pending counts from the BDS 2026 response Sheet.
- Summary Report: lists students who took admission and students not yet admitted.
- Create Google Form: creates or updates the BDS 2026 admission Form and matching response Sheet.
- Add Dummy Data: adds two sample BDS 2026 students when they are not already present.
- View Submitted Admission: select a student name, view/edit the full record, and print the office data sheet or student copy.
- Mark Completed: marks the searched admission as admitted and sets all certificate checklist fields to Yes.

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
2. Enable the Google Drive API, Google Sheets API, and Google Forms API.
3. Configure the OAuth consent screen and add the Google account as a test user when required.
4. Create an OAuth client with application type **Desktop app** and download its JSON file.
5. Open **Settings > Drive Settings**, select the JSON under **OAuth Client JSON**, then click **Login to Google Drive**.

The browser handles the Google sign-in. The local login is saved in `google_drive_token.json`; this file is excluded from source control. After login, **Load Names** can read the configured private Google Sheet URL and **Create Google Form** can create the **Admission BDS 2026** Google Form plus a matching response Google Sheet in Drive. If either file already exists, the app reuses it and adds any missing BDS admission fields instead of creating a duplicate. Certificate upload fields are created as Drive/file-link fields because the official Google Forms API does not support creating File Upload questions programmatically. The response Sheet includes an automatic age column calculated from Date of Birth as on 31 Dec 2026.
The same **Create Google Form** action is also available in **Settings > Drive Settings** after Google Drive login.
When a Google Drive login is saved, the app checks for a **GDC Admission** folder
in My Drive on startup and creates it if it is missing.
