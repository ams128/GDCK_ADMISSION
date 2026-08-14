import json
import re
import shutil
import subprocess
import tempfile
import csv
import os
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from PIL import Image, ImageTk
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


PACKAGE_DIR = Path(__file__).resolve().parent
APP_NAME = "GDCK Admission"


def user_data_dir():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming"
        return Path(base) / APP_NAME
    if os.name == "posix":
        return Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share")) / "gdck-admission"
    return Path.home() / ".gdck-admission"


COURSES = ("BDS", "MDS", "DORA", "DCDM")
CATEGORIES = ("General", "SC", "ST", "OEC")
DATA_DIR = user_data_dir()
CONFIG_FILE = DATA_DIR / "settings.json"
LEGACY_CONFIG_FILE = PACKAGE_DIR.parent / "settings.json"
DEFAULT_CONFIG_FILE = PACKAGE_DIR / "default_settings.json"
RECEIPT_DIR = DATA_DIR / "receipts"
GOOGLE_DRIVE_TOKEN_FILE = DATA_DIR / "google_drive_token.json"
GOOGLE_DRIVE_SCOPES = (
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
)
DEFAULT_PRINTER_LABEL = "Windows default printer"
FIXED_ID_CARD_FEE = "150"
FEE_FIELDS = (
    "Tuition Fee",
    "Misc. Fee",
    "Caution Deposit (Refundable)",
    "Van Fee",
    "University Fee",
    "Verification Fee",
    "ID Card Fee",
    "Fee for SC/ST/OEC Students",
    "Total",
)
BASE_TOTAL_FIELDS = (
    "Tuition Fee",
    "Misc. Fee",
    "Caution Deposit (Refundable)",
    "Van Fee",
    "University Fee",
    "Verification Fee",
)
DEFAULT_FEES = {
    "BDS": {
        "Tuition Fee": "20840",
        "Misc. Fee": "1740",
        "Caution Deposit (Refundable)": "2320",
        "Van Fee": "1740",
        "University Fee": "2665",
        "Verification Fee": "100",
        "ID Card Fee": FIXED_ID_CARD_FEE,
        "Fee for SC/ST/OEC Students": "5085",
        "Total": "29555",
    },
    "MDS": {},
    "DORA": {},
    "DCDM": {},
}


def extract_pdf_text(pdf_path):
    text_parts = []
    try:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            text_parts.append(page.extract_text() or "")
    except Exception:
        pass

    text = "\n".join(part for part in text_parts if part).strip()
    if text:
        return text

    return extract_text_with_windows_ocr(pdf_path) or extract_text_with_tesseract(pdf_path)


def extract_text_with_windows_ocr(pdf_path):
    pdftoppm = shutil.which("pdftoppm")
    powershell = shutil.which("powershell")
    script_path = PACKAGE_DIR / "tools" / "windows_ocr.ps1"
    if not pdftoppm or not powershell or not script_path.exists():
        return ""

    with tempfile.TemporaryDirectory() as tmp:
        prefix = str(Path(tmp) / "page")
        subprocess.run(
            [pdftoppm, "-png", "-r", "250", str(pdf_path), prefix],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        text_parts = []
        for image_path in sorted(Path(tmp).glob("page-*.png")):
            completed = subprocess.run(
                [
                    powershell,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                    "-ImagePath",
                    str(image_path),
                ],
                check=False,
                capture_output=True,
                text=True,
            )
            if completed.stdout:
                text_parts.append(completed.stdout)
        return "\n".join(text_parts).strip()


def extract_text_with_tesseract(pdf_path):
    pdftoppm = shutil.which("pdftoppm")
    tesseract = shutil.which("tesseract")
    if not pdftoppm or not tesseract:
        return ""

    with tempfile.TemporaryDirectory() as tmp:
        prefix = str(Path(tmp) / "page")
        subprocess.run(
            [pdftoppm, "-png", "-r", "250", str(pdf_path), prefix],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        text_parts = []
        for image_path in sorted(Path(tmp).glob("page-*.png")):
            completed = subprocess.run(
                [tesseract, str(image_path), "stdout"],
                check=False,
                capture_output=True,
                text=True,
            )
            if completed.stdout:
                text_parts.append(completed.stdout)
        return "\n".join(text_parts).strip()


def parse_tr12a_pdf(pdf_path):
    text = normalize_text(extract_pdf_text(pdf_path))
    details = {
        "File": pdf_path.name,
        "GRN": find_grn(text),
        "Date": find_date(text),
        "Department": find_department(text),
        "Type of Payment": find_payment_type(text),
        "Office Name": find_office_name(text),
        "Year": find_year(text),
        "Full Name": find_full_name(text),
        "Account Head": find_account_head(text),
        "Amount": find_amount(text),
        "Remarks": find_field(text, r"\bREMARKS\s*\(If Any\)", (r"\bAmount in Words\b",)),
        "CIN No": find_cin(text),
        "Bank Branch/Treasury": find_bank_branch(text),
    }
    details = {key: value for key, value in details.items() if value}
    if not text:
        details["Note"] = "No searchable text or OCR text found."
    return details


def normalize_text(text):
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def find_field(text, label_pattern, stop_patterns):
    if not text:
        return ""
    stop = "|".join(stop_patterns)
    match = re.search(label_pattern + r"\s*:?\s*(.+?)(?=" + stop + r"|$)", text, re.IGNORECASE | re.DOTALL)
    if not match:
        return ""
    value = re.sub(r"\s+", " ", match.group(1)).strip(" :.-")
    return value


def find_amount(text):
    if not text:
        return ""
    total_match = re.search(r"\bTotal\b\s+([0-9,]+(?:\.\d+)?)", text, re.IGNORECASE)
    if total_match:
        return total_match.group(1)
    amount_match = re.search(r"\bAmount in Rs\.?\s*\n?\s*([0-9,]+(?:\.\d+)?)", text, re.IGNORECASE)
    if amount_match:
        return amount_match.group(1)
    amount_words_index = text.lower().find("amount in words")
    search_text = text[:amount_words_index] if amount_words_index > -1 else text
    amounts = re.findall(r"(?<![-0-9])([0-9]{2,6})(?![-0-9])", search_text)
    if amounts:
        return amounts[-1]
    return ""


def find_department(text):
    if re.search(r"\bMedical Education\b", text, re.IGNORECASE):
        return "Medical Education"
    return find_field(text, r"\bDepartment\b", (r"\bPayer Details\b", r"\bType of Payment\b"))


def find_payment_type(text):
    if re.search(r"\bTreasury Receipts\b", text, re.IGNORECASE):
        return "Treasury Receipts"
    return find_field(text, r"\bType of Payment\b", (r"\bTIN\b", r"\bOffice Name\b"))


def find_office_name(text):
    match = re.search(r"\bGOVT\s+DENTAL\s+COLLEGE\s+KOTTAYAM\b", text, re.IGNORECASE)
    if match:
        return match.group(0).upper()
    return find_field(text, r"\bOffice Name\b", (r"\bPAN No\b", r"\bLocation\b"))


def find_year(text):
    match = re.search(r"\b20[0-9]{2}-20[0-9]{2}\b", text)
    if match:
        return match.group(0)
    return find_field(text, r"\bYear\b", (r"\bFlat/Block\b", r"\bAccount Head Details\b"))


def find_account_head(text):
    match = re.search(r"\b[0-9]{4}-[0-9]{2}-[0-9]{3}-[0-9]{2}-[0-9]{2}\b", text)
    if match:
        return match.group(0)
    return find_field(text, r"\bAccount Head Details\b", (r"\bAmount in Rs\b",))


def find_grn(text):
    match = re.search(r"\bKL[0-9A-Z]{10,}\b", text, re.IGNORECASE)
    if match:
        return match.group(0).upper()
    return find_field(text, r"\bGRN\b", (r"\bBARCODE\b", r"\bDate\b"))


def find_date(text):
    match = re.search(r"\bDate\s+([0-9]{2}[-/][0-9]{2}[-/ ]?[0-9]{4})\b", text, re.IGNORECASE)
    if match:
        return match.group(1).replace(" ", "-")
    match = re.search(r"\b(20[0-9]{2}[-/][0-9]{2}[-/][0-9]{2})\b", text)
    if match:
        return match.group(1)
    return find_field(text, r"\bDate\b", (r"\bDept\.?Ref\b", r"\bDepartment\b"))


def find_cin(text):
    match = re.search(r"\b[0-9]{10,16}\b", text)
    if match:
        return match.group(0)
    return find_field(text, r"\bCIN No\b", (r"\bDate\b", r"\bBank Branch/Treasury\b"))


def find_full_name(text):
    value = find_field(text, r"\bFull Name\b", (r"\bFlat/Block\b", r"\bRoad/Street\b"))
    if value and looks_like_name(value):
        return value

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        if re.search(r"\bAccount Head Details\b", line, re.IGNORECASE) and index > 0:
            candidate = lines[index - 1]
            if looks_like_name(candidate):
                return candidate

    for line in lines:
        if looks_like_name(line):
            return line
    return ""


def find_bank_branch(text):
    if re.search(r"\bQR CODE PAYMENT\b", text, re.IGNORECASE):
        return "QR CODE PAYMENT GATEWAY"
    return find_field(text, r"\bBank Branch/Treasury\b", (r"\bScroll No\b",))


def looks_like_name(value):
    if not re.fullmatch(r"[A-Z][A-Z .]{4,80}", value):
        return False
    blocked_words = {
        "GOVERNMENT",
        "KERALA",
        "CHALAN",
        "BARCODE",
        "GRN",
        "DEPARTMENT",
        "TREASURY",
        "PAYMENT",
        "DETAILS",
        "COLLEGE",
        "KOTTAYAM",
        "RECEIPTS",
        "GATEWAY",
    }
    words = set(value.split())
    return not words.intersection(blocked_words)


def format_details(details):
    if not details:
        return "No details could be parsed."
    rows = ["Details from the single PDF file:"]
    for key, value in details.items():
        rows.append(f"{key}: {value}")
    return "\n".join(rows)


def clean_filename(name):
    name = re.sub(r"[<>:\"/\\|?*]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120] or "Unnamed"


def unique_pdf_path(path):
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def unique_text_path(path):
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def unique_path(path):
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def name_from_filename(path):
    stem = re.sub(r"[_-]+", " ", path.stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    ignored = {"sheet", "try", "tr", "12", "12a", "bims"}
    words = [word for word in stem.split() if word.lower() not in ignored]
    return " ".join(words)


def default_fee_settings():
    fees = {}
    for course in COURSES:
        course_defaults = DEFAULT_FEES.get(course, {})
        fees[course] = {field: course_defaults.get(field, "") for field in FEE_FIELDS}
    return normalize_fee_settings(fees)


def normalize_fee_settings(fees):
    for course in COURSES:
        course_fees = fees.setdefault(course, {})
        course_fees["ID Card Fee"] = FIXED_ID_CARD_FEE
        component_values = []
        for field in BASE_TOTAL_FIELDS:
            value = str(course_fees.get(field, "")).strip().replace(",", "")
            if value.isdigit():
                component_values.append(int(value))
        if len(component_values) == len(BASE_TOTAL_FIELDS):
            course_fees["Total"] = str(sum(component_values) + int(FIXED_ID_CARD_FEE))
    return fees


def load_settings():
    source = next(
        (path for path in (CONFIG_FILE, LEGACY_CONFIG_FILE, DEFAULT_CONFIG_FILE) if path.exists()),
        DEFAULT_CONFIG_FILE,
    )
    try:
        data = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        data = {}
    return data if isinstance(data, dict) else {}


def save_settings(data):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def load_default_folder():
    data = load_settings()
    folder = data.get("default_folder", "")
    return folder if isinstance(folder, str) else ""


def save_default_folder(folder):
    data = load_settings()
    data["default_folder"] = folder
    save_settings(data)


def load_fee_settings():
    fees = default_fee_settings()
    data = load_settings()
    saved_fees = data.get("fees", {})
    if isinstance(saved_fees, dict):
        for course in COURSES:
            saved_course = saved_fees.get(course, {})
            if isinstance(saved_course, dict):
                for field in FEE_FIELDS:
                    value = saved_course.get(field)
                    if value is not None:
                        fees[course][field] = str(value)
    return normalize_fee_settings(fees)


def save_fee_settings(fees):
    data = load_settings()
    data["fees"] = normalize_fee_settings(fees)
    save_settings(data)


def load_google_sheet_settings():
    data = load_settings()
    sheet_settings = data.get("google_sheet", {})
    if not isinstance(sheet_settings, dict):
        sheet_settings = {}
    return {
        "url": str(sheet_settings.get("url", "")),
        "name_column": str(sheet_settings.get("name_column", "Name") or "Name"),
    }


def save_google_sheet_settings(url, name_column):
    data = load_settings()
    data["google_sheet"] = {"url": url, "name_column": name_column or "Name"}
    save_settings(data)


def load_google_drive_settings():
    drive = load_settings().get("google_drive", {})
    if not isinstance(drive, dict):
        drive = {}
    return {"credentials_file": str(drive.get("credentials_file", ""))}


def save_google_drive_settings(credentials_file):
    data = load_settings()
    data["google_drive"] = {"credentials_file": credentials_file}
    save_settings(data)


def load_google_drive_credentials():
    if not GOOGLE_DRIVE_TOKEN_FILE.exists():
        return None
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    credentials = Credentials.from_authorized_user_file(str(GOOGLE_DRIVE_TOKEN_FILE), GOOGLE_DRIVE_SCOPES)
    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        GOOGLE_DRIVE_TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        GOOGLE_DRIVE_TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")
    return credentials if credentials.valid else None


def login_to_google_drive(credentials_file):
    credentials_path = Path(credentials_file)
    if not credentials_path.is_file():
        raise FileNotFoundError("Select the Desktop OAuth client JSON file first.")
    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), GOOGLE_DRIVE_SCOPES)
    credentials = flow.run_local_server(port=0, open_browser=True, prompt="consent")
    GOOGLE_DRIVE_TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
    GOOGLE_DRIVE_TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")
    return google_drive_user_email(credentials)


def google_drive_user_email(credentials=None):
    credentials = credentials or load_google_drive_credentials()
    if not credentials:
        return ""
    from googleapiclient.discovery import build

    drive = build("drive", "v3", credentials=credentials, cache_discovery=False)
    about = drive.about().get(fields="user(displayName,emailAddress)").execute()
    user = about.get("user", {})
    return user.get("emailAddress") or user.get("displayName") or "Google Drive account"


def load_printer_setting():
    printer = load_settings().get("printer", DEFAULT_PRINTER_LABEL)
    return str(printer or DEFAULT_PRINTER_LABEL)


def save_printer_setting(printer):
    data = load_settings()
    data["printer"] = printer or DEFAULT_PRINTER_LABEL
    save_settings(data)


def load_receipt_settings():
    data = load_settings()
    receipt = data.get("receipt", {})
    if not isinstance(receipt, dict):
        receipt = {}
    return {
        "admission_year": str(receipt.get("admission_year", datetime.now().year)),
        "principal_name": str(receipt.get("principal_name", "Manoj Joseph Michel") or "Manoj Joseph Michel"),
    }


def save_receipt_settings(admission_year, principal_name):
    data = load_settings()
    data["receipt"] = {
        "admission_year": admission_year or str(datetime.now().year),
        "principal_name": principal_name or "Manoj Joseph Michel",
    }
    save_settings(data)


def list_available_printers():
    powershell = shutil.which("powershell")
    if not powershell:
        return []
    try:
        result = subprocess.run(
            [
                powershell,
                "-NoProfile",
                "-Command",
                "Get-CimInstance Win32_Printer | Select-Object -ExpandProperty Name",
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        return []
    if result.returncode != 0:
        return []
    return sorted({line.strip() for line in result.stdout.splitlines() if line.strip()}, key=str.lower)


def category_fee_label(category):
    return "Fee for SC/ST/OEC Students" if category in {"SC", "ST", "OEC"} else "Total"


def format_fee_structure(course, category, fees):
    course_fees = fees.get(course, {})
    rows = [f"{course} Fee Structure", f"Category: {category}"]
    if category in {"SC", "ST", "OEC"}:
        special_fee = str(course_fees.get("Fee for SC/ST/OEC Students", "")).strip()
        if special_fee:
            rows.append("Fee for SC/ST/OEC Students")
            rows.append("(CD + University Fee + Verification Fee)")
            rows.append(f"Payable Fee: {special_fee}/-")
        else:
            rows.append("No SC/ST/OEC fee value set. Open Settings to enter value.")
        return "\n".join(rows)

    has_values = False
    for field in FEE_FIELDS:
        if field == "Fee for SC/ST/OEC Students":
            continue
        value = str(course_fees.get(field, "")).strip()
        if value:
            has_values = True
            rows.append(f"{field}: {value}/-")
    if not has_values:
        rows.append("No fee values set. Open Settings to enter values.")
    payable_field = category_fee_label(category)
    payable_value = str(course_fees.get(payable_field, "")).strip()
    if payable_value:
        rows.append(f"Payable Fee: {payable_value}/-")
    return "\n".join(rows)


def format_fee_summary(course, category, fees):
    lines = format_fee_structure(course, category, fees).splitlines()[2:]
    editable_fields = ("Tuition Fee:", "ID Card Fee:")
    return "\n".join(line for line in lines if not line.startswith(editable_fields))


def payable_fee(course, category, fees):
    course_fees = fees.get(course, {})
    return str(course_fees.get(category_fee_label(category), "")).strip()


def google_sheet_csv_url(url):
    url = url.strip()
    if not url:
        return ""
    if "docs.google.com/spreadsheets" not in url:
        return url
    match = re.search(r"/spreadsheets/d/([^/]+)", url)
    if not match:
        return url
    spreadsheet_id = match.group(1)
    parsed = urllib.parse.urlparse(url)
    query = urllib.parse.parse_qs(parsed.query)
    gid = query.get("gid", ["0"])[0]
    if not gid and parsed.fragment:
        fragment_query = urllib.parse.parse_qs(parsed.fragment.replace("gid=", "gid="))
        gid = fragment_query.get("gid", ["0"])[0]
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid or '0'}"


def sheet_id_and_gid(url):
    match = re.search(r"/spreadsheets/d/([^/]+)", url)
    if not match:
        return "", "0"
    parsed = urllib.parse.urlparse(url)
    gid_match = re.search(r"(?:^|[?#&])gid=(\d+)", url)
    return match.group(1), gid_match.group(1) if gid_match else "0"


def names_from_rows(rows, name_column):
    if not rows:
        return []
    headers = [str(value).strip() for value in rows[0]]
    lowered_column = name_column.strip().lower()
    name_index = next((i for i, header in enumerate(headers) if header.lower() == lowered_column), None)
    if name_index is None:
        name_index = next((i for i, header in enumerate(headers) if "name" in header.lower()), 0)
    names = []
    for row in rows[1:]:
        name = str(row[name_index]).strip() if name_index < len(row) else ""
        if name and name not in names:
            names.append(name)
    return names


def fetch_private_google_sheet_names(url, name_column, credentials):
    from googleapiclient.discovery import build

    spreadsheet_id, gid = sheet_id_and_gid(url)
    if not spreadsheet_id:
        raise ValueError("The Google Sheet URL is not valid.")
    sheets = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    metadata = sheets.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets.properties(sheetId,title)",
    ).execute()
    properties = [sheet.get("properties", {}) for sheet in metadata.get("sheets", [])]
    selected = next((item for item in properties if str(item.get("sheetId")) == gid), None)
    selected = selected or (properties[0] if properties else None)
    if not selected:
        return []
    sheet_title = str(selected.get("title", "Sheet1")).replace("'", "''")
    values = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_title}'",
    ).execute().get("values", [])
    return names_from_rows(values, name_column)


def fetch_student_names_from_sheet(url, name_column, credentials=None):
    if credentials and "docs.google.com/spreadsheets" in url:
        return fetch_private_google_sheet_names(url, name_column, credentials)
    csv_url = google_sheet_csv_url(url)
    if not csv_url:
        return []
    request = urllib.request.Request(csv_url, headers={"User-Agent": "GDCK Admission"})
    with urllib.request.urlopen(request, timeout=20) as response:
        content = response.read().decode("utf-8-sig")
    csv_rows = list(csv.reader(content.splitlines()))
    return names_from_rows(csv_rows, name_column)


def create_receipt_pdf(
    path,
    roll_number,
    student,
    course,
    category,
    fee_value,
    fee_details,
    admission_year,
    principal_name,
):
    pdf = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    margin = 50
    table_width = width - (margin * 2)
    navy = (0.08, 0.20, 0.29)
    light_blue = (0.91, 0.95, 0.98)
    border = (0.68, 0.73, 0.78)

    pdf.setFillColorRGB(*navy)
    pdf.rect(0, height - 92, width, 92, fill=1, stroke=0)
    pdf.setFillColorRGB(1, 1, 1)
    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(margin, height - 48, "GDCK Admission Fee Acknowledgement")
    pdf.setFont("Helvetica", 9)
    pdf.drawString(margin, height - 68, f"{admission_year} {course} Admission")

    def draw_section_title(title, y):
        pdf.setFillColorRGB(*navy)
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(margin, y, title)
        return y - 12

    def draw_detail_table(rows, top_y):
        row_height = 27
        label_width = 145
        bottom_y = top_y - (row_height * len(rows))
        pdf.setStrokeColorRGB(*border)
        pdf.rect(margin, bottom_y, table_width, row_height * len(rows), fill=0, stroke=1)
        pdf.line(margin + label_width, bottom_y, margin + label_width, top_y)
        for index, (label, value) in enumerate(rows):
            row_top = top_y - (index * row_height)
            row_bottom = row_top - row_height
            if index:
                pdf.line(margin, row_top, margin + table_width, row_top)
            pdf.setFillColorRGB(*light_blue)
            pdf.rect(margin, row_bottom, label_width, row_height, fill=1, stroke=0)
            pdf.setFillColorRGB(*navy)
            pdf.setFont("Helvetica-Bold", 9)
            pdf.drawString(margin + 10, row_bottom + 9, label)
            pdf.setFillColorRGB(0.12, 0.15, 0.18)
            pdf.setFont("Helvetica", 9)
            pdf.drawString(margin + label_width + 12, row_bottom + 9, value)
        return bottom_y

    details_top = draw_section_title("Student Details", height - 122)
    detail_rows = [
        ("Student Name", student),
        ("Course", course),
        ("Roll Number", roll_number or "-"),
        ("Category", category),
        ("Date", datetime.now().strftime("%d-%m-%Y")),
    ]
    details_bottom = draw_detail_table(detail_rows, details_top)

    fee_lines = []
    pending_label = ""
    for raw_line in fee_details.splitlines():
        line = raw_line.strip()
        if not line or line.endswith("Fee Structure") or line.startswith("Category:"):
            continue
        if ":" in line:
            label, value = line.rsplit(":", 1)
            fee_lines.append((label.strip(), value.strip()))
            pending_label = ""
        elif line.startswith("(") and pending_label:
            fee_lines[-1] = (f"{fee_lines[-1][0]} {line}", fee_lines[-1][1])
            pending_label = ""
        else:
            fee_lines.append((line, f"{fee_value}/-" if line.startswith("Fee for ") else ""))
            pending_label = line

    fee_title_y = draw_section_title("Fee Details", details_bottom - 28)
    header_height = 27
    row_height = 25
    fee_bottom = fee_title_y - header_height - (row_height * len(fee_lines))
    amount_width = 125
    amount_x = margin + table_width - amount_width

    pdf.setStrokeColorRGB(*border)
    pdf.rect(margin, fee_bottom, table_width, fee_title_y - fee_bottom, fill=0, stroke=1)
    pdf.setFillColorRGB(*navy)
    pdf.rect(margin, fee_title_y - header_height, table_width, header_height, fill=1, stroke=0)
    pdf.setFillColorRGB(1, 1, 1)
    pdf.setFont("Helvetica-Bold", 9)
    pdf.drawString(margin + 10, fee_title_y - 18, "Fee Particular")
    pdf.drawString(amount_x + 10, fee_title_y - 18, "Amount")

    for index, (label, value) in enumerate(fee_lines):
        row_top = fee_title_y - header_height - (index * row_height)
        row_bottom = row_top - row_height
        pdf.setStrokeColorRGB(*border)
        pdf.line(margin, row_top, margin + table_width, row_top)
        pdf.line(amount_x, row_bottom, amount_x, row_top)
        if label == "Payable Fee":
            pdf.setFillColorRGB(*light_blue)
            pdf.rect(margin, row_bottom, table_width, row_height, fill=1, stroke=0)
            pdf.setStrokeColorRGB(*border)
            pdf.line(amount_x, row_bottom, amount_x, row_top)
            pdf.setFillColorRGB(*navy)
            pdf.setFont("Helvetica-Bold", 9)
        else:
            pdf.setFillColorRGB(0.12, 0.15, 0.18)
            pdf.setFont("Helvetica", 9)
        pdf.drawString(margin + 10, row_bottom + 8, label)
        pdf.drawString(amount_x + 10, row_bottom + 8, value)

    pdf.setFillColorRGB(*navy)
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawRightString(margin + table_width, max(fee_bottom - 25, 154), f"Total Payable: {fee_value}/-")

    signature_x = width - 220
    pdf.setStrokeColorRGB(*border)
    pdf.line(signature_x, 105, width - 60, 105)
    pdf.setFillColorRGB(*navy)
    principal_font_size = 10
    while pdf.stringWidth(principal_name, "Helvetica-Bold", principal_font_size) > 155 and principal_font_size > 7:
        principal_font_size -= 1
    pdf.setFont("Helvetica-Bold", principal_font_size)
    pdf.drawCentredString((signature_x + width - 60) / 2, 89, principal_name)
    pdf.setFont("Helvetica", 9)
    pdf.drawCentredString((signature_x + width - 60) / 2, 75, "Principal")

    pdf.setFont("Helvetica", 9)
    pdf.drawString(60, 45, "Generated by GDCK Admission")
    pdf.save()


class AdmissionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GDCK Admission")
        self.geometry("1220x720")
        self.minsize(1060, 680)

        self.configure(bg="#f5f7fb")
        self._configure_styles()

        self.selected_folder = tk.StringVar(value=load_default_folder())
        self.selected_course = tk.StringVar(value=COURSES[0])
        self.selected_category = tk.StringVar(value=CATEGORIES[0])
        self.selected_student = tk.StringVar()
        self.roll_number = tk.StringVar()
        self.fee_settings = load_fee_settings()
        self.google_sheet_settings = load_google_sheet_settings()
        self.google_drive_settings = load_google_drive_settings()
        self.selected_printer = load_printer_setting()
        self.receipt_settings = load_receipt_settings()
        self.student_names = []
        self.current_preview_pdf = None
        self.preview_image = None
        self.tuition_fee = tk.StringVar(
            value=self.fee_settings.get(self.selected_course.get(), {}).get("Tuition Fee", "")
        )
        self.id_card_fee = tk.StringVar(
            value=self.fee_settings.get(self.selected_course.get(), {}).get("ID Card Fee", "")
        )
        self.fee_text = tk.StringVar(
            value=format_fee_summary(self.selected_course.get(), self.selected_category.get(), self.fee_settings)
        )
        self.container = ttk.Frame(self, style="App.TFrame")
        self.container.pack(fill="both", expand=True)

        self.show_main_panel()

    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("App.TFrame", background="#f5f7fb")
        style.configure("Panel.TFrame", background="#ffffff", relief="flat")
        style.configure("Summary.TFrame", background="#eef5fb", relief="flat")
        style.configure(
            "Title.TLabel",
            background="#f5f7fb",
            foreground="#14324a",
            font=("Segoe UI", 22, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background="#f5f7fb",
            foreground="#526579",
            font=("Segoe UI", 11),
        )
        style.configure(
            "PanelTitle.TLabel",
            background="#ffffff",
            foreground="#14324a",
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "PreviewTitle.TLabel",
            background="#ffffff",
            foreground="#14324a",
            font=("Segoe UI", 14, "bold"),
        )
        style.configure(
            "Body.TLabel",
            background="#ffffff",
            foreground="#334155",
            font=("Segoe UI", 11),
        )
        style.configure(
            "Section.TLabel",
            background="#ffffff",
            foreground="#14324a",
            font=("Segoe UI", 11, "bold"),
        )
        style.configure(
            "Summary.TLabel",
            background="#eef5fb",
            foreground="#243b53",
            font=("Segoe UI", 9),
        )
        style.configure(
            "Muted.TLabel",
            background="#ffffff",
            foreground="#64748b",
            font=("Segoe UI", 9),
        )
        style.configure(
            "Action.TButton",
            font=("Segoe UI", 13, "bold"),
            padding=(18, 14),
            background="#1f6feb",
            foreground="#ffffff",
        )
        style.map(
            "Action.TButton",
            background=[("active", "#185abc"), ("pressed", "#174ea6")],
            foreground=[("active", "#ffffff")],
        )
        style.configure(
            "CompactAction.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(14, 9),
            background="#1f6feb",
            foreground="#ffffff",
        )
        style.map(
            "CompactAction.TButton",
            background=[("active", "#185abc"), ("pressed", "#174ea6")],
            foreground=[("active", "#ffffff")],
        )
        style.configure("Back.TButton", font=("Segoe UI", 10), padding=(10, 8))
        style.configure("Course.TCombobox", padding=(8, 6))
        style.configure("Input.TEntry", padding=(8, 6))

    def clear_container(self):
        for child in self.container.winfo_children():
            child.destroy()

    def show_main_panel(self):
        self.clear_container()

        header = ttk.Frame(self.container, style="App.TFrame")
        header.pack(fill="x", padx=36, pady=(32, 18))

        ttk.Label(header, text="GDCK Admission Process", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Admission management for BDS, MDS, DORA, and Dental Mechanic courses",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(6, 0))

        panel = ttk.Frame(self.container, style="Panel.TFrame", padding=28)
        panel.pack(fill="both", expand=True, padx=36, pady=(0, 36))

        ttk.Label(panel, text="Main Panel", style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(
            panel,
            text="Select a section to continue with admission work.",
            style="Body.TLabel",
        ).pack(anchor="w", pady=(8, 24))

        button_area = ttk.Frame(panel, style="Panel.TFrame")
        button_area.pack(expand=True)

        academic_button = ttk.Button(
            button_area,
            text="Academic Section",
            style="Action.TButton",
            command=self.show_academic_panel,
        )
        academic_button.grid(row=0, column=0, padx=16, pady=12, ipadx=28, ipady=18)

        account_button = ttk.Button(
            button_area,
            text="Account Section",
            style="Action.TButton",
            command=self.show_account_panel,
        )
        account_button.grid(row=0, column=1, padx=16, pady=12, ipadx=28, ipady=18)

        settings_button = ttk.Button(
            button_area,
            text="Settings",
            style="Action.TButton",
            command=self.open_settings_window,
        )
        settings_button.grid(row=0, column=2, padx=16, pady=12, ipadx=28, ipady=18)

    def select_default_folder(self):
        folder = filedialog.askdirectory(title="Select folder for BIMS TR 12A sheets")
        if folder:
            self.selected_folder.set(folder)
            save_default_folder(folder)

    def load_student_names(self):
        try:
            credentials = load_google_drive_credentials()
            names = fetch_student_names_from_sheet(
                self.google_sheet_settings.get("url", ""),
                self.google_sheet_settings.get("name_column", "Name"),
                credentials=credentials,
            )
        except Exception as exc:
            messagebox.showerror("Google Sheet Error", f"Could not load student names.\n\n{exc}")
            return
        self.student_names = names
        if hasattr(self, "student_dropdown") and self.student_dropdown.winfo_exists():
            self.student_dropdown.configure(values=self.student_names)
        if self.student_names and not self.selected_student.get():
            self.selected_student.set(self.student_names[0])
        self.write_details(f"Loaded {len(self.student_names)} student names from Google Sheet.")

    def generate_receipt(self):
        if not self.save_tuition_fee():
            return
        student = self.selected_student.get().strip()
        if not student:
            messagebox.showwarning("Student required", "Please select or enter a student name.")
            return

        course = self.selected_course.get()
        category = self.selected_category.get()
        fee_value = payable_fee(course, category, self.fee_settings)
        if not fee_value:
            messagebox.showwarning("Fee required", "Please set the fee value in Settings first.")
            return

        fee_details = format_fee_structure(course, category, self.fee_settings)
        RECEIPT_DIR.mkdir(parents=True, exist_ok=True)
        receipt_name = clean_filename(f"{student}_{course}_{category}_fee_acknowledgement")
        receipt_path = unique_path(RECEIPT_DIR / f"{receipt_name}.pdf")
        create_receipt_pdf(
            receipt_path,
            self.roll_number.get().strip(),
            student,
            course,
            category,
            fee_value,
            fee_details,
            self.receipt_settings["admission_year"],
            self.receipt_settings["principal_name"],
        )
        self.show_pdf_preview(receipt_path)
        self.write_details(f"Receipt PDF generated:\n{receipt_path}")

    def write_details(self, text):
        if not hasattr(self, "details_box") or not self.details_box.winfo_exists():
            return
        self.details_box.configure(state="normal")
        self.details_box.delete("1.0", "end")
        self.details_box.insert("1.0", text)
        self.details_box.configure(state="disabled")

    def generate_from_folder(self):
        folder_text = self.selected_folder.get().strip()
        if not folder_text:
            messagebox.showwarning("Folder required", "Please select the default folder first.")
            return

        folder = Path(folder_text)
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("Invalid folder", "The selected folder does not exist.")
            return

        pdf_files = sorted(folder.glob("*.pdf"))
        if not pdf_files:
            self.write_details("No PDF files found in the selected folder.")
            return
        if len(pdf_files) > 1:
            selected_file = filedialog.askopenfilename(
                title="Select one BIMS TR 12A PDF",
                initialdir=str(folder),
                filetypes=(("PDF files", "*.pdf"),),
            )
            if not selected_file:
                self.write_details("Multiple PDF files found. No file was selected.")
                return
            selected_path = Path(selected_file)
            if selected_path.suffix.lower() != ".pdf":
                messagebox.showerror("Invalid file", "Please select a PDF file.")
                return
            pdf_files = [selected_path]

        results = []
        parsed_items = []
        preview_pdf = None
        for pdf_file in pdf_files:
            details = parse_tr12a_pdf(pdf_file)
            parsed_items.append((pdf_file, details))
            full_name = details.get("Full Name") or name_from_filename(pdf_file)
            processed_pdf = pdf_file
            if full_name:
                target_path = folder / f"{clean_filename(full_name)}.pdf"
                if pdf_file.resolve() == target_path.resolve():
                    results.append(f"Already named: {pdf_file.name}")
                else:
                    new_path = unique_pdf_path(target_path)
                    pdf_file.rename(new_path)
                    processed_pdf = new_path
                    results.append(f"Renamed: {pdf_file.name} -> {new_path.name}")
            else:
                results.append(f"Skipped rename, full name not found: {pdf_file.name}")
            preview_pdf = processed_pdf

        if len(parsed_items) == 1:
            _, details = parsed_items[0]
            results.append("")
            results.append(format_details(details))
            results.append("")
            results.append(format_fee_structure(self.selected_course.get(), self.selected_category.get(), self.fee_settings))
        else:
            results.append("")
            results.append(f"Processed {len(pdf_files)} PDF files.")

        self.write_details("\n".join(results))
        if preview_pdf:
            self.show_pdf_preview(preview_pdf)

    def show_academic_panel(self):
        self.show_section_panel(
            title="Academic Section",
            description="Handle academic admission details, course selection, documents, and merit status.",
            fields=(
                "Student Name",
                "Course",
                "Application Number",
                "Qualification",
                "Document Status",
                "Admission Status",
            ),
        )

    def show_account_panel(self):
        self.clear_container()

        header = ttk.Frame(self.container, style="App.TFrame")
        header.pack(fill="x", padx=36, pady=(24, 14))

        ttk.Button(header, text="Back to Main Panel", style="Back.TButton", command=self.show_main_panel).pack(
            anchor="w"
        )
        ttk.Label(header, text="Account Section", style="Title.TLabel").pack(anchor="w", pady=(14, 0))

        panel = ttk.Frame(self.container, style="Panel.TFrame", padding=24)
        panel.pack(fill="both", expand=True, padx=36, pady=(0, 36))
        content = ttk.Frame(panel, style="Panel.TFrame")
        content.pack(fill="both", expand=True)
        left_panel = ttk.Frame(content, style="Panel.TFrame")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        right_panel = ttk.Frame(content, style="Panel.TFrame")
        right_panel.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        self.add_bims_controls(left_panel)
        self.add_preview_panel(right_panel)

    def show_section_panel(self, title, description, fields, include_bims=False):
        self.clear_container()

        header = ttk.Frame(self.container, style="App.TFrame")
        header.pack(fill="x", padx=36, pady=(28, 16))

        ttk.Button(header, text="Back to Main Panel", style="Back.TButton", command=self.show_main_panel).pack(
            anchor="w"
        )
        ttk.Label(header, text=title, style="Title.TLabel").pack(anchor="w", pady=(18, 0))
        ttk.Label(header, text=description, style="Subtitle.TLabel").pack(anchor="w", pady=(6, 0))

        panel = ttk.Frame(self.container, style="Panel.TFrame", padding=28)
        panel.pack(fill="both", expand=True, padx=36, pady=(0, 36))

        form = ttk.Frame(panel, style="Panel.TFrame")
        form.pack(fill="x", anchor="n")

        for row, label_text in enumerate(fields):
            ttk.Label(form, text=label_text, style="Body.TLabel").grid(
                row=row, column=0, sticky="w", padx=(0, 18), pady=9
            )
            if label_text == "Course":
                input_widget = ttk.Combobox(form, values=COURSES, style="Course.TCombobox", state="readonly")
                input_widget.current(0)
            else:
                input_widget = ttk.Entry(form, width=42)
            input_widget.grid(row=row, column=1, sticky="ew", pady=9)

        form.columnconfigure(1, weight=1)

        footer = ttk.Frame(panel, style="Panel.TFrame")
        footer.pack(fill="x", pady=(24, 0))
        ttk.Button(footer, text="Save Record", style="Action.TButton").pack(side="right")

        if include_bims:
            self.add_bims_controls(panel)

    def add_bims_controls(self, panel):
        overview = ttk.Frame(panel, style="Panel.TFrame")
        overview.pack(fill="x", pady=(0, 12))
        details_panel = ttk.Frame(overview, style="Panel.TFrame")
        details_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        summary_panel = ttk.Frame(overview, style="Panel.TFrame")
        summary_panel.grid(row=0, column=1, sticky="nsew")
        overview.columnconfigure(0, weight=3)
        overview.columnconfigure(1, weight=2)

        details_heading = ttk.Frame(details_panel, style="Panel.TFrame")
        details_heading.pack(fill="x", pady=(0, 8))
        ttk.Label(details_heading, text="Admission Details", style="Section.TLabel").pack(side="left")
        ttk.Button(
            details_heading,
            text="Settings",
            style="Back.TButton",
            command=self.open_settings_window,
        ).pack(side="right")

        details_form = ttk.Frame(details_panel, style="Panel.TFrame")
        details_form.pack(fill="x")
        ttk.Label(details_form, text="Course", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4
        )
        course_dropdown = ttk.Combobox(
            details_form,
            textvariable=self.selected_course,
            values=COURSES,
            style="Course.TCombobox",
            state="readonly",
            width=14,
        )
        course_dropdown.grid(row=0, column=1, sticky="ew", pady=4)
        course_dropdown.bind("<<ComboboxSelected>>", self.update_course_fee_controls)
        ttk.Label(details_form, text="Category", style="Body.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4
        )
        category_dropdown = ttk.Combobox(
            details_form,
            textvariable=self.selected_category,
            values=CATEGORIES,
            style="Course.TCombobox",
            state="readonly",
            width=14,
        )
        category_dropdown.grid(row=1, column=1, sticky="ew", pady=4)
        category_dropdown.bind("<<ComboboxSelected>>", lambda _event: self.update_fee_display())
        ttk.Label(details_form, text="Student Name", style="Body.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.student_dropdown = ttk.Combobox(
            details_form,
            textvariable=self.selected_student,
            values=self.student_names,
            style="Course.TCombobox",
        )
        self.student_dropdown.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=4)
        ttk.Button(details_form, text="Load", style="Back.TButton", command=self.load_student_names).grid(
            row=2, column=2, pady=4
        )
        ttk.Label(details_form, text="Roll Number", style="Body.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 10), pady=4
        )
        ttk.Entry(details_form, textvariable=self.roll_number, style="Input.TEntry").grid(
            row=3, column=1, columnspan=2, sticky="ew", pady=4
        )
        details_form.columnconfigure(1, weight=1)

        ttk.Label(summary_panel, text="Fee Summary", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        fee_box = ttk.Frame(summary_panel, style="Summary.TFrame", padding=(12, 9))
        fee_box.pack(fill="both", expand=True)
        tuition_row = ttk.Frame(fee_box, style="Summary.TFrame")
        tuition_row.pack(fill="x", pady=(0, 7))
        ttk.Label(tuition_row, text="Tuition Fee", style="Summary.TLabel").pack(side="left")
        ttk.Label(tuition_row, text="/-", style="Summary.TLabel").pack(side="right", padx=(5, 0))
        tuition_entry = ttk.Entry(
            tuition_row,
            textvariable=self.tuition_fee,
            style="Input.TEntry",
            width=11,
            justify="right",
        )
        tuition_entry.pack(side="right")
        tuition_entry.bind("<Return>", self.save_tuition_fee)
        tuition_entry.bind("<FocusOut>", self.save_tuition_fee)
        id_card_row = ttk.Frame(fee_box, style="Summary.TFrame")
        id_card_row.pack(fill="x", pady=(0, 7))
        ttk.Label(id_card_row, text="ID Card Fee", style="Summary.TLabel").pack(side="left")
        ttk.Label(id_card_row, text="/-", style="Summary.TLabel").pack(side="right", padx=(5, 0))
        id_card_entry = ttk.Entry(
            id_card_row,
            textvariable=self.id_card_fee,
            style="Input.TEntry",
            width=11,
            justify="right",
            state="readonly",
        )
        id_card_entry.pack(side="right")
        ttk.Label(fee_box, textvariable=self.fee_text, style="Summary.TLabel", justify="left").pack(anchor="w")

        ttk.Separator(panel).pack(fill="x", pady=(0, 10))
        ttk.Label(panel, text="BIMS TR 12A Sheet", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        folder_area = ttk.Frame(panel, style="Panel.TFrame")
        folder_area.pack(fill="x", pady=(0, 12))
        ttk.Entry(folder_area, textvariable=self.selected_folder, style="Input.TEntry").grid(
            row=0, column=0, sticky="ew", padx=(0, 10)
        )
        ttk.Button(folder_area, text="Browse", style="Back.TButton", command=self.select_default_folder).grid(
            row=0, column=1, padx=(0, 10)
        )
        ttk.Button(folder_area, text="Generate", style="Back.TButton", command=self.generate_from_folder).grid(
            row=0, column=2
        )
        folder_area.columnconfigure(0, weight=1)

        receipt_area = ttk.Frame(panel, style="Panel.TFrame")
        receipt_area.pack(fill="x", pady=(0, 12))
        ttk.Label(receipt_area, text="Fee Acknowledgement", style="Section.TLabel").pack(side="left")
        ttk.Button(
            receipt_area,
            text="Generate Receipt",
            style="CompactAction.TButton",
            command=self.generate_receipt,
        ).pack(side="right")

        ttk.Label(panel, text="Activity", style="Muted.TLabel").pack(anchor="w", pady=(0, 4))
        self.details_box = ScrolledText(
            panel,
            height=3,
            wrap="word",
            font=("Segoe UI", 9),
            background="#f8fafc",
            foreground="#475569",
            relief="flat",
            borderwidth=0,
            padx=8,
            pady=6,
        )
        self.details_box.pack(fill="both", expand=True)
        self.write_details("Load names from Google Sheet, select category, then generate receipt.")

    def add_preview_panel(self, panel):
        top_row = ttk.Frame(panel, style="Panel.TFrame")
        top_row.pack(fill="x", pady=(0, 10))
        ttk.Label(top_row, text="Receipt Preview", style="PreviewTitle.TLabel").pack(side="left")
        ttk.Button(top_row, text="Print", style="Back.TButton", command=self.print_current_pdf).pack(side="right")
        ttk.Button(top_row, text="Save PDF", style="Back.TButton", command=self.save_current_pdf).pack(
            side="right", padx=(0, 10)
        )

        self.preview_canvas = tk.Canvas(panel, bg="#f8fafc", highlightthickness=1, highlightbackground="#cbd5e1")
        self.preview_canvas.pack(fill="both", expand=True)
        self.preview_canvas.create_text(
            220,
            260,
            text="Generate or select a PDF to preview here.",
            fill="#526579",
            font=("Segoe UI", 11),
            width=320,
        )

    def show_pdf_preview(self, pdf_path):
        self.current_preview_pdf = Path(pdf_path)
        pdftoppm = shutil.which("pdftoppm")
        if not pdftoppm:
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(
                220,
                260,
                text=f"PDF ready:\n{self.current_preview_pdf}\n\nPreview needs Poppler pdftoppm.",
                fill="#526579",
                font=("Segoe UI", 10),
                width=340,
            )
            return

        with tempfile.TemporaryDirectory() as tmp:
            prefix = str(Path(tmp) / "preview")
            subprocess.run(
                [pdftoppm, "-png", "-singlefile", "-r", "120", str(self.current_preview_pdf), prefix],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            image_path = Path(tmp) / "preview.png"
            if not image_path.exists():
                self.preview_canvas.delete("all")
                self.preview_canvas.create_text(
                    220,
                    260,
                    text=f"Could not render preview.\n{self.current_preview_pdf}",
                    fill="#526579",
                    font=("Segoe UI", 10),
                    width=340,
                )
                return

            image = Image.open(image_path)
            self.preview_canvas.update_idletasks()
            canvas_width = max(self.preview_canvas.winfo_width(), 120)
            canvas_height = max(self.preview_canvas.winfo_height(), 120)
            image.thumbnail((max(canvas_width - 24, 96), max(canvas_height - 24, 96)))
            self.preview_image = ImageTk.PhotoImage(image)

        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(
            canvas_width // 2,
            canvas_height // 2,
            image=self.preview_image,
            anchor="center",
        )

    def save_current_pdf(self):
        if not self.current_preview_pdf or not self.current_preview_pdf.exists():
            messagebox.showwarning("No PDF", "Please generate or select a PDF first.")
            return

        destination = filedialog.asksaveasfilename(
            parent=self,
            title="Save Receipt PDF",
            initialfile=self.current_preview_pdf.name,
            defaultextension=".pdf",
            filetypes=(("PDF files", "*.pdf"),),
        )
        if not destination:
            return

        try:
            source = self.current_preview_pdf.resolve()
            target = Path(destination).resolve()
            if source != target:
                shutil.copy2(source, target)
            messagebox.showinfo("PDF Saved", f"Receipt saved to:\n{target}")
        except Exception as exc:
            messagebox.showerror("Save Error", f"Could not save PDF.\n\n{exc}")

    def print_current_pdf(self):
        if not self.current_preview_pdf or not self.current_preview_pdf.exists():
            messagebox.showwarning("No PDF", "Please generate or select a PDF first.")
            return
        try:
            if self.selected_printer and self.selected_printer != DEFAULT_PRINTER_LABEL:
                os.startfile(str(self.current_preview_pdf), "printto", f'"{self.selected_printer}"')
            else:
                os.startfile(str(self.current_preview_pdf), "print")
        except Exception as exc:
            messagebox.showerror("Print Error", f"Could not print PDF.\n\n{exc}")

    def update_fee_display(self):
        self.fee_text.set(
            format_fee_summary(self.selected_course.get(), self.selected_category.get(), self.fee_settings)
        )

    def update_course_fee_controls(self, _event=None):
        course = self.selected_course.get()
        self.tuition_fee.set(self.fee_settings.get(course, {}).get("Tuition Fee", ""))
        self.id_card_fee.set(FIXED_ID_CARD_FEE)
        self.fee_settings.setdefault(course, {})["ID Card Fee"] = FIXED_ID_CARD_FEE
        self.update_fee_display()

    def save_tuition_fee(self, _event=None):
        course = self.selected_course.get()
        raw_value = self.tuition_fee.get().strip().replace(",", "")
        if raw_value and not raw_value.isdigit():
            current_value = self.fee_settings.get(course, {}).get("Tuition Fee", "")
            self.tuition_fee.set(current_value)
            messagebox.showwarning("Invalid Tuition Fee", "Enter the tuition fee using numbers only.")
            return False

        self.tuition_fee.set(raw_value)
        self.id_card_fee.set(FIXED_ID_CARD_FEE)
        self.fee_settings.setdefault(course, {})["Tuition Fee"] = raw_value
        self.fee_settings[course]["ID Card Fee"] = FIXED_ID_CARD_FEE
        component_values = []
        for field in BASE_TOTAL_FIELDS:
            value = str(self.fee_settings[course].get(field, "")).strip().replace(",", "")
            if value.isdigit():
                component_values.append(int(value))
        if len(component_values) == len(BASE_TOTAL_FIELDS):
            self.fee_settings[course]["Total"] = str(sum(component_values) + int(FIXED_ID_CARD_FEE))
        save_fee_settings(self.fee_settings)
        self.update_fee_display()
        if not raw_value and _event is None:
            messagebox.showwarning("Tuition Fee Required", "Enter the tuition fee before generating a receipt.")
            return False
        return True

    def open_settings_window(self):
        settings_window = tk.Toplevel(self)
        settings_window.title("Settings")
        settings_window.geometry("680x600")
        settings_window.minsize(600, 600)
        settings_window.transient(self)
        settings_window.grab_set()

        wrapper = ttk.Frame(settings_window, style="Panel.TFrame", padding=20)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text="Settings", style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(wrapper, text="Configure printing, fees, and drive connections.", style="Body.TLabel").pack(
            anchor="w", pady=(6, 14)
        )

        notebook = ttk.Notebook(wrapper)
        notebook.pack(fill="both", expand=True)

        printer_frame = ttk.Frame(notebook, style="Panel.TFrame", padding=18)
        notebook.add(printer_frame, text="Printers")
        ttk.Label(printer_frame, text="Receipt Printer", style="PanelTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )
        printer_var = tk.StringVar(value=self.selected_printer or DEFAULT_PRINTER_LABEL)
        printer_dropdown = ttk.Combobox(
            printer_frame,
            textvariable=printer_var,
            style="Course.TCombobox",
            state="readonly",
        )
        printer_dropdown.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        def refresh_printers():
            printers = [DEFAULT_PRINTER_LABEL, *list_available_printers()]
            printer_dropdown.configure(values=printers)
            if printer_var.get() not in printers:
                printer_var.set(DEFAULT_PRINTER_LABEL)

        ttk.Button(
            printer_frame,
            text="Refresh",
            style="Back.TButton",
            command=refresh_printers,
        ).grid(row=1, column=1)
        ttk.Label(
            printer_frame,
            text="The Print button will send the displayed PDF to this printer.",
            style="Body.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(14, 0))
        printer_frame.columnconfigure(0, weight=1)
        refresh_printers()

        fee_frame = ttk.Frame(notebook, style="Panel.TFrame", padding=18)
        notebook.add(fee_frame, text="Fee Settings")
        selected_settings_course = tk.StringVar(value=self.selected_course.get())
        admission_year_var = tk.StringVar(value=self.receipt_settings["admission_year"])
        principal_name_var = tk.StringVar(value=self.receipt_settings["principal_name"])
        ttk.Label(fee_frame, text="Course", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 10), padx=(0, 10)
        )
        course_dropdown = ttk.Combobox(
            fee_frame,
            textvariable=selected_settings_course,
            values=COURSES,
            style="Course.TCombobox",
            state="readonly",
            width=18,
        )
        course_dropdown.grid(row=0, column=1, sticky="w", pady=(0, 10))
        ttk.Label(fee_frame, text="Admission Year", style="Body.TLabel").grid(
            row=0, column=2, sticky="w", pady=(0, 10), padx=(18, 10)
        )
        ttk.Entry(fee_frame, textvariable=admission_year_var, style="Input.TEntry", width=15).grid(
            row=0, column=3, sticky="ew", pady=(0, 10)
        )
        ttk.Label(fee_frame, text="Principal Name", style="Body.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 12), padx=(0, 10)
        )
        ttk.Entry(fee_frame, textvariable=principal_name_var, style="Input.TEntry").grid(
            row=1, column=1, columnspan=3, sticky="ew", pady=(0, 12)
        )
        ttk.Separator(fee_frame).grid(row=2, column=0, columnspan=4, sticky="ew", pady=(0, 8))

        fee_drafts = {
            course: {field: self.fee_settings.get(course, {}).get(field, "") for field in FEE_FIELDS}
            for course in COURSES
        }
        normalize_fee_settings(fee_drafts)
        fee_entry_vars = {field: tk.StringVar() for field in FEE_FIELDS}
        active_fee_course = {"value": selected_settings_course.get()}
        for index, field in enumerate(FEE_FIELDS):
            row = 3 + (index // 2)
            label_column = 0 if index % 2 == 0 else 2
            entry_column = label_column + 1
            ttk.Label(fee_frame, text=field, style="Body.TLabel").grid(
                row=row, column=label_column, sticky="w", pady=7, padx=(0 if label_column == 0 else 18, 10)
            )
            entry_state = "readonly" if field == "ID Card Fee" else "normal"
            ttk.Entry(fee_frame, textvariable=fee_entry_vars[field], style="Input.TEntry", state=entry_state).grid(
                row=row, column=entry_column, sticky="ew", pady=7
            )
        fee_frame.columnconfigure(1, weight=1)
        fee_frame.columnconfigure(3, weight=1)

        def store_active_fee_values():
            course = active_fee_course["value"]
            fee_drafts[course] = {field: fee_entry_vars[field].get().strip() for field in FEE_FIELDS}
            fee_drafts[course]["ID Card Fee"] = FIXED_ID_CARD_FEE
            normalize_fee_settings(fee_drafts)

        def load_selected_fee_values(_event=None):
            store_active_fee_values()
            course = selected_settings_course.get()
            active_fee_course["value"] = course
            fee_drafts[course]["ID Card Fee"] = FIXED_ID_CARD_FEE
            for field in FEE_FIELDS:
                fee_entry_vars[field].set(fee_drafts[course][field])

        course_dropdown.bind("<<ComboboxSelected>>", load_selected_fee_values)
        for field in FEE_FIELDS:
            fee_entry_vars[field].set(fee_drafts[active_fee_course["value"]][field])

        drive_frame = ttk.Frame(notebook, style="Panel.TFrame", padding=18)
        notebook.add(drive_frame, text="Drive Settings")
        folder_var = tk.StringVar(value=self.selected_folder.get())
        ttk.Label(drive_frame, text="BIMS TR 12A Folder", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", pady=6, padx=(0, 12)
        )
        ttk.Entry(drive_frame, textvariable=folder_var, style="Input.TEntry").grid(
            row=0, column=1, sticky="ew", pady=6, padx=(0, 10)
        )

        def browse_drive_folder():
            folder = filedialog.askdirectory(parent=settings_window, title="Select BIMS TR 12A folder")
            if folder:
                folder_var.set(folder)

        ttk.Button(drive_frame, text="Browse", style="Back.TButton", command=browse_drive_folder).grid(
            row=0, column=2, pady=6
        )
        ttk.Label(drive_frame, text="Google Sheet URL or CSV URL", style="Body.TLabel").grid(
            row=1, column=0, sticky="w", pady=6, padx=(0, 12)
        )
        sheet_url_var = tk.StringVar(value=self.google_sheet_settings.get("url", ""))
        ttk.Entry(drive_frame, textvariable=sheet_url_var, style="Input.TEntry").grid(
            row=1, column=1, columnspan=2, sticky="ew", pady=6
        )
        ttk.Label(drive_frame, text="Student Name Column", style="Body.TLabel").grid(
            row=2, column=0, sticky="w", pady=6, padx=(0, 12)
        )
        name_column_var = tk.StringVar(value=self.google_sheet_settings.get("name_column", "Name"))
        ttk.Entry(drive_frame, textvariable=name_column_var, style="Input.TEntry").grid(
            row=2, column=1, columnspan=2, sticky="ew", pady=6
        )
        ttk.Separator(drive_frame).grid(row=3, column=0, columnspan=3, sticky="ew", pady=14)
        ttk.Label(drive_frame, text="Google Drive Login", style="Section.TLabel").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )
        ttk.Label(drive_frame, text="OAuth Client JSON", style="Body.TLabel").grid(
            row=5, column=0, sticky="w", pady=6, padx=(0, 12)
        )
        credentials_file_var = tk.StringVar(value=self.google_drive_settings.get("credentials_file", ""))
        ttk.Entry(drive_frame, textvariable=credentials_file_var, style="Input.TEntry").grid(
            row=5, column=1, sticky="ew", pady=6, padx=(0, 10)
        )

        def browse_oauth_file():
            selected = filedialog.askopenfilename(
                parent=settings_window,
                title="Select Google Desktop OAuth client JSON",
                filetypes=(("JSON files", "*.json"),),
            )
            if selected:
                credentials_file_var.set(selected)

        ttk.Button(drive_frame, text="Browse", style="Back.TButton", command=browse_oauth_file).grid(
            row=5, column=2, pady=6
        )
        initial_drive_status = "Connected login saved" if GOOGLE_DRIVE_TOKEN_FILE.exists() else "Not connected"
        drive_status_var = tk.StringVar(value=initial_drive_status)
        ttk.Label(drive_frame, textvariable=drive_status_var, style="Muted.TLabel").grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(4, 8)
        )
        drive_buttons = ttk.Frame(drive_frame, style="Panel.TFrame")
        drive_buttons.grid(row=7, column=0, columnspan=3, sticky="w")

        def login_drive():
            credentials_file = credentials_file_var.get().strip()
            self.google_drive_settings = {"credentials_file": credentials_file}
            save_google_drive_settings(credentials_file)
            drive_status_var.set("Waiting for Google login in your browser...")
            settings_window.update_idletasks()
            try:
                account = login_to_google_drive(credentials_file)
            except Exception as exc:
                drive_status_var.set("Login failed")
                messagebox.showerror("Google Drive Login", f"Could not connect to Google Drive.\n\n{exc}")
                return
            drive_status_var.set(f"Connected: {account}")
            messagebox.showinfo("Google Drive", f"Connected to {account}.")

        def disconnect_drive():
            try:
                GOOGLE_DRIVE_TOKEN_FILE.unlink(missing_ok=True)
            except OSError as exc:
                messagebox.showerror("Google Drive", f"Could not remove the saved login.\n\n{exc}")
                return
            drive_status_var.set("Not connected")

        ttk.Button(drive_buttons, text="Login to Google Drive", style="CompactAction.TButton", command=login_drive).pack(
            side="left", padx=(0, 10)
        )
        ttk.Button(drive_buttons, text="Disconnect", style="Back.TButton", command=disconnect_drive).pack(side="left")
        drive_frame.columnconfigure(1, weight=1)

        def save_values():
            store_active_fee_values()
            for course in COURSES:
                self.fee_settings[course] = fee_drafts[course]
            save_fee_settings(self.fee_settings)
            self.selected_printer = printer_var.get() or DEFAULT_PRINTER_LABEL
            save_printer_setting(self.selected_printer)
            self.receipt_settings = {
                "admission_year": admission_year_var.get().strip() or str(datetime.now().year),
                "principal_name": principal_name_var.get().strip() or "Manoj Joseph Michel",
            }
            save_receipt_settings(
                self.receipt_settings["admission_year"],
                self.receipt_settings["principal_name"],
            )
            self.selected_folder.set(folder_var.get().strip())
            save_default_folder(self.selected_folder.get())
            self.google_sheet_settings = {
                "url": sheet_url_var.get().strip(),
                "name_column": name_column_var.get().strip() or "Name",
            }
            save_google_sheet_settings(
                self.google_sheet_settings["url"],
                self.google_sheet_settings["name_column"],
            )
            self.google_drive_settings = {"credentials_file": credentials_file_var.get().strip()}
            save_google_drive_settings(self.google_drive_settings["credentials_file"])
            self.update_course_fee_controls()
            settings_window.destroy()

        button_row = ttk.Frame(wrapper, style="Panel.TFrame")
        button_row.pack(fill="x", pady=(14, 0))
        ttk.Button(button_row, text="Cancel", style="Back.TButton", command=settings_window.destroy).pack(
            side="right", padx=(10, 0)
        )
        ttk.Button(button_row, text="Save Settings", style="Action.TButton", command=save_values).pack(side="right")


def main():
    try:
        app = AdmissionApp()
        app.mainloop()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
