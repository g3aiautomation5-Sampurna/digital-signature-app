"""
Accept both .xlsx and .ods file uploads
Load data from either format using the existing load_spreadsheet_data() function
Immediately convert to a new Excel (.xlsx) workbook
Process the spreadsheet (add signature)
Download as .xlsx with proper MIME type
"""

import streamlit as st
import pandas as pd
import hashlib
import base64
import os
import tempfile
import zipfile
import io
import subprocess
import shutil
from datetime import datetime

from openpyxl import cell, load_workbook, Workbook
from odf.opendocument import load as load_ods
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf import teletype

from cryptography.hazmat.primitives.asymmetric import rsa, padding
from cryptography.hazmat.primitives import serialization, hashes


import re

# =========================================================
# FIND APPROVER
# =========================================================

def find_approver(file_path):
    data = load_spreadsheet_data(file_path)

    pattern = r"^\s*(.*?)\s*:\s*(ok|approved)\s*$"

    for sheet_name, rows in data.items():

        if sheet_name == "Digital Signature":
            continue

        if not rows or not rows[0]:
            continue

        value = rows[0][0]

        if value is None:
            continue

        value = str(value)

        match = re.match(
            pattern,
            value,
            re.IGNORECASE
        )

        if match:
            approver_name = match.group(1).strip()
            return approver_name

    return None

# =========================================================
# HASH GENERATION
# =========================================================

def generate_hash_from_excel(file_path):
    data = load_spreadsheet_data(file_path)

    collected = []

    for sheet_name in sorted(data.keys()):
        if sheet_name == "Digital Signature":
            continue

        rows = data[sheet_name]
        collected.append(f"---SHEET:{sheet_name}---")

        for row in rows:
            if row is None:
                continue
            for value in row:
                if value is None:
                    value = ""

                value = str(value).strip()
                value = value.replace("\r", "")
                value = value.replace("\n", " ")

                try:
                    numeric = float(value)
                    if numeric.is_integer():
                        value = str(int(numeric))
                except:
                    pass

                collected.append(value)

    content = "|".join(collected)
    hash_value = hashlib.sha256(content.encode("utf-8")).hexdigest()
    return hash_value


# =========================================================
# ODS SUPPORT HELPERS
# =========================================================


def is_ods_file(file_path):
    return os.path.splitext(str(file_path))[1].lower() == ".ods"


def get_ods_cell_text(cell):

    text = teletype.extractText(cell)

    if text is not None:
        text = text.strip()

    if text:
        return text

    value = cell.getAttribute("value")

    if value is not None:
        return str(value)

    return ""


def load_ods_data(file_path):

    doc = load_ods(file_path)

    data = {}

    for table in doc.spreadsheet.getElementsByType(Table):

        sheet_name = table.getAttribute("name")

        rows = []

        for row in table.getElementsByType(TableRow):

            values = []

            for cell in row.getElementsByType(TableCell):

                # Handle repeated columns
                repeat = cell.getAttribute("numbercolumnsrepeated")
                repeat = int(repeat) if repeat else 1

                cell_value = get_ods_cell_text(cell)

                for _ in range(repeat):
                    values.append(cell_value)

            # Remove unnecessary trailing empty cells
            while values and values[-1] == "":
                values.pop()

            # Handle repeated rows
            row_repeat = row.getAttribute("numberrowsrepeated")
            row_repeat = int(row_repeat) if row_repeat else 1

            for _ in range(row_repeat):
                rows.append(values.copy())

        data[sheet_name] = rows

    return data


def _convert_with_excel(input_path, output_path):
    import win32com.client

    excel = None
    workbook = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(
            os.path.abspath(input_path),
            ReadOnly=True
        )

        workbook.SaveAs(
            os.path.abspath(output_path),
            FileFormat=51
        )
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def _find_libreoffice_command():
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\OpenOffice\program\soffice.exe",
        r"C:\Program Files (x86)\OpenOffice\program\soffice.exe"
    ]

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate

    return None


def _convert_with_libreoffice(input_path, output_path):
    libreoffice_cmd = _find_libreoffice_command()

    if not libreoffice_cmd:
        raise RuntimeError("LibreOffice is not available on this system.")

    command = [
        libreoffice_cmd,
        "--headless",
        "--convert-to",
        "xlsx",
        "--outdir",
        os.path.dirname(os.path.abspath(output_path)),
        os.path.abspath(input_path)
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed: {result.stderr.strip() or result.stdout.strip()}"
        )

    converted_name = os.path.splitext(os.path.basename(input_path))[0] + ".xlsx"
    converted_path = os.path.join(
        os.path.dirname(os.path.abspath(output_path)),
        converted_name
    )

    if not os.path.exists(converted_path):
        raise RuntimeError("LibreOffice conversion did not produce an XLSX file.")

    os.replace(converted_path, output_path)


def convert_ods_to_xlsx(input_path, output_path):
    try:
        _convert_with_excel(input_path, output_path)
        return
    except Exception:
        pass

    try:
        _convert_with_libreoffice(input_path, output_path)
        return
    except Exception:
        pass

    try:
        _convert_with_python(input_path, output_path)
        return
    except Exception as exc:
        raise RuntimeError(
            "Unable to convert ODS file to XLSX."
        ) from exc


def _convert_with_python(input_path, output_path):
    data = load_ods_data(input_path)
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, rows in data.items():
        ws = wb.create_sheet(sheet_name)
        for row_idx, row in enumerate(rows, 1):
            for col_idx, value in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx).value = value

    wb.save(output_path)



def load_spreadsheet_data(file_path):
    if is_ods_file(file_path):
        return load_ods_data(file_path)

    wb = load_workbook(file_path, data_only=False)
    data = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []

        for row in ws.iter_rows():
            rows.append([cell.value for cell in row])

        data[sheet_name] = rows

    return data


# =========================================================
# GENERATE NEW KEY PAIR
# =========================================================

def generate_new_keys():

    private_key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048
    )

    public_key = private_key.public_key()

    private_key_pem = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption()
    )

    public_key_pem = public_key.public_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PublicFormat.SubjectPublicKeyInfo
    )

    return private_key_pem, public_key_pem



# =========================================================
# SIGN HASH
# =========================================================

def sign_hash(hash_value, private_key):

    signature = private_key.sign(

        hash_value.encode(),

        padding.PKCS1v15(),

        hashes.SHA256()
    )

    signature_b64 = base64.b64encode(signature).decode()

    return signature_b64


# =========================================================
# STORE SIGNATURE IN EXCEL
# =========================================================

def store_signature(
    file_path,
    signature_b64,
    approver_name
):
    # File is already in .xlsx format at this point
    wb = load_workbook(file_path)

    if "Digital Signature" in wb.sheetnames:
        del wb["Digital Signature"]

    ws = wb.create_sheet("Digital Signature")

    crc_value = hashlib.sha256(
        signature_b64.encode()
    ).hexdigest()[:16]

    ws["A1"] = crc_value
    ws["A2"] = signature_b64

    current_date = datetime.now().strftime(
        "%d-%m-%Y"
    )

    ws["A3"] = (
        f"Approved by {approver_name} "
        f"on {current_date}"
    )

    wb.save(file_path)





# =========================================================
# STREAMLIT UI
# =========================================================

st.set_page_config(
    page_title="Digital Signature Generator",
    page_icon="✍️",
    layout="centered"
)

st.title("Digital Signature Generator")

uploaded_file = st.file_uploader(
    "Upload Spreadsheet File",
    type=["xlsx", "ods"]
)


uploaded_private_key = st.file_uploader(
    "Upload Private Key (.pem)",
    type=["pem"]
)

generate_new = st.button("Generate And Download New Keys")


# =========================================================
# GENERATE NEW KEY
# =========================================================

if generate_new:

    private_key_pem, public_key_pem = generate_new_keys()

    public_key_text = public_key_pem.decode()

    st.success("New key pair generated")

    st.subheader("PUBLIC KEY")

    st.text_area(
        "Share this public key with receiver",
        public_key_text,
        height=250
    )

    # =========================================
    # CREATE ZIP IN MEMORY
    # =========================================

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        zip_file.writestr(
            "private_key.pem",
            private_key_pem
        )

        zip_file.writestr(
            "public_key.pem",
            public_key_pem
        )

    zip_buffer.seek(0)

    # =========================================
    # SINGLE DOWNLOAD BUTTON
    # =========================================

    st.download_button(
        label="Download Keys ZIP",
        data=zip_buffer,
        file_name="RSA_Keys.zip",
        mime="application/zip"
    )



if uploaded_file is not None and uploaded_private_key is not None:

    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    if file_ext not in [".xlsx", ".ods"]:
        file_ext = ".xlsx"

    # Create temp file with original extension to load the file
    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=file_ext
    ) as tmp:
        tmp.write(uploaded_file.read())
        temp_file_path = tmp.name

    if file_ext == ".ods":
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as tmp:
            temp_path = tmp.name

        try:
            convert_ods_to_xlsx(temp_file_path, temp_path)
        except RuntimeError as exc:
            try:
                os.remove(temp_file_path)
            except Exception:
                pass
            st.error(str(exc))
            st.stop()
        finally:
            try:
                os.remove(temp_file_path)
            except Exception:
                pass
    else:
        temp_path = temp_file_path

    try:

        # =========================================
        # CHECK APPROVAL
        # =========================================

        approver_name = find_approver(
            temp_path
        )

        if approver_name is None:

            st.error(
                "Approval not found. "
                "A1 of at least one sheet must contain "
                "'Name : OK' or 'Name : Approved'"
            )

            st.stop()



        # Load uploaded private key
        private_key = serialization.load_pem_private_key(
            uploaded_private_key.read(),
            password=None
        )

        # Generate hash
        hash_value = generate_hash_from_excel(
            temp_path
        )

        # Sign hash
        signature_b64 = sign_hash(
            hash_value,
            private_key
        )

        # Store signature
        store_signature(
            temp_path,
            signature_b64,
            approver_name
        )

        st.success(
            "Digital signature stored successfully"
        )

        with open(temp_path, "rb") as f:
            base, ext = os.path.splitext(uploaded_file.name)
            # Always download as .xlsx
            st.download_button(
                label="Download Signed Spreadsheet",
                data=f,
                file_name=f"{base}_Hashed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error(str(e))
