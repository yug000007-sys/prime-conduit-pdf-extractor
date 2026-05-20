import re
from io import BytesIO

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Prime Conduit PDF Extractor", page_icon=":bar_chart:", layout="wide")

HEADERS = [
    "Supplier_name", "Distname", "CustName", "City", "State", "CustAccNbr",
    "InvoiceNumber", "PO_Number", "UnitCost", "Qty", "Commissions"
]

INVOICE_RE = re.compile(r"\b9\d{7,8}\b")
STATE_RE = re.compile(r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b")
ACCOUNT_RE = re.compile(r"\b\d{5,6}\b")
MONEY_RE = re.compile(r"-?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})|-?\$?\d+\.\d{2}")
PERCENT_RE = re.compile(r"\d+(?:\.\d+)?%")

CUSTOMER_WORDS = [
    "ECHO", "BORDER", "GRAYBAR", "CED", "VIKING", "VAN METER", "ELECTRIC",
    "SUPPLY", "SPLY", "WHOLESALE", "DISTRIBUTING", "INC", "LLC", "CO", "COMPANY"
]


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def to_number(value):
    if value in (None, ""):
        return ""
    value = str(value).replace("$", "").replace(",", "").strip()
    try:
        return float(value)
    except ValueError:
        return value


def looks_like_customer(text):
    upper = text.upper()
    if INVOICE_RE.search(upper):
        before_invoice = upper[:INVOICE_RE.search(upper).start()].strip()
        return len(before_invoice) >= 3
    return any(word in upper for word in CUSTOMER_WORDS) and not PERCENT_RE.search(upper)


def customer_from_line(line):
    match = INVOICE_RE.search(line)
    if match:
        candidate = clean(line[:match.start()])
        if candidate:
            return candidate
    return ""


def parse_city_state_account(text):
    state = ""
    city = ""
    account = ""

    state_match = STATE_RE.search(text)
    if state_match:
        state = state_match.group(1)
        before = text[:state_match.start()].strip().split()
        if before:
            city = before[-1]

    account_match = ACCOUNT_RE.search(text)
    if account_match:
        possible = account_match.group(0)
        if not possible.startswith("9"):
            account = possible

    return city, state, account


def parse_invoice_line(line, current_customer, distname=""):
    line = clean(line)
    invoices = INVOICE_RE.findall(line)
    if not invoices:
        return []

    customer = customer_from_line(line) or current_customer
    city, state, account = parse_city_state_account(line)

    money_values = MONEY_RE.findall(line)
    money_values = [m for m in money_values if not m.endswith("%")]

    # In the Prime report lines, the final dollar amount is usually commission.
    # The first amount after invoice is usually sale/unit value. The second is often qty or related value.
    unit_cost = to_number(money_values[0]) if len(money_values) >= 1 else ""
    qty = to_number(money_values[1]) if len(money_values) >= 2 else ""
    commission = to_number(money_values[-1]) if len(money_values) >= 1 else ""

    rows = []
    for invoice in invoices:
        rows.append({
            "Supplier_name": "Prime",
            "Distname": distname,
            "CustName": customer,
            "City": city,
            "State": state,
            "CustAccNbr": account,
            "InvoiceNumber": invoice,
            "PO_Number": "",
            "UnitCost": unit_cost,
            "Qty": qty,
            "Commissions": commission,
        })
    return rows


def extract_from_tables(page):
    rows = []
    try:
        tables = page.extract_tables() or []
    except Exception:
        return rows

    for table in tables:
        current_customer = ""
        for raw_row in table:
            cells = [clean(c) for c in raw_row if clean(c)]
            if not cells:
                continue
            line = " ".join(cells)
            found_customer = customer_from_line(line)
            if found_customer:
                current_customer = found_customer
            elif looks_like_customer(line) and not INVOICE_RE.search(line):
                current_customer = line
            rows.extend(parse_invoice_line(line, current_customer))
    return rows


def extract_from_text(page):
    text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
    rows = []
    current_customer = ""
    distname = ""

    for raw_line in text.splitlines():
        line = clean(raw_line)
        if not line:
            continue

        found_customer = customer_from_line(line)
        if found_customer:
            current_customer = found_customer
        elif looks_like_customer(line) and not INVOICE_RE.search(line):
            current_customer = line

        rows.extend(parse_invoice_line(line, current_customer, distname))

    return rows


def extract_pdf(uploaded_file):
    all_rows = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            table_rows = extract_from_tables(page)
            text_rows = extract_from_text(page)
            all_rows.extend(table_rows if len(table_rows) >= len(text_rows) else text_rows)

    deduped = []
    seen = set()
    for row in all_rows:
        key = (row["InvoiceNumber"], row["CustName"], row["UnitCost"], row["Commissions"])
        if key not in seen:
            seen.add(key)
            deduped.append(row)
    return deduped


def create_excel(rows):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    ws.append(HEADERS)

    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append([row.get(header, "") for header in HEADERS])

    for col in range(1, len(HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18

    for row in ws.iter_rows(min_row=2):
        for idx in [9, 10, 11]:
            row[idx - 1].number_format = '#,##0.00'

    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output


st.title("Prime Conduit PDF Extractor")
st.write("Upload Prime Conduit Commission Report PDFs and download structured invoice data as Excel.")

uploaded_files = st.file_uploader("Upload PDF file(s)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Extract Data", type="primary"):
        all_rows = []
        with st.spinner("Extracting data..."):
            for uploaded_file in uploaded_files:
                try:
                    rows = extract_pdf(uploaded_file)
                    all_rows.extend(rows)
                    st.success(f"{uploaded_file.name}: {len(rows)} rows extracted")
                except Exception as exc:
                    st.error(f"{uploaded_file.name}: {exc}")

        if all_rows:
            st.dataframe(all_rows, use_container_width=True)
            excel_file = create_excel(all_rows)
            st.download_button(
                label="Download Excel File",
                data=excel_file,
                file_name="prime_conduit_extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No invoice numbers found. If this is a scanned PDF, OCR will be needed.")
else:
    st.info("Upload one or more PDF files to begin.")
