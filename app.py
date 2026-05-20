import re
import json
from io import BytesIO

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Prime Conduit PDF Extractor",
    page_icon="📊",
    layout="wide",
)

BOOK4_HEADERS = [
    "Distname", "Supplier_name", "direct_indirect", "in_out_territory", "CustAccNbr",
    "CustDunsID", "CustName", "Address1", "City", "State", "County", "Zip", "Phone",
    "Country", "NoOfEmployees", "WebAddress", "SIC", "NAICS", "LineOfBusiness",
    "ParentName", "AccountType", "UOM", "InvoiceNumber", "Qty", "UnitCost", "UnitResale",
    "InvoiceDate", "DateRecieved", "PartNumberSubmitted", "PartNumberDescription", "Branch",
    "SalesRep", "Latitude", "Longitude", "Brand", "PartNumberActual", "UPCCode", "rawcustname",
    "rawdistaddress", "rawdistcity", "rawdiststate", "rawdistpostalcode", "rawdistcountry",
    "currency", "contractID", "client_CustName", "Zip_4_digit", "dnb_trade_style",
    "dnb_sales_value", "google_CustName", "google_Address1", "google_State", "google_Zip",
    "google_Country", "google_Phone", "google_WebAddress", "Pay_Month", "Pay_Year",
    "Ship_Month", "Ship_Year", "Industry", "Commissions", "Commission_Rate", "Cust_AM",
    "CEM", "Sales", "In_Out", "Commission_split_percentage", "Distributor_part_number",
    "Category", "google_City", "Billings", "Cheque_Number", "Pay_Date", "meta_data_json",
    "SO_Number", "PO_Number", "ship_date", "searched_on_google"
]

PREVIEW_COLUMNS = [
    "Supplier_name", "CustAccNbr", "CustName", "City", "State", "InvoiceNumber",
    "Qty", "UnitCost", "Commissions", "SO_Number", "PO_Number"
]

INVOICE_RE = re.compile(r"\b9\d{7,8}\b")
AMOUNT_RE = re.compile(r"-?\d{1,3}(?:,\d{3})*(?:\.\d{2})|-?\d+\.\d{2}")
ACCOUNT_RE = re.compile(r"^(\d{5,6})\s*-\s*(.*)$")
STATE_CODES = set("AL AK AZ AR CA CO CT DE FL GA IA ID IL IN KS KY LA MA MD ME MI MN MO MS MT NC ND NE NH NJ NM NV NY OH OK OR PA RI SC SD TN TX UT VA VT WA WI WV WY".split())
SKIP_PREFIXES = (
    "ZSDB0010", "Agency Name", "Agency No", "Agreement No", "Sales Organization",
    "Distribution Channel", "Customer Name", "City / State", "Ship To Party",
    "Customer Totals", "Commission Totals", "Group Description", "Freight / Basis", "Frt Allow"
)


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def to_number(value):
    if value in (None, ""):
        return ""
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return value


def to_int(value):
    try:
        return int(value)
    except Exception:
        return value or ""


def blank_row():
    return {header: "" for header in BOOK4_HEADERS}


def parse_order_line(line, current_city, current_state):
    tokens = clean_text(line).split()
    city = current_city
    state = current_state
    order_no = ""
    po_no = ""

    state_pos = None
    for i, token in enumerate(tokens):
        if token in STATE_CODES:
            state_pos = i

    if state_pos is not None:
        city = " ".join(tokens[:state_pos])
        state = tokens[state_pos]
        after = tokens[state_pos + 1:]
    else:
        after = tokens

    if len(after) >= 1:
        order_no = after[0]
    if len(after) >= 2 and not re.fullmatch(r"-?\d{1,3}(?:,\d{3})*(?:\.\d{2})|-?\d+\.\d{2}", after[1]):
        po_no = after[1]

    return city, state, order_no, po_no


def extract_prime_conduit_rows(uploaded_pdf):
    all_lines = []

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_lines.extend([clean_text(line) for line in text.splitlines() if clean_text(line)])

    rows = []
    current_customer = ""
    current_city = ""
    current_state = ""

    for index, line in enumerate(all_lines):
        if line.startswith(SKIP_PREFIXES):
            continue

        invoice_match = INVOICE_RE.search(line)
        if not invoice_match:
            continue

        invoice_number = invoice_match.group(0)
        customer_before_invoice = clean_text(line[:invoice_match.start()])
        amount_text = clean_text(line[invoice_match.end():])
        amounts = AMOUNT_RE.findall(amount_text)

        if len(amounts) < 4:
            continue

        if customer_before_invoice:
            current_customer = customer_before_invoice

        unit_cost = to_number(amounts[2])
        commission = to_number(amounts[-1])
        commission_rate = ""
        rate_match = re.search(r"(\d+(?:\.\d+)?)%", amount_text)
        if rate_match:
            commission_rate = f"{rate_match.group(1)}%"

        city = current_city
        state = current_state
        order_no = ""
        po_no = ""

        if index + 1 < len(all_lines):
            city, state, order_no, po_no = parse_order_line(all_lines[index + 1], current_city, current_state)
            if city:
                current_city = city
            if state:
                current_state = state

        account_number = ""
        ship_to_name = ""
        if index + 2 < len(all_lines):
            account_match = ACCOUNT_RE.match(all_lines[index + 2])
            if account_match:
                account_number = account_match.group(1)
                ship_to_name = clean_text(account_match.group(2))

        row = blank_row()
        row.update({
            "Supplier_name": "Prime",
            "CustAccNbr": to_int(account_number),
            "CustName": current_customer,
            "City": city,
            "State": state,
            "InvoiceNumber": to_int(invoice_number),
            "Qty": 1,
            "UnitCost": unit_cost,
            "Commissions": commission,
            "Commission_Rate": commission_rate,
            "SO_Number": order_no,
            "PO_Number": po_no,
            "rawcustname": ship_to_name,
            "currency": "USD",
            "meta_data_json": json.dumps({
                "ship_to_name": ship_to_name,
                "source_line": line,
                "order_line": all_lines[index + 1] if index + 1 < len(all_lines) else "",
                "ship_to_line": all_lines[index + 2] if index + 2 < len(all_lines) else "",
            })
        })
        rows.append(row)

    # Remove exact duplicate invoice rows if PDF text is repeated.
    unique = []
    seen = set()
    for row in rows:
        key = (row["InvoiceNumber"], row["CustAccNbr"], row["SO_Number"], row["UnitCost"], row["Commissions"])
        if key not in seen:
            seen.add(key)
            unique.append(row)

    return unique


def create_excel(rows):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    ws.append(BOOK4_HEADERS)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        ws.append([row.get(header, "") for header in BOOK4_HEADERS])

    currency_cols = ["UnitCost", "Commissions", "Billings"]
    int_cols = ["CustAccNbr", "InvoiceNumber", "Qty", "SO_Number"]

    for col_index, header in enumerate(BOOK4_HEADERS, start=1):
        letter = get_column_letter(col_index)
        width = max(12, min(32, len(header) + 2))
        if header in ["CustName", "rawcustname", "meta_data_json"]:
            width = 34
        ws.column_dimensions[letter].width = width

        for cell in ws[letter][1:]:
            if header in currency_cols:
                cell.number_format = '#,##0.00'
            elif header in int_cols:
                cell.number_format = '0'

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(output)
    output.seek(0)
    return output


st.title("📊 Prime Conduit PDF Extractor")
st.write("Upload the Prime Conduit commission report PDF and download a Book4-style Excel output.")

uploaded_files = st.file_uploader(
    "Upload Prime Conduit PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("Extract Data", type="primary"):
        all_rows = []

        with st.spinner("Extracting PDF data..."):
            for uploaded_file in uploaded_files:
                try:
                    rows = extract_prime_conduit_rows(uploaded_file)
                    all_rows.extend(rows)
                    st.success(f"✅ {uploaded_file.name}: {len(rows)} rows extracted")
                except Exception as error:
                    st.error(f"❌ {uploaded_file.name}: {error}")

        if all_rows:
            st.subheader("Preview")
            st.dataframe(
                [{col: row.get(col, "") for col in PREVIEW_COLUMNS} for row in all_rows],
                use_container_width=True,
            )

            excel_file = create_excel(all_rows)
            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_file,
                file_name="prime_conduit_extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No invoice rows found. Make sure this is a text-based Prime Conduit ZSDB0010 PDF, not a scanned image.")
else:
    st.info("Upload a PDF to begin.")
