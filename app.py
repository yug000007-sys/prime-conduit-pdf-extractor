import re
from io import BytesIO

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Prime Conduit PDF Extractor", page_icon="📊", layout="wide")

# Header order matched to your Book5.xlsx manual file.
# These columns are intentionally left blank: meta_data_json, SO_Number,
# Commission_Rate, currency, and most enrichment/DNB/Google columns.
HEADERS = [
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

INVOICE_RE = re.compile(r"\b9\d{7,8}\b")
STATE_RE = re.compile(r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b")
ACCOUNT_LINE_RE = re.compile(r"^(\d{5,6})\s*-\s*(.+)$")
NUMBER_RE = re.compile(r"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?-?|-?\d+(?:\.\d+)?-?")

SKIP_LINES = (
    "ZSDB0010", "Commission Customer Report", "Commission Group Report", "Agency Name",
    "Agency No.", "Agreement No.", "Sales Organization", "Distribution Channel",
    "Customer Name", "City / State", "Ship To Party", "Customer Totals", "Commission Totals",
    "Group Description", "Manual Accruals", "Page :", "Time :"
)


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def to_number(value):
    value = clean(value).replace(",", "").replace("$", "")
    if not value:
        return ""

    is_negative = False
    if value.endswith("-"):
        is_negative = True
        value = value[:-1]
    if value.startswith("-"):
        is_negative = True
        value = value[1:]

    try:
        number = float(value)
        return -number if is_negative else number
    except ValueError:
        return ""


def is_skip_line(line):
    return any(word in line for word in SKIP_LINES)


def is_customer_group_header(line):
    """Detect customer group lines such as BORDER STATES or VIKING ELECTRIC SPLY INC."""
    line = clean(line)
    if not line or is_skip_line(line):
        return False
    if INVOICE_RE.search(line):
        return False
    if ACCOUNT_LINE_RE.search(line):
        return False
    if re.match(r"^\d{6,8}\b", line):
        return False
    # City/order lines look like: FARGO ND 2367330 4505851450 524.76
    if STATE_RE.search(line) and re.search(r"\b\d{6,8}\b", line):
        return False
    # Customer group lines can include numbers, e.g. CED 4662 or WESCO - FARGO ND 7852,
    # so allow digits if the line has letters and no invoice/order structure.
    return bool(re.search(r"[A-Za-z]", line))


def parse_invoice_line(line):
    """Return invoice, optional customer group from same line, commission basis, commission total."""
    invoice_match = INVOICE_RE.search(line)
    if not invoice_match:
        return None

    invoice_number = invoice_match.group(0)
    customer_on_line = clean(line[:invoice_match.start()])
    after_invoice = clean(line[invoice_match.end():])
    tokens = after_invoice.split()

    rate_index = None
    for index, token in enumerate(tokens):
        if token.endswith("%"):
            rate_index = index
            break

    unit_cost = ""
    commissions = ""

    if rate_index is not None:
        values_before_rate = []
        for token in tokens[:rate_index]:
            parsed = to_number(token)
            if parsed != "":
                values_before_rate.append(parsed)

        values_after_rate = []
        for token in tokens[rate_index + 1:]:
            parsed = to_number(token)
            if parsed != "":
                values_after_rate.append(parsed)

        # Manual Book5 maps UnitCost to Commission Basis, not Net Value.
        if values_before_rate:
            unit_cost = values_before_rate[-1]
        if values_after_rate:
            commissions = values_after_rate[0]
    else:
        values = [to_number(x) for x in NUMBER_RE.findall(after_invoice)]
        values = [x for x in values if x != ""]
        if len(values) >= 2:
            unit_cost = values[-2]
        if values:
            commissions = values[-1]

    return {
        "invoice_number": invoice_number,
        "customer_on_line": customer_on_line,
        "unit_cost": unit_cost,
        "commissions": commissions,
    }


def parse_city_state_order(line):
    """Parse the line below an invoice. It may contain City/State + order/PO/freight."""
    line = clean(line)
    city = ""
    state = ""

    state_match = STATE_RE.search(line)
    if state_match:
        city = clean(line[:state_match.start()])
        state = state_match.group(1)

    # In Book5, PO_Number is the first order number on this line.
    order_match = re.search(r"\b\d{6,8}\b", line)
    po_number = order_match.group(0) if order_match else ""

    return city, state, po_number


def find_account_line(lines, start_index):
    for j in range(start_index, min(start_index + 5, len(lines))):
        match = ACCOUNT_LINE_RE.match(clean(lines[j]))
        if match:
            return match.group(1), clean(match.group(2))
    return "", ""


def extract_rows_from_lines(lines):
    rows = []
    current_customer_group = ""
    current_city = ""
    current_state = ""

    i = 0
    while i < len(lines):
        line = clean(lines[i])

        if not line or is_skip_line(line):
            i += 1
            continue

        parsed_invoice = parse_invoice_line(line)
        if parsed_invoice:
            if parsed_invoice["customer_on_line"]:
                current_customer_group = parsed_invoice["customer_on_line"]

            next_line = clean(lines[i + 1]) if i + 1 < len(lines) else ""
            parsed_city, parsed_state, po_number = parse_city_state_order(next_line)

            if parsed_city and parsed_state:
                current_city = parsed_city
                current_state = parsed_state

            cust_acc_nbr, ship_to_name = find_account_line(lines, i + 1)

            row = {header: "" for header in HEADERS}
            row.update({
                "Supplier_name": "Prime",
                "CustAccNbr": int(cust_acc_nbr) if cust_acc_nbr else "",
                "CustName": current_customer_group,
                "City": current_city,
                "State": current_state,
                "InvoiceNumber": int(parsed_invoice["invoice_number"]),
                "Qty": 1,
                "UnitCost": parsed_invoice["unit_cost"],
                "Commissions": parsed_invoice["commissions"],
                "PO_Number": int(po_number) if str(po_number).isdigit() else po_number,
                # Leave these blank based on your correction request:
                "SO_Number": "",
                "Commission_Rate": "",
                "currency": "",
                "meta_data_json": "",
                "rawcustname": "",
            })
            rows.append(row)
            i += 1
            continue

        if is_customer_group_header(line):
            current_customer_group = line
            current_city = ""
            current_state = ""

        i += 1

    return rows


def extract_pdf(uploaded_file):
    all_rows = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            lines = [clean(line) for line in text.splitlines() if clean(line)]
            all_rows.extend(extract_rows_from_lines(lines))

    # De-duplicate exact repeated invoice rows without dropping legitimate credits.
    unique_rows = []
    seen = set()
    for row in all_rows:
        key = (
            row.get("InvoiceNumber"),
            row.get("CustAccNbr"),
            row.get("UnitCost"),
            row.get("Commissions"),
            row.get("PO_Number"),
        )
        if key not in seen:
            seen.add(key)
            unique_rows.append(row)

    return unique_rows


def create_excel(rows):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    ws.append(HEADERS)

    header_fill = PatternFill("solid", fgColor="FFFF00")
    header_font = Font(bold=True, color="000000")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in rows:
        ws.append([row.get(header, "") for header in HEADERS])

    money_columns = {"UnitCost", "UnitResale", "Commissions", "Billings"}
    integer_columns = {"CustAccNbr", "InvoiceNumber", "Qty", "PO_Number"}

    for col_index, header in enumerate(HEADERS, start=1):
        letter = get_column_letter(col_index)
        ws.column_dimensions[letter].width = min(max(len(header) + 2, 12), 26)

        if header in money_columns:
            for cell in ws[letter][1:]:
                cell.number_format = '#,##0.00;[Red]-#,##0.00'
        elif header in integer_columns:
            for cell in ws[letter][1:]:
                cell.number_format = '0'

    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output


st.title("📊 Prime Conduit PDF Extractor")
st.write("Upload Prime Conduit Commission Customer Report PDFs and download Excel using the Book5 header format.")

uploaded_files = st.file_uploader("Upload PDF file(s)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Extract Data", type="primary"):
        rows = []
        with st.spinner("Extracting data..."):
            for uploaded_file in uploaded_files:
                try:
                    file_rows = extract_pdf(uploaded_file)
                    rows.extend(file_rows)
                    st.success(f"✅ {uploaded_file.name}: {len(file_rows)} rows extracted")
                except Exception as error:
                    st.error(f"❌ {uploaded_file.name}: {error}")

        if rows:
            st.subheader("Preview")
            st.dataframe(rows, use_container_width=True)

            excel_data = create_excel(rows)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_data,
                file_name="prime_conduit_correct_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No invoice rows found. Make sure this is a selectable-text Prime Conduit commission PDF.")
else:
    st.info("Upload PDF file(s) to begin.")
