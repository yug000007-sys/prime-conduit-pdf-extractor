import re
from io import BytesIO

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(page_title="Prime Conduit PDF Extractor", page_icon="📊", layout="wide")

FIXED_HEADERS = ['Distname', 'Supplier_name', 'direct_indirect', 'in_out_territory', 'CustAccNbr', 'CustDunsNumber', 'CustName', 'Address1', 'City', 'State', 'County', 'Zip', 'Phone', 'Country', 'NoOfEmployees', 'WebAddress', 'SIC', 'NAICS', 'LineOfBusiness', 'ParentName', 'AccountType', 'UOM', 'InvoiceNumber', 'Qty', 'UnitCost', 'UnitResale', 'InvoiceDate', 'DateReceived', 'PartNumber', 'PartNumberDesc', 'Branch', 'SalesRep', 'Latitude', 'Longitude', 'Brand', 'PartNumber2', 'UPCCode', 'rawcustname', 'rawdistaddress', 'rawdistcity', 'rawdiststate', 'rawdistpostal', 'rawdistcountry', 'contractID', 'client_CustNumber', 'Zip_4_digits', 'dnb_tradestyle', 'dnb_sales_volume', 'google_CustName', 'google_Address', 'google_State', 'google_Zip', 'google_Country', 'google_Phone', 'google_WebAddress', 'Pay_Month', 'Pay_Year', 'Ship_Month', 'Ship_Year', 'Industry', 'Commissions', 'CommissionRate', 'Cust_AM', 'CEM', 'Sales', 'In_Out', 'CommissionNotes', 'Distributor', 'Category', 'google_City', 'Billings', 'Cheque_Number', 'Pay_Date', 'searched_address']

INVOICE_RE = re.compile(r"\b9\d{7,8}\b")
STATE_RE = re.compile(r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b")
MONEY_RE = re.compile(r"-?\d{1,3}(?:,\d{3})*(?:\.\d{2})")
ACCOUNT_LINE_RE = re.compile(r"^(\d{5,6})\s*-\s*(.+)$")
HEADER_SKIP_WORDS = (
    "ZSDB0010", "Commission Customer Report", "Commission Group Report",
    "Agency Name", "Agency No.", "Agreement No.", "Sales Organization",
    "Distribution Channel", "Customer Name", "City / State", "Ship To Party",
    "Customer Totals", "Commission Totals", "Page :", "Time :",
    "Group Description"
)


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def to_number(value):
    value = clean(value).replace(",", "").replace("$", "")
    if not value:
        return ""
    try:
        return float(value)
    except ValueError:
        return value


def is_header_or_total(line):
    return any(word in line for word in HEADER_SKIP_WORDS)


def looks_like_customer_header(line):
    if is_header_or_total(line):
        return False
    if INVOICE_RE.search(line):
        return False
    if ACCOUNT_LINE_RE.search(line):
        return False
    if re.match(r"^\d{6,}\s+", line):
        return False
    return bool(re.search(r"[A-Za-z]", line))


def parse_invoice_line(line):
    invoice_match = INVOICE_RE.search(line)
    if not invoice_match:
        return None

    invoice = invoice_match.group(0)
    before_invoice = clean(line[:invoice_match.start()])
    after_invoice = clean(line[invoice_match.end():])
    nums = MONEY_RE.findall(after_invoice)

    unit_cost = ""
    unit_resale = ""
    commission = ""

    if len(nums) >= 5:
        unit_cost = to_number(nums[0])
        unit_resale = to_number(nums[2])
        commission = to_number(nums[-1])
    elif nums:
        commission = to_number(nums[-1])

    return {
        "invoice": invoice,
        "customer_on_line": before_invoice,
        "unit_cost": unit_cost,
        "unit_resale": unit_resale,
        "commission": commission,
    }


def parse_order_po_line(line):
    parts = clean(line).split()
    order_no = ""
    po_no = ""

    if parts and re.fullmatch(r"\d{6,8}", parts[0]):
        order_no = parts[0]
    if len(parts) >= 2:
        po_no = parts[1]

    return order_no, po_no


def extract_rows_from_lines(lines):
    rows = []
    current_dist = ""
    current_city = ""
    current_state = ""

    i = 0
    while i < len(lines):
        line = clean(lines[i])

        if not line or is_header_or_total(line):
            i += 1
            continue

        parsed = parse_invoice_line(line)

        if parsed:
            if parsed["customer_on_line"]:
                current_dist = parsed["customer_on_line"]

            next_line = clean(lines[i + 1]) if i + 1 < len(lines) else ""
            state_match = STATE_RE.search(next_line)
            city = current_city
            state = current_state

            if state_match:
                state = state_match.group(1)
                city = clean(next_line[:state_match.start()])
                current_city = city
                current_state = state

            order_no, po_no = parse_order_po_line(next_line)

            cust_acc = ""
            cust_name = ""
            for j in range(i + 1, min(i + 5, len(lines))):
                acc_match = ACCOUNT_LINE_RE.search(clean(lines[j]))
                if acc_match:
                    cust_acc = acc_match.group(1)
                    cust_name = clean(acc_match.group(2))
                    break

            if not cust_name:
                cust_name = current_dist

            row = {h: "" for h in FIXED_HEADERS}
            row.update({
                "Distname": current_dist,
                "Supplier_name": "Prime",
                "direct_indirect": "Indirect",
                "CustAccNbr": cust_acc,
                "CustName": cust_name,
                "City": city,
                "State": state,
                "Country": "USA",
                "InvoiceNumber": parsed["invoice"],
                "UnitCost": parsed["unit_cost"],
                "UnitResale": parsed["unit_resale"],
                "Brand": "Prime Conduit",
                "rawcustname": cust_name,
                "Commissions": parsed["commission"],
                "CommissionRate": "",
                "Distributor": current_dist,
                "Billings": parsed["unit_resale"],
            })

            rows.append(row)
            i += 1
            continue

        if looks_like_customer_header(line):
            current_dist = line
            if i + 1 < len(lines):
                nxt = clean(lines[i + 1])
                sm = STATE_RE.search(nxt)
                if sm and not INVOICE_RE.search(nxt):
                    current_city = clean(nxt[:sm.start()])
                    current_state = sm.group(1)

        i += 1

    return rows


def extract_pdf(uploaded_file):
    all_rows = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            lines = [clean(x) for x in text.splitlines() if clean(x)]
            all_rows.extend(extract_rows_from_lines(lines))

    seen = set()
    unique = []
    for row in all_rows:
        key = (row.get("InvoiceNumber"), row.get("CustAccNbr"), row.get("Commissions"))
        if key not in seen:
            seen.add(key)
            unique.append(row)

    return unique


def create_excel(rows):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    ws.append(FIXED_HEADERS)

    fill = PatternFill("solid", fgColor="FFFF00")
    font = Font(bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append([row.get(h, "") for h in FIXED_HEADERS])

    for idx, header in enumerate(FIXED_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = min(max(len(header) + 2, 12), 24)

    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output


st.title("📊 Prime Conduit PDF Extractor")
st.write("Upload Prime Conduit Commission Report PDF(s). Output uses your fixed header only.")

uploaded_files = st.file_uploader("Upload PDF file(s)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Extract Data", type="primary"):
        rows = []

        with st.spinner("Extracting data..."):
            for file in uploaded_files:
                try:
                    file_rows = extract_pdf(file)
                    rows.extend(file_rows)
                    st.success(f"✅ {file.name}: {len(file_rows)} rows extracted")
                except Exception as e:
                    st.error(f"❌ {file.name}: {e}")

        if rows:
            st.dataframe(rows, use_container_width=True)
            excel_file = create_excel(rows)

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_file,
                file_name="prime_conduit_fixed_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No invoice rows found.")
else:
    st.info("Upload PDF file(s) to begin.")
