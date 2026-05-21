import re
from io import BytesIO

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(page_title="Prime Conduit PDF Extractor", page_icon="📊", layout="wide")

# Fixed output header. Removed unwanted columns:
# meta_data_json, Commission_Rate, currency, rawcustname
# SO_Number is kept blank if present in your template logic.
HEADERS = [
    "Distname", "Supplier_name", "direct_indirect", "in_out_territory",
    "CustAccNbr", "CustDunsNumber", "CustName", "Address1", "City", "State",
    "County", "Zip", "Phone", "Country", "NoOfEmployees", "WebAddress",
    "SIC", "NAICS", "LineOfBusiness", "ParentName", "AccountType", "UOM",
    "InvoiceNumber", "Qty", "UnitCost", "UnitResale", "InvoiceDate",
    "DateReceived", "PartNumber", "PartNumberDesc", "Branch", "SalesRep",
    "Latitude", "Longitude", "Brand", "PartNumber2", "UPCCode",
    "rawdistaddress", "rawdistcity", "rawdiststate", "rawdistpostal",
    "rawdistcountry", "contractID", "client_CustNumber", "Zip_4_digits",
    "dnb_tradestyle", "dnb_sales_volume", "google_CustName", "google_Address",
    "google_State", "google_Zip", "google_Country", "google_Phone",
    "google_WebAddress", "Pay_Month", "Pay_Year", "Ship_Month", "Ship_Year",
    "Industry", "Commissions", "CommissionRate", "Cust_AM", "CEM", "Sales",
    "In_Out", "CommissionNotes", "Distributor", "Category", "google_City",
    "Billings", "Cheque_Number", "Pay_Date", "SO_Number", "PO_Number",
    "ship_date", "searched_address"
]

INVOICE_RE = re.compile(r"\b9\d{7,8}\b")
ACCOUNT_RE = re.compile(r"^(\d{5,6})\s*-\s*(.+)$")
STATE_RE = re.compile(r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b")
MONEY_RE = re.compile(r"-?\d{1,3}(?:,\d{3})*(?:\.\d{2})-?")
DATE_RE = re.compile(r"Commission Customer Report For Period\s*:\s*(\d{2})/(\d{2})/(\d{4})\s*to\s*(\d{2})/(\d{2})/(\d{4})")

SKIP_WORDS = (
    "ZSDB0010", "Commission Customer Report", "Commission Group Report",
    "Agency Name", "Agency No.", "Agreement No.", "Sales Organization",
    "Distribution Channel", "Customer Name", "City / State", "Ship To Party",
    "Customer Totals", "Commission Totals", "Group Description", "Frt Allow",
    "Freight / Basis", "Page :", "Time :"
)


def clean(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()


def to_float(x):
    x = clean(x).replace("$", "").replace(",", "")
    if x.endswith("-"):
        x = "-" + x[:-1]
    if not x:
        return ""
    try:
        return float(x)
    except ValueError:
        return x


def is_skip(line):
    return any(w in line for w in SKIP_WORDS)


def is_group_header(line):
    line = clean(line)
    if not line or is_skip(line):
        return False
    if INVOICE_RE.search(line):
        return False
    if ACCOUNT_RE.search(line):
        return False
    if re.match(r"^\d{6,8}\s+", line):
        return False
    if re.match(r"^Manual Accruals", line):
        return False
    return bool(re.search(r"[A-Za-z]", line))


def parse_report_date(full_text):
    m = DATE_RE.search(full_text)
    if not m:
        return "", "", "", ""
    start_m, start_d, start_y, end_m, end_d, end_y = m.groups()
    invoice_date = f"{end_m}/{end_d}/{end_y}"
    pay_month = end_m
    pay_year = end_y
    return invoice_date, pay_month, pay_year, end_m, end_y


def parse_invoice_line(line):
    m = INVOICE_RE.search(line)
    if not m:
        return None

    invoice = m.group(0)
    dist_on_same_line = clean(line[:m.start()])
    rest = clean(line[m.end():])

    # Normal Prime customer rows after invoice:
    # Net Value, Discount/Freight, Commission Basis, Rate, Commission Total
    nums = MONEY_RE.findall(rest)

    net_value = ""
    commission_basis = ""
    commission_total = ""

    if len(nums) >= 5:
        net_value = to_float(nums[0])
        commission_basis = to_float(nums[2])
        commission_total = to_float(nums[-1])
    elif len(nums) >= 4:
        net_value = to_float(nums[0])
        commission_basis = to_float(nums[1])
        commission_total = to_float(nums[-1])
    elif nums:
        commission_total = to_float(nums[-1])

    return {
        "invoice": invoice,
        "dist_on_same_line": dist_on_same_line,
        "net_value": net_value,
        "commission_basis": commission_basis,
        "commission_total": commission_total,
    }


def parse_city_state_order_po(line):
    """
    Handles both:
    MINNEAPOLIS MN 2368091 P001829508 1,089.84
    2368092 P001829516 1,910.06
    """
    line = clean(line)
    city = ""
    state = ""
    order_no = ""
    po_no = ""

    sm = STATE_RE.search(line)
    rest = line

    if sm:
        city = clean(line[:sm.start()])
        state = sm.group(1)
        rest = clean(line[sm.end():])

    parts = rest.split()

    if parts and re.fullmatch(r"\d{6,8}", parts[0]):
        order_no = parts[0]

    if len(parts) >= 2:
        po_no = parts[1]

    return city, state, order_no, po_no


def find_next_account(lines, start_index):
    for j in range(start_index, min(start_index + 7, len(lines))):
        m = ACCOUNT_RE.search(clean(lines[j]))
        if m:
            return m.group(1), clean(m.group(2))
    return "", ""


def extract_rows_from_text(text):
    invoice_date, pay_month, pay_year, ship_month, ship_year = parse_report_date(text)
    lines = [clean(x) for x in text.splitlines() if clean(x)]

    rows = []
    current_dist = ""
    current_city = ""
    current_state = ""

    for i, line in enumerate(lines):
        if not line or is_skip(line):
            continue

        parsed = parse_invoice_line(line)

        if parsed:
            if parsed["dist_on_same_line"]:
                current_dist = parsed["dist_on_same_line"]

            next_line = clean(lines[i + 1]) if i + 1 < len(lines) else ""
            city, state, order_no, po_no = parse_city_state_order_po(next_line)

            # IMPORTANT FIX:
            # If city/state are not printed on every invoice row, carry forward the
            # last valid city/state from the customer block.
            if city:
                current_city = city
            if state:
                current_state = state

            cust_acc, cust_name = find_next_account(lines, i + 1)

            # IMPORTANT FIX:
            # CustName must come from "123456 - Customer Name" line.
            # Never leave CustName blank if account exists.
            if not cust_name:
                cust_name = current_dist

            row = {h: "" for h in HEADERS}

            row.update({
                "Distname": current_dist,
                "Supplier_name": "Prime",
                "direct_indirect": "Indirect",
                "in_out_territory": "",
                "CustAccNbr": cust_acc,
                "CustName": cust_name,
                "City": current_city,
                "State": current_state,
                "Country": "USA",
                "UOM": "EA",
                "InvoiceNumber": parsed["invoice"],
                "Qty": 1,
                # Book5 mapping: UnitCost should be commission basis
                "UnitCost": parsed["commission_basis"],
                "UnitResale": parsed["net_value"],
                "InvoiceDate": invoice_date,
                "Brand": "Prime Conduit",
                "Pay_Month": pay_month,
                "Pay_Year": pay_year,
                "Ship_Month": ship_month,
                "Ship_Year": ship_year,
                "Commissions": parsed["commission_total"],
                "CommissionRate": "",
                "Sales": "",
                "In_Out": "",
                "CommissionNotes": "",
                "Distributor": current_dist,
                "Billings": parsed["commission_basis"],
                "SO_Number": "",
                "PO_Number": po_no,
            })

            rows.append(row)
            continue

        if is_group_header(line):
            current_dist = line

            # Look ahead for a city/state line. This fixes blocks where only the
            # first invoice has city/state and following invoices omit it.
            for j in range(i + 1, min(i + 5, len(lines))):
                probe = clean(lines[j])
                if INVOICE_RE.search(probe):
                    city, state, _, _ = parse_city_state_order_po(clean(lines[j + 1]) if j + 1 < len(lines) else "")
                    if city:
                        current_city = city
                    if state:
                        current_state = state
                    break
                sm = STATE_RE.search(probe)
                if sm and not ACCOUNT_RE.search(probe):
                    current_city = clean(probe[:sm.start()])
                    current_state = sm.group(1)
                    break

    return rows


def extract_pdf(uploaded_files):
    all_rows = []
    for file in uploaded_files:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
                all_rows.extend(extract_rows_from_text(text))

    # Deduplicate safely
    seen = set()
    final_rows = []
    for r in all_rows:
        key = (r.get("InvoiceNumber"), r.get("CustAccNbr"), r.get("PO_Number"), r.get("Commissions"))
        if key not in seen:
            seen.add(key)
            final_rows.append(r)
    return final_rows


def create_excel(rows):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    ws.append(HEADERS)

    header_fill = PatternFill("solid", fgColor="FFFF00")
    header_font = Font(bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append([row.get(h, "") for h in HEADERS])

    for col_idx, header in enumerate(HEADERS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(header) + 2, 12), 28)

    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output


st.title("📊 Prime Conduit PDF Extractor")
st.write("Fixed version: CustName, City, and State are carried correctly through each customer block.")

uploaded_files = st.file_uploader("Upload Prime Conduit PDF file(s)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Extract Data", type="primary"):
        with st.spinner("Extracting data..."):
            try:
                rows = extract_pdf(uploaded_files)

                if rows:
                    st.success(f"Extracted {len(rows)} invoice rows")
                    st.dataframe(rows, use_container_width=True)

                    excel_data = create_excel(rows)
                    st.download_button(
                        label="⬇️ Download Excel",
                        data=excel_data,
                        file_name="prime_conduit_fixed_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No invoice rows found.")
            except Exception as e:
                st.error(f"Error: {e}")
else:
    st.info("Upload PDF file(s) to begin.")
