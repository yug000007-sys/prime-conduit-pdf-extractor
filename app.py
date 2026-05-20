
import re
from io import BytesIO
import pdfplumber
import streamlit as st
from openpyxl import Workbook

st.set_page_config(page_title="Prime Conduit PDF Extractor", page_icon="📊")

HEADERS = [
    "Supplier_name","Distname","CustName","City","State",
    "CustAccNbr","InvoiceNumber","PO_Number","UnitCost","Qty","Commissions"
]

INVOICE_RE = re.compile(r"\b9\d{7,8}\b")

def extract_rows(text):
    rows = []
    lines = text.splitlines()

    for line in lines:
        invoices = INVOICE_RE.findall(line)

        if invoices:
            for invoice in invoices:
                rows.append({
                    "Supplier_name": "Prime",
                    "Distname": "",
                    "CustName": line[:50],
                    "City": "",
                    "State": "",
                    "CustAccNbr": "",
                    "InvoiceNumber": invoice,
                    "PO_Number": "",
                    "UnitCost": "",
                    "Qty": "",
                    "Commissions": ""
                })

    return rows

def create_excel(rows):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    ws.append(HEADERS)

    for row in rows:
        ws.append([row.get(h, "") for h in HEADERS])

    wb.save(output)
    output.seek(0)
    return output

st.title("📊 Prime Conduit PDF Extractor")

uploaded_files = st.file_uploader(
    "Upload PDF Files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Extract Data"):
        all_rows = []

        for uploaded_file in uploaded_files:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    all_rows.extend(extract_rows(text))

        if all_rows:
            st.success(f"Extracted {len(all_rows)} rows")

            excel_data = create_excel(all_rows)

            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_data,
                file_name="prime_conduit_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(all_rows)
        else:
            st.warning("No invoice numbers found.")
