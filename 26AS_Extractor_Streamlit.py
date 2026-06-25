import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="26AS Extractor", layout="wide")

st.title("26AS PDF → Excel Extractor")

uploaded_file = st.file_uploader(
    "Upload Form 26AS PDF",
    type=["pdf"]
)

if uploaded_file:

    full_text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # --------------------------------------------------
    # Assessee Details
    # --------------------------------------------------

    assessee = {}

    pan_match = re.search(
        r"Permanent Account Number \(PAN\)\s+([A-Z0-9]+)",
        full_text
    )

    name_match = re.search(
        r"Name of Assessee\s+(.+?)\n",
        full_text
    )

    fy_match = re.search(
        r"Financial Year\s+([0-9\-]+)",
        full_text
    )

    ay_match = re.search(
        r"Assessment Year\s+([0-9\-]+)",
        full_text
    )

    status_match = re.search(
        r"Current Status of PAN\s+(.+?)\s+Financial Year",
        full_text
    )

    address_match = re.search(
        r"Address of Assessee\s+(.+?)Above data",
        full_text,
        re.S
    )

    assessee["PAN"] = pan_match.group(1) if pan_match else ""
    assessee["Name"] = name_match.group(1).strip() if name_match else ""
    assessee["Financial Year"] = fy_match.group(1) if fy_match else ""
    assessee["Assessment Year"] = ay_match.group(1) if ay_match else ""
    assessee["PAN Status"] = status_match.group(1).strip() if status_match else ""

    if address_match:
        assessee["Address"] = (
            address_match.group(1)
            .replace("\n", " ")
            .strip()
        )
    else:
        assessee["Address"] = ""

    assessee_df = pd.DataFrame(
        list(assessee.items()),
        columns=["Field", "Value"]
    )

    # --------------------------------------------------
    # TDS SUMMARY
    # --------------------------------------------------

    tds_summary = []

    summary_pattern = re.compile(
        r'(\d+)\s+'
        r'([A-Z0-9 &.,()/-]+?)\s+'
        r'([A-Z]{4}\d{5}[A-Z])\s+'
        r'([\d,.]+)\s+'
        r'([\d,.]+)\s+'
        r'([\d,.]+)'
    )

    matches = summary_pattern.findall(full_text)

    for m in matches:

        sno = m[0]
        name = m[1].strip()
        tan = m[2]

        amount = float(m[3].replace(",", ""))
        tds = float(m[4].replace(",", ""))
        deposited = float(m[5].replace(",", ""))

        tds_summary.append([
            sno,
            "",
            name,
            tan,
            amount,
            tds,
            deposited
        ])

    tds_df = pd.DataFrame(
        tds_summary,
        columns=[
            "S.No",
            "Section",
            "Name",
            "TAN",
            "Amount Paid/Credited",
            "TDS Deducted",
            "TDS Deposited"
        ]
    )

    # --------------------------------------------------
    # TRANSACTIONS
    # --------------------------------------------------

    transactions = []

    current_name = ""
    current_tan = ""

    lines = full_text.split("\n")

    company_pattern = re.compile(
        r'([A-Z0-9 &.,()/-]+?)\s+([A-Z]{4}\d{5}[A-Z])'
    )

    transaction_pattern = re.compile(
        r'(\d+)\s+'
        r'(194[A-Z]*)\s+'
        r'(\d{2}-[A-Za-z]{3}-\d{4})\s+'
        r'([A-Z])\s+'
        r'(\d{2}-[A-Za-z]{3}-\d{4})\s+'
        r'-\s+'
        r'([\d,.]+)\s+'
        r'([\d,.]+)\s+'
        r'([\d,.]+)'
    )

    for line in lines:

        company_match = company_pattern.search(line)

        if company_match:
            current_name = company_match.group(1).strip()
            current_tan = company_match.group(2)

        trx_match = transaction_pattern.search(line)

        if trx_match:

            transactions.append([
                trx_match.group(1),
                current_name,
                current_tan,
                trx_match.group(2),
                trx_match.group(3),
                trx_match.group(4),
                trx_match.group(5),
                "-",
                float(trx_match.group(6).replace(",", "")),
                float(trx_match.group(7).replace(",", "")),
                float(trx_match.group(8).replace(",", ""))
            ])

    transaction_df = pd.DataFrame(
        transactions,
        columns=[
            "S.No",
            "Deductor Name",
            "TAN",
            "Section",
            "Transaction Date",
            "Booking Status",
            "Booking Date",
            "Remarks",
            "Amount Paid/Credited",
            "Tax Deducted",
            "TDS Deposited"
        ]
    )

    # --------------------------------------------------
    # PREVIEW
    # --------------------------------------------------

    st.subheader("Assessee Details")
    st.dataframe(assessee_df)

    st.subheader("TDS Summary")
    st.dataframe(tds_df)

    st.subheader("Transaction Details")
    st.dataframe(transaction_df)

    # --------------------------------------------------
    # EXCEL EXPORT
    # --------------------------------------------------

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        assessee_df.to_excel(
            writer,
            sheet_name="Assessee Details",
            index=False
        )

        tds_df.to_excel(
            writer,
            sheet_name="TDS Summary",
            index=False
        )

        transaction_df.to_excel(
            writer,
            sheet_name="Transaction Details",
            index=False
        )

    output.seek(0)

    st.download_button(
        "Download Excel",
        output,
        file_name="26AS_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )