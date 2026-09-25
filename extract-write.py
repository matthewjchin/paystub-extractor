#!/usr/bin/env python3.12.2

import os
import re
import sys

import pandas as pd
from pypdf import PdfReader


OUTPUT_FILE = "paystubs.xlsx"
COLUMNS = ["Earnings", "FedWthg", "FedMED", "FedOASDI", "CAWthg", "CAOASDI", "Medical", \
    "AccIns", "CriIll", "DNE", "LFE", "VSE", "NetPay"]

# Each field is matched against paystub text using the first pattern that hits.
# Patterns capture the first dollar amount that follows the label on its line.

'''
----------NOTE TO USERS----------
The official labels that are used in paystubs may vary by employer or payroll organization
that the employer uses. Please be sure to check how pay and deductions are labeled in the 
paystub based on the PDF you have downloaded. 

If the labels are different, you may need to modify the regex patterns in the FIELD_PATTERNS 
dictionary below to match your statement's labels.
'''
FIELD_PATTERNS = {
    "FedWthg": [
        r"Fed(?:eral)?\s*With(?:holding|hldg|holdng)?[^\d\-]*([\d,]+\.\d{2})",
        r"Federal\s*Income\s*Tax[^\d\-]*([\d,]+\.\d{2})",
    ],
    "FedMED": [
        r"Fed\s*MED(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
        r"Medicare[^\d\-]*([\d,]+\.\d{2})",
    ],
    "FedOASDI": [
        r"Fed\s*OASDI(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
        r"Social\s*Security[^\d\-]*([\d,]+\.\d{2})",
    ],
    "CAWthg": [
        r"CA\s*With(?:holding|hldg|holdng)?[^\d\-]*([\d,]+\.\d{2})",
        r"State\s*Income\s*Tax[^\d\-]*([\d,]+\.\d{2})",
    ],
    "CAOASDI": [
        r"CA\s*OASDI(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
        r"CA\s*SDI[^\d\-]*([\d,]+\.\d{2})",
    ],
    # Patterns for medical benefits deduction
    "Medical":[
        r"1-MSE-BT(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
    ],
    "AccIns" : [
        r"Acc\s*Ins(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
    ],
    "CriIll" : [
        r"Cri\s*Ill(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
    ],
    # Dental benefits deduction
    "DNE" : [
        r"1-DNE-AT(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
    ],
    # Life insurance deduction
    "LFE" : [
        r"1-LFE-AT(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
    ],
    # Vision benefits deduction
    "VSE" : [
        r"1-VSE-AT(?:/EE)?[^\d\-]*([\d,]+\.\d{2})",
    ]
}

# Bottom summary row: "Current  <TotalGross>  <FedTaxableGross>  <TotalTaxes>  <TotalDeductions>  <NetPay>"
SUMMARY_ROW_PATTERN = (
    r"Current\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+"
    r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})"
)

EARNINGS_FALLBACK_PATTERNS = [
    r"(?:Gross\s*Pay|Total\s*Earnings|Gross\s*Earnings)[^\d\-]*([\d,]+\.\d{2})",
]
NET_PAY_FALLBACK_PATTERNS = [
    r"Net\s*Pay[^\d\-]*([\d,]+\.\d{2})",
    r"Take\s*Home\s*Pay[^\d\-]*([\d,]+\.\d{2})",
]


def extract_text(pdf_path):
    with PdfReader(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_field(text, patterns):
    for pattern in patterns:
        if match := re.search(pattern, text, re.IGNORECASE):
            return float(match[1].replace(",", ""))
    return None


def extract_paystub_values(pdf_path):
    text = extract_text(pdf_path)
    values = {field: extract_field(text, patterns) for field, patterns in FIELD_PATTERNS.items()}

    if summary := re.search(SUMMARY_ROW_PATTERN, text):
        earnings = float(summary[1].replace(",", ""))
        net_pay = float(summary[5].replace(",", ""))
    else:
        earnings = extract_field(text, EARNINGS_FALLBACK_PATTERNS)
        net_pay = extract_field(text, NET_PAY_FALLBACK_PATTERNS)

    values["Earnings"] = earnings
    values["NetPay"] = net_pay

    return values


def write_paystub_row(output_path, values):
    rounded_values = {
        col: (round(values[col], 2) if values[col] is not None else None) for col in COLUMNS
    }
    new_row = pd.DataFrame([rounded_values], columns=COLUMNS)
    if os.path.exists(output_path):
        existing = pd.read_excel(output_path)
        combined = pd.concat([existing, new_row], ignore_index=True)
    else:
        combined = new_row

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=len(COLUMNS)):
            for cell in row:
                cell.number_format = "0.00"


def main():
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Path to paystub PDF: ").strip()

    values = extract_paystub_values(pdf_path)

    if missing := [
        field for field, amount in values.items() if amount is None
    ]:
        print(f"Warning: Could not extract values for fields: {', '.join(missing)}")
    write_paystub_row(OUTPUT_FILE, values)
    print(f"Wrote paystub data to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
