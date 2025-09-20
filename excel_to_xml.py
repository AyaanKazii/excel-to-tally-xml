import pandas as pd
import xml.etree.ElementTree as ET
import xml.dom.minidom
import os
from datetime import datetime

company_name = "Your Company Name"
receipt_bank_ledger = "Your Bank Ledger Name"

def clean_amount(val):
    if pd.isnull(val):
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return None

def excel_to_tally_xml(excel_path, xml_path, mode="Sales", manual_date=None):
    df = pd.read_excel(excel_path)

    print(f"\nFile: {excel_path}")
    print(f"Columns detected ({mode} mode):")
    for c in df.columns:
        print(f"- {c}")

    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"

    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    request_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(request_desc, "REPORTNAME").text = "All Masters"
    static_vars = ET.SubElement(request_desc, "STATICVARIABLES")
    ET.SubElement(static_vars, "SVCURRENTCOMPANY").text = company_name

    request_data = ET.SubElement(import_data, "REQUESTDATA")

    if mode.lower() == "sales":
        col_party = "PARTYNAME"
        col_date = "DATE"
        col_vno = "VOUCHERNUMBER"
        col_narr = "Narration"
        col_total = "Total including interest"

        for idx, row in df.iterrows():
            try:
                party_name = str(row[col_party]).strip()
                narration = str(row[col_narr])
                amount = clean_amount(row[col_total])
                if amount is None:
                    continue
            except KeyError as e:
                print(f"Missing expected column: {e}")
                continue

            tally_msg = ET.SubElement(request_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
            voucher = ET.SubElement(tally_msg, "VOUCHER", {
                "VCHTYPE": "Sales",
                "ACTION": "Create",
                "OBJVIEW": "Invoice Voucher View"
            })

            ET.SubElement(voucher, "DATE").text = "20250401"
            ET.SubElement(voucher, "EFFECTIVEDATE").text = "20250401"
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = "Sales"
            ET.SubElement(voucher, "PARTYLEDGERNAME").text = party_name
            ET.SubElement(voucher, "BASICBASEPARTYNAME").text = party_name
            ET.SubElement(voucher, "PERSISTEDVIEW").text = "Invoice Voucher View"
            ET.SubElement(voucher, "ISINVOICE").text = "Yes"

            party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(party_entry, "LEDGERNAME").text = party_name
            ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(party_entry, "ISPARTYLEDGER").text = "Yes"
            ET.SubElement(party_entry, "AMOUNT").text = f"-{amount:.2f}"

            bill_alloc = ET.SubElement(party_entry, "BILLALLOCATIONS.LIST")
            vno_val = row[col_vno]
            if pd.notnull(vno_val):
                bill_name = str(int(float(vno_val)))
            else:
                bill_name = f"BILL-{idx+1}"

            ET.SubElement(bill_alloc, "NAME").text = bill_name
            ET.SubElement(bill_alloc, "BILLTYPE").text = "Agst Ref"
            ET.SubElement(bill_alloc, "AMOUNT").text = f"-{amount:.2f}"

            exclude = {col_party, col_date, col_vno, col_narr, col_total}
            charge_cols = []
            for c in df.columns:
                if c in exclude:
                    continue
                if pd.api.types.is_numeric_dtype(df[c]):
                    charge_cols.append(c)

            for col in df.columns:
                if "Service Charges" in col and col not in charge_cols and col not in exclude:
                    charge_cols.append(col)

            for charge in charge_cols:
                amt = clean_amount(row[charge])
                if amt is None or amt == 0:
                    continue

                ledger_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                ET.SubElement(ledger_entry, "LEDGERNAME").text = charge
                ET.SubElement(ledger_entry, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(ledger_entry, "ISPARTYLEDGER").text = "No"
                ET.SubElement(ledger_entry, "AMOUNT").text = f"{amt:.2f}"
                ET.SubElement(ledger_entry, "VATEXPAMOUNT").text = f"{amt:.2f}"

    else:
        col_map = {
            "Party Name": "FlatNo.",
            "Narration1": "Narration",
            "Narration2": "Narration.1",
            "Received Amount": "Debit",
            "Voucher Number": "Voucher Number",
            "Date": "Date"
        }

        for idx, row in df.iterrows():
            party_name = str(row[col_map["Party Name"]]).strip()
            narration1 = str(row[col_map["Narration1"]]).strip()
            narration2 = str(row[col_map["Narration2"]]).strip()
            narration = (narration1 + " " + narration2).strip()
            voucher_no = str(row[col_map["Voucher Number"]]).strip()

            if manual_date:
                date_str = manual_date
            else:
                try:
                    date_str = pd.to_datetime(row[col_map["Date"]]).strftime("%Y%m%d")
                except:
                    date_str = datetime.now().strftime("%Y%m%d")

            amount = clean_amount(row[col_map["Received Amount"]])
            if amount is None or amount == 0:
                continue

            tally_msg = ET.SubElement(request_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
            voucher = ET.SubElement(tally_msg, "VOUCHER", {
                "VCHTYPE": "Receipt",
                "ACTION": "Create",
                "OBJVIEW": "Accounting Voucher View"
            })

            ET.SubElement(voucher, "DATE").text = date_str
            ET.SubElement(voucher, "EFFECTIVEDATE").text = date_str
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = "Receipt"
            ET.SubElement(voucher, "VOUCHERNUMBER").text = voucher_no
            ET.SubElement(voucher, "PARTYLEDGERNAME").text = receipt_bank_ledger
            ET.SubElement(voucher, "PERSISTEDVIEW").text = "Accounting Voucher View"
            ET.SubElement(voucher, "ISINVOICE").text = "No"

            party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(party_entry, "LEDGERNAME").text = party_name
            ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(party_entry, "ISPARTYLEDGER").text = "Yes"
            ET.SubElement(party_entry, "AMOUNT").text = f"-{amount:.2f}"

            bill_alloc = ET.SubElement(party_entry, "BILLALLOCATIONS.LIST")
            ET.SubElement(bill_alloc, "NAME").text = f"RECEIPT-{idx+1}"
            ET.SubElement(bill_alloc, "BILLTYPE").text = "Agst Ref"
            ET.SubElement(bill_alloc, "AMOUNT").text = f"-{amount:.2f}"

            bank_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(bank_entry, "LEDGERNAME").text = receipt_bank_ledger
            ET.SubElement(bank_entry, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(bank_entry, "ISPARTYLEDGER").text = "No"
            ET.SubElement(bank_entry, "AMOUNT").text = f"{amount:.2f}"

    raw_xml = ET.tostring(envelope, "utf-8")
    parsed_xml = xml.dom.minidom.parseString(raw_xml)
    pretty_xml = parsed_xml.toprettyxml(indent="  ")

    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(pretty_xml)

    print(f"{mode} XML saved to: {xml_path}")
    try:
        os.startfile(xml_path)
    except:
        pass

if __name__ == "__main__":
    choice = input("Enter voucher type (Sales/Receipt): ").strip().lower()

    if choice == "sales":
        excel_file = r"input_sales.xlsx"
        xml_file = r"sales.xml"
        excel_to_tally_xml(excel_file, xml_file, mode="Sales")

    elif choice == "receipt":
        excel_file = r"input_receipt.xlsx"
        xml_file = r"receipt.xml"

        manual_date_choice = input("Do you want to enter the date manually? (y/n): ").strip().lower()
        if manual_date_choice == "y":
            while True:
                manual_date_str = input("Enter date (dd-mm-yyyy): ").strip()
                try:
                    manual_date = datetime.strptime(manual_date_str, "%d-%m-%Y").strftime("%Y%m%d")
                    print(f"Using manually entered date: {manual_date}")
                    break
                except ValueError:
                    print("Invalid date format. Please enter in dd-mm-yyyy format.")
        else:
            manual_date = None

        excel_to_tally_xml(excel_file, xml_file, mode="Receipt", manual_date=manual_date)

    else:
        print("Invalid choice. Please type 'Sales' or 'Receipt'.")
