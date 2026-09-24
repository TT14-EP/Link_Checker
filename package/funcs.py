import pandas as pd
import requests
import openpyxl
from pypdf import PdfReader
import os

import package


def notification(func):
    efname = input("Please input the name of the Excel file you want to check, e.g. EU KID Links 20251231.xlsx ")
    efpath = os.path.join(os.getcwd(), efname)

    def wrapper(*args, **kwargs):
        print(f">>Running link check for {efname} ")
        func(package.dlpath, efpath)
        print(f">>Finished link check for {efname}")

    return wrapper


def network_call(link):
    callobj = requests.get(str(link))
    code = callobj.status_code
    doc = "n/a" if code != 200 else callobj.headers["Content-Type"]
    ldate = "n/a" if code != 200 else callobj.headers["Last-Modified"]
    return code, doc, ldate, callobj


def pdf_check(dpath, fpath):
    reader = PdfReader(dpath + fpath)
    text = ""
    for page_num, page in enumerate(reader.pages):
        text = text + reader.pages[page_num].extract_text() + "\n"

    date_word = "The key information in this document is as of "
    isin_word = "ISIN: "

    if text.find(date_word) != -1:
        pos1 = text.find(date_word)  # for EU KIDS, UK KIDS
        d = text[pos1 + len(date_word):pos1 + len(date_word) + 10]
    else:
        pos1 = text.find("This Key Investor Information is accurate as at ")  # for UK KIIDS
        d = text[pos1 + len("This Key Investor Information is accurate as at "):pos1 + len("This Key Investor "
                                                                                           "Information is "
                                                                                           "accurate as at ") + 10]

    pos2 = text.find(isin_word)
    i = text[pos2 + len(isin_word):pos2 + len(isin_word) + 12]

    return d, i



@notification
def check(dpath, efpath):
    """Amend download folder in __init__ as needed."""

    pd.set_option('display.max_columns', None)  # Show all columns of dataframe in output panel
    pd.set_option('display.width', None)  # No wrapped columns of dataframe in output panel

    # Retrieve source list file and create dataframe
    df = pd.read_excel(efpath)

    # Create new columns in dataframe
    df["status_code"] = ""
    df["valid_link"] = ""
    df["doc_type"] = ""
    df["last_upload_date"] = ""
    df["as_of_date"] = ""
    df["pdf_isin"] = ""

    # Loop each row in dataframe to populate the newly created columns and download PDF
    for r in range(len(df)):
        print(r)
        try:
            call_result = network_call(str(list(df.iloc[r])[4]))
            df.at[r, "status_code"] = call_result[0]
            df.at[r, "valid_link"] = "Valid" if call_result[0] == 200 else "Invalid"
            df.at[r, "doc_type"] = call_result[1]
            df.at[r, "last_upload_date"] = call_result[2]

            if call_result[0] == 200 and call_result[1] == "application/pdf":
                filepath = "\\" + str(df.iloc[r, 2]).replace('"', "") + ".pdf"
                with open(dpath + filepath, "wb") as f:
                    f.write(call_result[3].content)
                    print("Download successful")

                pdf_result = pdf_check(dpath, filepath)
                df.at[r, "as_of_date"] = pdf_result[0]
                df.at[r, "pdf_isin"] = pdf_result[1]
            else:
                df.at[r, "as_of_date"] = "n/a"
                df.at[r, "pdf_isin"] = "n/a"

        except:
            # When the shown url in list is not a valid url, extract and use underlying url instead.
            print("exception")
            wb = openpyxl.load_workbook(efpath, data_only=True)
            targetcell = wb.worksheets[0].cell(row=r + 2, column=5)
            url = targetcell.hyperlink.target
            print(url)
            call_result = network_call(str(url))
            df.at[r, "status_code"] = call_result[0]
            df.at[r, "valid_link"] = "Valid" if call_result[0] == 200 else "Invalid"
            df.at[r, "doc_type"] = call_result[1]
            df.at[r, "last_upload_date"] = call_result[2]

            if call_result[0] == 200 and call_result[1] == "application/pdf":
                filepath = "\\" + str(df.iloc[r, 2]).replace('"', "") + ".pdf"
                with open(dpath + filepath, "wb") as f:
                    f.write(call_result[3].content)
                    print("Download successful")

                pdf_result = pdf_check(dpath, filepath)
                df.at[r, "as_of_date"] = pdf_result[0]
                df.at[r, "pdf_isin"] = pdf_result[1]
            else:
                df.at[r, "as_of_date"] = "n/a"
                df.at[r, "pdf_isin"] = "n/a"

    # Append populated dataframe to a new tab in source list file
    with pd.ExcelWriter(
            efpath,
            mode="a",
            engine="openpyxl",
            if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name="Output", index=False)