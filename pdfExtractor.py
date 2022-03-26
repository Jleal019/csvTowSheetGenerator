import re
import os
import csv
from csv import writer
import fitz  # this is pymupdf

# This program requires the pymupdf dependency

# a_csv Defines function to append a row to the file.


def a_csv(row):
    with open("towReport.csv", 'a', newline='') as file:
        wr = writer(file)
        wr.writerow(row)
        file.close()


# analyze Defines function to find relevant fields and adds them to
# a list that will then pass the list to the a_csv function.


def analyze(path):

        with fitz.open(path) as doc:
            text = ""
            for page in doc:
                text += page.getText()
                # print(text)
                tow_company = re.search(r'ROADWAY INFORMATION\n(.*?)\nNAME OF LOCATION TOWED TO', text).group(1)
                address = re.search(r'WRECKER NAME\(IF DIFERENT\)\n(.*?)\nADDRESS', text).group(1)

                yr = re.search(r'OTHER COMMENTS, OTHER AGENCY CASE  #\n(.*?)\nVEH. YR.', text)
                if yr is None:
                    yr = ""
                else:
                    yr = re.search(r'OTHER COMMENTS, OTHER AGENCY CASE  #\n(.*?)\nVEH. YR.', text).group(1)

                make = re.search(r'REG. EXP. DATE\n(.*?)\nVEHICLE MAKE', text)
                if make is None:
                    make = ""
                else:
                    make = re.search(r'REG. EXP. DATE\n(.*?)\nVEHICLE MAKE', text).group(1)

                model = re.search(r'VEHICLE MAKE\n(.*?)\nMODEL', text)
                if model is None:
                    model = ""
                else:
                    model = re.search(r'VEHICLE MAKE\n(.*?)\nMODEL', text).group(1)

                clr = re.search(r'VEH. YR.\n(.*?|)\nVEHICLE COLOR', text).group(1)

                tag = re.search(r'VEHICLE COLOR\n(.*?)\nVEHICLE TAG #', text)
                if tag is None:
                    tag = ""
                else:
                    tag = re.search(r'VEHICLE COLOR\n(.*?)\nVEHICLE TAG #', text).group(1)

                tken1 = re.search(r'DRIVER OR LAST PERSON IN POSSESSION INFORMATION\n(.*?)\nADDRESS #', text)
                if tken1 is None:
                    tken1 = ""
                else:
                    tken1 = re.search(r'DRIVER OR LAST PERSON IN POSSESSION INFORMATION\n(.*?)\nADDRESS #', text).group(1)

                tken3 = re.search(r'ADDRESS #\n(.*?)\nTOW OCCURRED ON STREET, ROAD, HIGHWAY', text)
                if tken3 is None:
                    tken3 = ""
                else:
                    tken3 = re.search(r'ADDRESS #\n(.*?)\nTOW OCCURRED ON STREET, ROAD, HIGHWAY', text).group(1)

                tken2 = re.search(r'DIRECTION\n(.*?|)\nAT/FROM INTERSECTION WITH STREET, ROAD,HIGHWAY', text)
                if tken2 is None:
                    tken2 = ""
                else:
                    tken2 = re.search(r'DIRECTION\n(.*?|)\nAT/FROM INTERSECTION WITH STREET, ROAD,HIGHWAY', text).group(1)

                dt = re.search(r'DATE\n(.*?)\n', text).group(1)
                tim = re.search(r'\n(.*?)\nTIME\n', text)
                if tim is None:
                    tim = ""
                else:
                    tim = re.search(r'\n(.*?)\nTIME\n', text).group(1)
                cn = re.search(r'\n(.*?)\nCASE NUMBER\n', text).group(1)
                rsn = re.search(r'TOW INFORMATION\n(.*?)\nCIRCUMSTANCES\n', text).group(1)
                list = [tow_company, address, yr, make, model, clr, tag, (tken1+" "+tken3+" "+tken2), (dt+" "+tim), cn, rsn]
                a_csv(list)
                # print("Completed analyzing file name: " + path)

# creates the towReport.csv file and appends the header fields


with open("towReport.csv", 'w') as file:
    header = ["Tow Company", "Address", "Year", "Make", "Model", "Color", "Tag", "Taken From", "Date/Time", "Police C\\N", "Reason"]
    dw = csv.DictWriter(file, delimiter=',', fieldnames=header)
    dw.writeheader()
    file.close()

# walks through the files in the current directory and passes the file to the analyze function

for file in os.listdir():
    if file.endswith(".pdf"):
        print("Found file: " + file + "\n")
        analyze(file)

print("Program finished running.")
