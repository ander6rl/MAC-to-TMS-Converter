import math
import numpy as np
import pandas as pd
import time
import string
from datetime import date
import logging

# Program for transforming xlsx file format to proper format for importing into TMS
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def check_keywords(column_val):
    """
    Checks for keywords in the column value and returns the corresponding accession method.
    """
    if "gift" in column_val.lower():
        return "Gift"
    elif "long-term loan" in column_val.lower():
        return "Long-term loan"
    elif "purchase" in column_val.lower():
        return "Purchase"
    elif "trade" in column_val.lower():
        return "Trade"
    else:
        return "(not assigned)"


def main():
    logger.info("This is the Convert to TMS format tool, for converting the MAC collection excel format to the proper TMS formatting the the Importer Tool.")
    filename = "FINAL_MAC"
    filename = filename + ".xlsx"
    print(f"File to convert: {filename}\n")

    # if (input("Is this correct? (y/n)") == 'n'):

    sheetname = input("Name of sheet: ")
    year = 2024
    if sheetname.isdigit():
        year = int(sheetname)
    else:
        year = input("What is the accession year for this set of objects?")
    # sheetname = "Sheet1"

    # rowstart = int(input("Row to start at: "))
    # rowend = int(input("Row to end at: "))

    # sheetname = sheetname + '_' + str(rowstart)
    # sheetname = sheetname + '_' + str(rowend)

    logger.info(f"Your new file will be named {sheetname}.xlsx")


    st = time.time()
    xls = pd.ExcelFile(filename)
    old_df = pd.read_excel(xls, sheet_name=sheetname)
    nd = old_df.to_numpy()

    Column_Index_Dictionary_old = dict(
        zip(list(range(0, len(old_df.columns))), old_df.columns)
    )
    # print(Column_Index_Dictionary_old)
    # print()
    # Column_Index_Dictionary_new = dict(zip(new_df.columns,
    #         list(range(0,len(new_df.columns)))))
    # print(Column_Index_Dictionary_new)

    # print("dict shape: ", Column_Index_Dictionary_new.shape)

    # make key value pairs for reasonable equivalents from old to new database
    # 21 seperate places
    old_order = [
        "Accession Number",
        "On loan / not at MAC",
        "2020-21 Inventory",
        "Row",
        "Shelf/Cabinet",
        "Box/Bay/Drawer",
        "Classification",  # new thing!
        "Artist (if known)",
        "Object Name",
        "Date",
        "Geography Type (Place Made, Place Used, Place Depicted, Place Found). Default is 'Place Made'",  # was Culture # divided into culture and geography, needs to be done by hand, continent country city
        "Geography (Continent)",
        "Geography (Country)",
        "Geography (state/region/island)",
        "Geography (City)",
        "Other Geography (to be added to TMS manually)",
        "Culture",
        "Materials",
        "Dimensions (H x W x D cm)",
        "Dimensions (H x W x D in)",
        "Object Condition",
        "Condition Report?",
        "Images?",
        "Research Doc?",
        "Flat File Folder?",
        "Scanned?",
        "Valuation",
        "Object Needs", # not getting added to notes
        "Object Notes", #not getting added to notes?
    ]

    """
    ObjectID	Object Number	Sort Number	Department	Classification	Accession Method	Accession Date	Object Status	Constituent1	Constituent2	Constituent3	Constituent4	Constituent5	Title1	Title2	Object Name	Date	Begin ISO Date	End ISO Date	Date  Date Remarks	Medium	Dimensions	Description	Credit Line	Catalogue Raisonné	Portfolio/Series	Paper/Support	Signed	Mark(s)	Inscription(s)	Alternate Number1	Alternate Number2	Culture	Period	Object Type	Notes	Label Text	Curatorial Remarks	Provenance	Bibliography	State/Proof	Exhibition History	Published References	Copyright	Curator Approved	Public Access	On View	Accountability	Virtual Object	Location	Location Date	Accession Value	Currency	Stated Date	Exchange Rate	Exchange Rate Date	Valuation Purpose2	Stated Value2	Currency2	Stated Date2	Exchange Rate2	Exchange Rate Date2	Text Entry1	Text Entry2	Text Entry3	Text Entry4	Text Entry5	Text Entry6	Text Entry7	Text Entry8	Text Entry9	Text Entry10	Field Value1	Value Date1	Field Value2	Value Date2	Field Value3	Value Date3	Field Value4	Value Date4	Field Value5	Value Date5	Field Value6	Value Date6	Field Value7	Value Date7	Field Value8	Value Date8	Field Value9	Value Date9	Field Value10	Value Date10	Attributes1	Attributes2	Attributes3	Attributes4	Attributes5	Attributes6	Attributes7	Attributes8	Geography Type	Continent	Country	State/Province	County/Subdivision	City	Township	River	Locale	Locus	Verbatim Latitude	Verbatim Longitude	Elevation	Site	Event	Loan	Lender Object Number

    """

    new_order = [
        "ObjectID",
        "Object Number",
        "Sort Number",
        "Department",
        "Classification",
        "Accession Method",
        "Accession Date",
        "Object Status",
        "Constituent1",
        "Constituent2",
        "Constituent3",
        "Constituent4",
        "Constituent5",
        "Title1",
        "Title2",
        "Object Name",
        "Date",
        "Begin ISO Date",
        "End ISO Date",
        "Date  Date Remarks",
        "Medium",
        "Dimensions",
        "Description",
        "Credit Line",
        "Catalogue Raisonné",
        "Portfolio/Series",
        "Paper/Support",
        "Signed",
        "Mark(s)",
        "Inscription(s)",
        "Alternate Number1",
        "Alternate Number2",
        "Culture",
        "Period",
        "Object Type",
        "Notes",
        "Label Text",
        "Curatorial Remarks",
        "Provenance",
        "Bibliography",
        "State/Proof",
        "Exhibition History",
        "Published References",
        "Copyright",
        "Curator Approved",
        "Public Access",
        "On View",
        "Accountability",
        "Virtual Object",
        "Location",
        "Location Date",
        "Accession Value",
        "Currency",
        "Stated Date",
        "Exchange Rate",
        "Exchange Rate Date",
        "Valuation Purpose2",
        "Stated Value2",
        "Currency2",
        "Stated Date2",
        "Exchange Rate2",
        "Exchange Rate Date2",
        "Text Entry1",
        "Text Entry2",
        "Text Entry3",
        "Text Entry4",
        "Text Entry5",
        "Text Entry6",
        "Text Entry7",
        "Text Entry8",
        "Text Entry9",
        "Text Entry10",
        "Field Value1",
        "Value Date1",
        "Field Value2",
        "Value Date2",
        "Field Value3",
        "Value Date3",
        "Field Value4",
        "Value Date4",
        "Field Value5",
        "Value Date5",
        "Field Value6",
        "Value Date6",
        "Field Value7",
        "Value Date7",
        "Field Value8",
        "Value Date8",
        "Field Value9",
        "Value Date9",
        "Field Value10",
        "Value Date10",
        "Attributes1",
        "Attributes2",
        "Attributes3",
        "Attributes4",
        "Attributes5",
        "Attributes6",
        "Attributes7",
        "Attributes8",
        "Geography Type",
        "Continent",
        "Country",
        "State/Province",
        "County/Subdivision",
        "City",
        "Township",
        "River",
        "Locale",
        "Locus",
        "Verbatim Latitude",
        "Verbatim Longitude",
        "Elevation",
        "Site",
        "Event",
        "Loan",
        "Lender Object Number",
    ]

    # new_order = [
    #     "Object Number",
    #     "Department",
    #     "Classification",
    #     "Accession Method",
    #     "Accession Date",
    #     "Object Status",
    #     "Constituent1",
    #     "Title1",
    #     "Date",
    #     "Begin ISO Date",
    #     "End ISO Date",
    #     "Date Date Remarks",
    #     "Medium",
    #     "Dimensions",
    #     "Description",
    #     "Credit Line",
    #     "Catalogue Raisonne",
    #     "Portfolio/Series",
    #     "Paper/Support",
    #     "Signed",
    #     "Mark(s)",
    #     "Inscription(s)",
    #     "Alternate Number1",
    #     "Alternate Number2",
    #     "Culture",
    #     "Period",
    #     "Object Type",
    #     "Notes",
    #     "Label Text",
    #     "Curatorial Remarks",
    #     "Provenance",
    #     "Bibliography",
    #     "State/Proof",
    #     "Exhibition History",
    #     "Published References",
    #     "Copyright",
    #     "Curator Approved",
    #     "Public Access",
    #     "On View",
    #     "Accountability",
    #     "Virtual Object",
    #     "Location",
    #     "Location Date",
    #     "Valuation Purpose1",
    #     "Stated Value1",
    #     "Currency1",
    #     "Stated Date1",
    #     "Attributes6",
    #     "Attributes7",
    #     "Attributes8",
    #     "Geography Type",  # georgarphy type as place made (needs to be rechecked), they are different categories
    #     # DEFAULT IS PLACE MADE, CAN ADD OTHER CATEGORIES
    #     # Place made (note: we will be consolidating Place and Place made)
    #     # Place manufactured
    #     # Place printed
    #     # Possible place
    #     # Probable place
    #     # Place photographed (? We were considering adding this)
    #     "Continent",
    #     "Country",
    #     "State/Province",
    #     "County/Subdivision",
    #     "City",
    #     "Township",
    #     "River",
    #     "Locale",
    #     "Locus",
    # ]
    tms_nd = np.empty((0, len(new_order)))
    # new_df = pd.DataFrame(columns=new_order)

    # if notes, then it does not fit perfectly into a column
    convert = {
        "Accession Number": "Object Number",
        "On loan / not at MAC": "Notes",
        "2020-21 Inventory": "Notes",
        "Row": "Notes",
        "Shelf/Cabinet": "Notes",
        "Box/Bay/Drawer": "Notes",
        "Artist (if known)": "Constituent1",
        "Object Name": "Title1",
        "Date": "Date",
        # "Culture": "Culture", # maybe have it be blank for now, only going to use for indeginous groups
        "Culture": "Culture",
        "Culture (Tribal Groups)": "Culture",
        "Materials": "Medium",
        "Dimensions (H x W x D cm)": "Dimensions",
        "Dimensions (H x W x D in)": "Notes",
        "Object Condition": "Notes",
        "Condition Report?": "Notes",
        "Images?": "Notes",
        "Research Doc?": "Notes",
        "Flat File Folder?": "Notes",
        "Scanned?": "Notes",
        "Valuation": "Accession Value",
        "Object Needs": "Notes",
        "Object Notes": "Notes",
        "2020-22 Inventory": "Notes",
        "2020-2022 Inventory": "Notes",
        "Geography Type": "Geography Type",
        "Geography Type (Place Made, Place Used, Place Depicted, Place Found). Default is 'Place Made'": "Geography Type",
        "Classification": "Classification",
        "Geography (Continent)": "Continent",
        "Geography (Country)": "Country",
        "Geography (state/region/island)": "State/Province",
        "Geography (State/Region/Island)": "State/Province",
        "Geography (City)": "City",
        "Other Geography (to be added to TMS manually)": "Notes",
        "On loan / not at MAC": "Notes",
        "On Loan/Not at MAC": "Notes",
        "Materials": "Medium",
        "Credit Line": "Credit Line",
    }

    # get list of locations
    # LOCATION IS A BIG PROBLEM

    # loops over the whole numpy2d list, for each element in the old array from the old excel sheet

    total_rows = len(old_df)

    for j in nd:
        # logger.info(f"Processing row {j} of {total_rows}")
        notes_str = ""
        location_str = ""
        tms_list = [None] * len(new_order)
        tms_list[new_order.index("Department")] = "Madison Art Collection"
        tms_list[new_order.index("Object Status")] = "Accessioned Object"
        tms_list[new_order.index("Location Date")] = date.today()

        # is_gift = False

        for i in range(0, len(j)):
            column_val = j[i]
            column_val = str(column_val)  # Convert column_val to string
            if pd.notnull(column_val):  # Check if the column value is not null
                old_col = Column_Index_Dictionary_old.get(i)
                new_col = convert.get(old_col)
                if new_col is None:  # If new column not found, skip
                    continue
                elif new_col == "Classification":
                    if column_val.strip() == "" or column_val.lower() == "nan":
                        tms_list[new_order.index(new_col)] = "(not assigned)"
                    elif column_val.lower() == "print":
                        tms_list[new_order.index(new_col)] = "Prints"
                    elif column_val.lower() == "textiles":
                        tms_list[new_order.index(new_col)] = "Textile"
                    else:
                        tms_list[new_order.index(new_col)] = column_val
                elif new_col == "Notes":
                    accession_method = check_keywords(column_val)
                    if accession_method:
                        tms_list[new_order.index("Accession Method")] = accession_method

                    # Append the notes to notes_str
                    notes_str += f"{old_col}: {column_val}\n"
                    if old_col.lower() == "row" or old_col.lower() == "shelf/cabinet" or old_col.lower() == "box/bay/drawer":
                        location_str += f"{old_col}: {column_val}\n"

                # ...
                elif new_col == "Date":
                    if column_val.lower() == "nan":
                        # If the value is 'nan', leave it empty
                        tms_list[new_order.index("Date")] = ""
                    else:
                        # Otherwise, process the date as usual
                        tms_list[new_order.index("Date")] = column_val
                        firstdate = ""
                        enddate = ""
                        splitor = None
                        if "or" in column_val:
                            splitor = column_val.split("or")
                        # ...
                        # ...
                        if "-" in column_val:
                            if splitor:
                                splitdate = splitor[0].split("-")
                                firstdate = splitdate[0].strip().lstrip("ca.")  # Strip "ca." from the start
                                splitdate = splitor[1].split("-")
                                if len(splitdate) == 2:
                                    enddate = splitdate[1].strip()
                                else:
                                    enddate = splitdate[0].strip()
                                # Check if "BC" or "BCE" is present in either date
                                if "BC" in firstdate or "BCE" in firstdate:
                                    firstdate = "-" + firstdate.replace("BC", "").replace("BCE", "").strip()  # Convert first date to negative number and strip "BC" or "BCE"
                                if "BC" in enddate or "BCE" in enddate:
                                    enddate = "-" + enddate.replace("BC", "").replace("BCE", "").strip()  # Convert end date to negative number and strip "BC" or "BCE"
                            else:
                                splitdate = column_val.split("-")
                                firstdate = splitdate[0].strip().lstrip("ca.")  # Strip "ca." from the start
                                enddate = splitdate[1].strip()
                                # Check if "BC" or "BCE" is present in either date
                                if "BC" in firstdate or "BCE" in firstdate:
                                    firstdate = "-" + firstdate.replace("BC", "").replace("BCE", "").strip()  # Convert first date to negative number and strip "BC" or "BCE"
                                if "BC" in enddate or "BCE" in enddate:
                                    enddate = "-" + enddate.replace("BC", "").replace("BCE", "").strip()  # Convert end date to negative number and strip "BC" or "BCE"
                            tms_list[new_order.index("Begin ISO Date")] = firstdate
                            tms_list[new_order.index("End ISO Date")] = enddate
# ...

# ...

                # ...


                else:
                    if column_val.lower() == "nan":  # Check if column_val is "nan"
                        column_val = ""  # Replace "nan" with empty string
                    if new_col == "Constituent1" or new_col == "Culture":
                        tms_list[new_order.index(new_col)] = string.capwords(column_val, sep=None)
                    else:
                        tms_list[new_order.index(new_col)] = column_val

                    if "gift" in column_val.lower():
                        is_gift = True

        # if is_gift:
        #     tms_list[new_order.index("Accession Method")] = "Gift"
        # else:
        #     tms_list[new_order.index("Accession Method")] = "(not assigned)"
        tms_list[new_order.index("Location")] = "Festival, Room 1000"
        tms_list[new_order.index("Notes")] = notes_str
        tms_nd = np.append(tms_nd, np.array([tms_list]), axis=0)

    # for j in nd:
    #     # print(j)
    #     notes_str = ""
    #     location_str = ""
    #     # makes a new tms_list for the object with enough spaces for each of the columns in the tms excel sheet
    #     tms_list = [None] * len(new_order)
    #     # print(len(tms_list))
    #     tms_list[new_order.index("Department")] = "Madison Art Collection"
    #     tms_list[new_order.index("Object Status")] = "Accessioned Object"
    #     # current_date = datetime.date.today()
    #     # tms_list[new_order.index("Location Date")] = current_date.strftime.today.strftime(
    #     #     "%m-%d-%Y"
    #     # )
    #     tms_list[new_order.index("Location Date")] = date.today()
    #     # loops through each of the elements for an object from the old spreadsheet
    #     for i in range(0, len(j)):
    #         # get the actual value in the current column for the object
    #         column_val = j[i]
    #         column_val = str(column_val)
    #         if not isinstance(column_val, float):
    #             # if True:
    #             # print(type(column_val), column_val)
    #             old_col = Column_Index_Dictionary_old.get(i)
    #             if old_col == "Object Notes":
    #                 notes_str = notes_str + old_col + ": " + str(column_val) + "\n"
    #                 # print(notes_str)
    #                 if "gift" in column_val.lower():
    #                     tms_list[new_order.index("Accession Method")] = "Gift"
    #             elif old_col == "Accession Number":
    #                 new_col = convert.get(old_col)
    #                 tms_list[new_order.index(new_col)] = column_val
    #                 splitdate = column_val.split(".")

    #                 tms_list[new_order.index("Accession Date")] = splitdate[0]

    #             elif old_col == "Date":
    #                 tms_list[new_order.index("Date")] = column_val
    #                 column_val = str(column_val)
    #                 firstdate = ""
    #                 enddate = ""
    #                 splitor = None
    #                 if "or" in column_val:
    #                     splitor = column_val.split("or")
    #                 # print(splitor)
    #                 if "-" in column_val:
    #                     # HANDLE OR CASES
    #                     if splitor != None:
    #                         splitdate = splitor[0].split("-")
    #                         firstdate = splitdate[0]

    #                         splitdate = splitor[1].split("-")
    #                         if len(splitdate) == 2:
    #                             enddate = splitdate[1]
    #                         else:
    #                             enddate = splitdate[0]

    #                     else:
    #                         splitdate = column_val.split("-")
    #                         firstdate = splitdate[0]
    #                         enddate = splitdate[1]

    #                         # if BCE is in enddate, add to firstdate
    #                         if "BCE" in enddate:
    #                             firstdate += " BCE"
    #                         elif "CE" in enddate or "CE" in firstdate:
    #                             firstdate = firstdate.replace("CE", "")
    #                             enddate = enddate.replace("CE", "")

    #                     # check to see if CE is in (not BCE) and remove CE
    #                     # if "CE" in firstdate and "BCE" not in firstdate:

    #                     tms_list[new_order.index("Begin ISO Date")] = firstdate
    #                     tms_list[new_order.index("End ISO Date")] = enddate
    #             else:
    #                 new_col = convert.get(old_col)
    #                 print(old_col)
    #                 if new_col == "Notes":
    #                     # print(type(notes_str), type(old_col), type(column_val))
    #                     notes_str = notes_str + old_col + ": " + str(column_val) + "\n"

    #                     if (
    #                         old_col == "Row"
    #                         or old_col == "Shelf/Cabinet"
    #                         or old_col == "Box/Bay/Drawer"
    #                     ):
    #                         location_str = (
    #                             location_str + old_col + ": " + str(column_val) + "\n"
    #                         )
    #                         # changed to default location
    #                         tms_list[new_order.index("Location")] = (
    #                             "Festival, Room 1000"
    #                         )

    #                     tms_list[new_order.index(new_col)] = str(column_val)
    #                     # print(notes_str)
    #                     # tms_list[new_order.index(new_col)] = notes_str.capitalize()
    #                 if new_col == "Constituent1" or new_col == "Culture":
    #                     tms_list[new_order.index(new_col)] = string.capwords(
    #                         column_val, sep=None
    #                     )
    #                 elif new_col == None:
    #                     break
    #                 else:
    #                     if isinstance(column_val, str):
    #                         tms_list[new_order.index(new_col)] = column_val
    #                     else:
    #                         tms_list[new_order.index(new_col)] = column_val
    #     tms_list[new_order.index("Location")] = location_str
    #     tms_list[new_order.index("Notes")] = notes_str
    #     # print(notes_str)
    #     tms_nd = np.append(tms_nd, np.array([tms_list]), axis=0)

    newer_df = pd.DataFrame(tms_nd, columns=new_order)
    # del newer_df[newer_df.columns[0]]
    sheetname = sheetname + ".xlsx"
    newer_df.to_excel(sheetname)
    et = time.time()
    elapsed_time = et - st
    print("Completed!")
    print("Execution time:", elapsed_time, "seconds")


if __name__ == "__main__":
    main()
