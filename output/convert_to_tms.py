import math
import numpy as np
import pandas as pd
import time
import string
from datetime import date
import logging
import os
import sys

# This program is for converting the MAC collection excel format to the proper TMS formatting for the Importer Tool.
# Author: Rebecca Anderson
# Version: 05/08/2024

# If you have any questions about this code, email me at ander6rl@dukes.jmu.edu

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


def split_and_save_excel_files(data, new_order, accession_date, sheetname):
    num_splits = math.ceil(len(data) / 200)
    for i in range(num_splits):
        start_index = i * 200
        end_index = min((i + 1) * 200, len(data))
        split_data = data[start_index:end_index]
        df = pd.DataFrame(split_data, columns=new_order)
        filename = f"{sheetname}_part_{i + 1}.xlsx"
        while True:
            try:
                df.to_excel(filename, index=False)
                print(f"File '{filename}' written successfully!")
                break
            except PermissionError:
                print(
                    "Permission Error: The file is already open. Please close the file and try again."
                )
                choice = input("Do you want to try again? (yes/no): ").lower()
                if choice != "yes":
                    print("Exiting program.")
                    sys.exit(1)
            except Exception as e:
                print(f"An error occurred: {e}")
                break


def main():
    logger.info(
        "This is the Convert to TMS format tool, for converting the MAC collection excel format to the proper TMS formatting the the Importer Tool."
    )
    filename = "FINAL_MAC"
    filename = filename + ".xlsx"
    print(f"File to convert: {filename}\n")

    # if (input("Is this correct? (y/n)") == 'n'):

    sheetname = input("Name of sheet: ")
    accession_date = input("What is the accession year? ").strip()
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
        "Object Needs",  # not getting added to notes
        "Object Notes",  # not getting added to notes?
    ]

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
                    # print("-----------------")
                    # print("Classification: ", column_val)
                    if column_val.strip() == "" or column_val.lower() == "nan":
                        tms_list[new_order.index(new_col)] = "(not assigned)"
                    elif column_val.lower() == "print":
                        tms_list[new_order.index(new_col)] = "Prints"
                    elif column_val.lower() == "textiles":
                        tms_list[new_order.index(new_col)] = "Textile"
                    # if column val contains arms or armour set to arms and armor
                    elif "armour" in column_val.lower():
                        tms_list[new_order.index(new_col)] = "Arms and Armor"
                    elif "decorative arts" in column_val.lower():
                        tms_list[new_order.index(new_col)] = "Decorative Art"
                    elif "paintings" in column_val.lower():
                        tms_list[new_order.index(new_col)] = "Painting"
                    elif (
                        "religious" in column_val.lower()
                        or "funerary" in column_val.lower()
                    ):
                        tms_list[new_order.index(new_col)] = (
                            "Religious and funerary items"
                        )
                    elif (
                        "jewelry" in column_val.lower()
                        or "personal" in column_val.lower()
                    ):
                        tms_list[new_order.index(new_col)] = (
                            "Jewelry and personal items"
                        )
                    # elif (
                    #     "books"
                    #     or "documentation" in column_val.lower()
                    # ):
                    #     tms_list[new_order.index(new_col)] = (
                    #         "Books, Manuscripts, Documents"
                    #     )
                    else:
                        tms_list[new_order.index(new_col)] = column_val
                    # print("Classification: ", tms_list[new_order.index(new_col)])
                    # print("-----------------")
                elif new_col == "Accession Value" and column_val.lower() != "nan":
                    print("Old Accession Value: ", column_val)
                    if column_val.strip() != "" and "(" not in column_val:
                        tms_list[new_order.index("Currency")] = "US $"
                        tms_list[new_order.index(new_col)] = column_val
                        tms_list[new_order.index("Stated Date")] = date.today()
                    elif "(" in column_val:
                        year = column_val.split("(")[-1].split(")")[0]
                        column_val = column_val.replace(f"({year})", "").strip()
                        tms_list[new_order.index("Stated Date")] = f"{year}-01-01"
                        tms_list[new_order.index("Currency")] = "US $"
                        tms_list[new_order.index(new_col)] = column_val
                    print("New Accession Value: ", tms_list[new_order.index(new_col)])
                elif new_col == "Stated Date":
                    continue
                elif new_col == "Currency":
                    continue
                elif new_col == "Notes":
                    accession_method = check_keywords(column_val)
                    if accession_method:
                        tms_list[new_order.index("Accession Method")] = accession_method

                    # Append the notes to notes_str

                    if column_val.lower() == "nan":
                        column_val = ""

                    notes_str += f"{old_col}: {column_val}\n"
                    if (
                        old_col.lower() == "row"
                        or old_col.lower() == "shelf/cabinet"
                        or old_col.lower() == "box/bay/drawer"
                    ):
                        location_str += f"{old_col}: {column_val}\n"

                # ...
                elif new_col == "Date":
                    if column_val.lower() == "nan":
                        # If the value is 'nan', leave it empty
                        tms_list[new_order.index("Date")] = ""
                    else:
                        # Otherwise, process the date as usual
                        tms_list[new_order.index("Date")] = column_val
                        # print("----column_val:", column_val)  # Debugging
                        firstdate = ""
                        enddate = ""
                        splitor = None
                        if "or" in column_val:
                            splitor = column_val.split("or")
                        # print("splitor:", splitor)  # Debugging
                        print(column_val)
                        if "-" in column_val:
                            if splitor:
                                splitdate = splitor[0].split("-")
                                # print("1splitdate:", splitdate)  # Debugging
                                firstdate = (
                                    splitdate[0].strip().lstrip("ca.")
                                )  # Strip "ca." from the start
                                # print("firstdate:", firstdate)  # Debugging
                                splitdate = splitor[1].split("-")
                                # print("2splitdate:", splitdate)  # Debugging
                                if len(splitdate) == 2:
                                    enddate = splitdate[1].strip()
                                else:
                                    enddate = splitdate[0].strip()
                                # print("enddate:", enddate)  # Debugging
                                # Check if "BC" or "BCE" is present in either date

                                # check end date fist if in end date make both negative
                                # then elif check first date and then only set first date to negative

                                if "BC" in enddate or "BCE" in enddate:
                                    firstdate = "-" + "".join(
                                        filter(str.isdigit, firstdate)
                                    )
                                    enddate = "-" + "".join(
                                        filter(str.isdigit, enddate)
                                    )
                                elif "BC" in firstdate or "BCE" in firstdate:
                                    firstdate = "-" + "".join(
                                        filter(str.isdigit, firstdate)
                                    )
                                    enddate = "".join(filter(str.isdigit, enddate))
                                else:
                                    firstdate = "".join(filter(str.isdigit, firstdate))
                                    enddate = "".join(filter(str.isdigit, enddate))

                            else:
                                splitdate = column_val.split("-")
                                # print("splitdate:", splitdate)  # Debugging
                                firstdate = (
                                    splitdate[0].strip().lstrip("ca.")
                                )  # Strip "ca." from the start
                                # print("firstdate:", firstdate)  # Debugging
                                enddate = splitdate[1].strip()
                                # print("enddate:", enddate)  # Debugging
                                # Check if "BC" or "BCE" is present in either date
                                # check end date fist if in end date make both negative
                                # then elif check first date and then only set first date to negative

                                if "BC" in enddate or "BCE" in enddate:
                                    firstdate = "-" + "".join(
                                        filter(str.isdigit, firstdate)
                                    )
                                    enddate = "-" + "".join(
                                        filter(str.isdigit, enddate)
                                    )
                                elif "BC" in firstdate or "BCE" in firstdate:
                                    firstdate = "-" + "".join(
                                        filter(str.isdigit, firstdate)
                                    )
                                    enddate = "".join(filter(str.isdigit, enddate))
                                else:
                                    firstdate = "".join(filter(str.isdigit, firstdate))
                                    enddate = "".join(filter(str.isdigit, enddate))

                            tms_list[new_order.index("Begin ISO Date")] = firstdate
                            tms_list[new_order.index("End ISO Date")] = enddate
                            # print("----Final firstdate:", firstdate)  # Debugging
                            # print("----Final enddate:", enddate)

                else:
                    if column_val.lower() == "nan":  # Check if column_val is "nan"
                        column_val = ""  # Replace "nan" with empty string

                    accession_method = check_keywords(column_val)
                    if accession_method:
                        tms_list[new_order.index("Accession Method")] = accession_method

                    if new_col == "Constituent1" or new_col == "Culture":
                        tms_list[new_order.index(new_col)] = string.capwords(
                            column_val, sep=None
                        )
                    else:
                        tms_list[new_order.index(new_col)] = column_val

                    if "gift" in column_val.lower():
                        is_gift = True

        tms_list[new_order.index("Accession Date")] = accession_date + "-01-01"
        tms_list[new_order.index("Location")] = "Festival, Room 1000"
        tms_list[new_order.index("Notes")] = notes_str
        tms_nd = np.append(tms_nd, np.array([tms_list]), axis=0)

    newer_df = pd.DataFrame(tms_nd, columns=new_order)
    # del newer_df[newer_df.columns[0]]
    sheetname = sheetname + ".xlsx"

    split_and_save_excel_files(tms_nd, new_order, accession_date, sheetname)

    et = time.time()
    elapsed_time = et - st
    print("Completed!")
    print("Execution time:", elapsed_time, "seconds")


if __name__ == "__main__":
    main()
