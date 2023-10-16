import math
import numpy as np
import pandas as pd
import time
import string


filename = "Copy_of_MAC_Collection"
if(input(f"Do you want a file name that is not {filename}? (y/yes): ")):
    filename = input("Enter new filename: ")

filename = filename + ".xlsx"
print(f"File to convert: {filename}")

pagename = input("Name of excel page to convert: ")
# print(f"Your new file will be named {pagename}.xlsx")
destname = input("Destination file name: ")
print(f"Your new file will be named {destname}.xlsx")
# print()

startrow = int(input("Enter the starting row: "))
endrow = int(input("Enter the ending row: "))

st = time.time()
xls = pd.ExcelFile(filename)
# old_df = pd.read_excel(xls, sheet_name=pagename, nrows=20)
old_df = pd.read_excel(xls, sheet_name=pagename, skiprows=range(1, startrow), nrows=endrow - startrow + 1)


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
    "Classification", # new thing!
    "Artist (if known)",
    "Object Name",
    "Date",
    "Geography Type (Place Made, Place Used, Place Depicted, Place Found). Default is 'Place Made'", # was Culture # divided into culture and geography, needs to be done by hand, continent country city
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
    "Object Needs",
    "Object Notes",
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
    "Geography Type (Place Made, Place Used, Place Depicted, Place Found). Default is 'Place Made'": "Geography Type",
    "Classification": "Classification",
    "Geography (Continent)": "Continent",
    "Geography (Country)": "Country",
    "Geography (state/region/island)": "State/Province",
    "Geography (City)": "City",
    "Other Geography (to be added to TMS manually)":"Notes",
    "On loan / not at MAC": "Notes",
}

# get list of locations
# LOCATION IS A BIG PROBLEM

# loops over the whole numpy2d list, for each element in the old array from the old excel sheet
for j in nd:
    notes_str = ""
    # makes a new tms_list for the object with enough spaces for each of the columns in the tms excel sheet
    tms_list = [None] * len(new_order)
    # print(len(tms_list))
    tms_list[new_order.index("Department")] = "Madison Art Collection"
    tms_list[new_order.index("Object Status")] = "Accessioned Object"
    # loops through each of the elements for an object from the old spreadsheet
    for i in range(0, len(j)):
        # get the actual value in the current column for the object
        column_val = j[i]
        if not isinstance(column_val, float):
            # print(type(column_val), column_val)
            old_col = Column_Index_Dictionary_old.get(i)
            if old_col == "Object Notes":
                if "gift" in column_val.lower():
                    tms_list[new_order.index("Accession Method")] = "Gift"
            elif old_col == "Accession Number":
                new_col = convert.get(old_col)
                tms_list[new_order.index(new_col)] = column_val
                splitdate = column_val.split(".")
                tms_list[new_order.index("Accession Date")] = int(pagename)
            elif old_col == "Date":
                tms_list[new_order.index("Date")] = column_val
                column_val = str(column_val)
                firstdate = ""
                enddate = ""
                splitor = None
                if "or" in column_val:
                    splitor = column_val.split("or")
                # print(splitor)
                if "-" in column_val:
                    # HANDLE OR CASES
                    if splitor != None:
                        splitdate = splitor[0].split("-")
                        firstdate = splitdate[0]

                        splitdate = splitor[1].split("-")
                        if len(splitdate) == 2:
                            enddate = splitdate[1]
                        else:
                            enddate = splitdate[0]

                    else:
                        splitdate = column_val.split("-")
                        firstdate = splitdate[0]
                        enddate = splitdate[1]
                    tms_list[new_order.index("Begin ISO Date")] = firstdate
                    tms_list[new_order.index("End ISO Date")] = enddate
            else:
                new_col = convert.get(old_col)
                if new_col == "Notes":
                    # print(type(notes_str), type(old_col), type(column_val))
                    notes_str = notes_str + old_col + ": " + str(column_val) + "\n"
                    # print(notes_str)
                    # tms_list[new_order.index(new_col)] = notes_str.capitalize()
                if new_col == "Constituent1" or new_col == "Culture":
                    tms_list[new_order.index(new_col)] = string.capwords(column_val, sep = None)
                elif new_col == None:
                    break
                else:
                    if isinstance(column_val, str):
                        tms_list[new_order.index(new_col)] = column_val.capitalize()
                    else:
                        tms_list[new_order.index(new_col)] = column_val
    tms_list[new_order.index("Notes")] = notes_str.capitalize()
    tms_nd = np.append(tms_nd, np.array([tms_list]), axis=0)


newer_df = pd.DataFrame(tms_nd, columns=new_order)
# del newer_df[newer_df.columns[0]]
sheetname = destname + ".xlsx"
newer_df.to_excel(sheetname)
et = time.time()
elapsed_time = et - st
print("Completed!")
print("Execution time:", elapsed_time, "seconds")
