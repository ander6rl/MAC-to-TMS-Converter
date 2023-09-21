import math
import numpy as np
import pandas as pd
import time


filename = input("Type name of file: ")
filename = filename + ".xlsx"
print(f"File to convert: {filename}\n")

sheetname = input("Name of sheet: ")
print(f"Your new file will be named {sheetname}.xlsx")
# print()

st = time.time()
xls = pd.ExcelFile(filename)
old_df = pd.read_excel(xls, sheet_name=sheetname, nrows=2000)

nd = old_df.to_numpy()




Column_Index_Dictionary_old = dict(zip(
        list(range(0,len(old_df.columns))), old_df.columns))
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
    "Artist (if known)",
    "Object Name",
    "Date",
    "Culture",  # divided into culture and geography, needs to be done by hand, continent country city
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


# 59 seperate places
new_order = [
    "Object Number",
    "Department",
    "Classification",
    "Accession Method",
    "Accession Date",
    "Object Status",
    "Constituent1",
    "Title1",
    "Date",
    "Begin ISO Date",
    "End ISO Date",
    "Date Date Remarks",
    "Medium",
    "Dimensions",
    "Description",
    "Credit Line",
    "Catalogue Raisonne",
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
    "Valuation Purpose1",
    "Stated Value1",
    "Currency1",
    "Stated Date1",
    "Attributes6",
    "Attributes7",
    "Attributes8",
    "Geography Type",  # georgarphy type as place made (needs to be rechecked), they are different categories
    # DEFAULT IS PLACE MADE, CAN ADD OTHER CATEGORIES
    # Place made (note: we will be consolidating Place and Place made)
    # Place manufactured
    # Place printed
    # Possible place
    # Probable place
    # Place photographed (? We were considering adding this)
    "Continent",
    "Country",
    "State/Province",
    "County/Subdivision",
    "City",
    "Township",
    "River",
    "Locale",
    "Locus",
]
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
    "Materials": "Medium",
    "Dimensions (H x W x D cm)": "Dimensions",
    "Dimensions (H x W x D in)": "Notes",
    "Object Condition": "Notes",
    "Condition Report?": "Notes",
    "Images?": "Notes",
    "Research Doc?": "Notes",
    "Flat File Folder?": "Notes",
    "Scanned?": "Notes",
    "Valuation": "Stated Value1",
    "Object Needs": "Notes",
    "Object Notes": "Notes",
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
    tms_list[new_order.index("Object Status")] = "Accessioned"
    # loops through each of the elements for an object from the old spreadsheet
    for i in range(0, len(j)):
        # get the actual value in the current column for the object
        column_val = j[i]
        if not isinstance(column_val, float):
            # print(type(column_val), column_val)
            old_col = Column_Index_Dictionary_old.get(i)
            if old_col == "Object Notes":
                if 'gift' in column_val.lower():
                    tms_list[new_order.index("Accession Method")] = "Gift"
            elif old_col == "Accession Number":
                new_col = convert.get(old_col)
                tms_list[new_order.index(new_col)] = column_val
                splitdate = column_val.split(".")
                if splitdate[0][:1] != 20:
                    tms_list[new_order.index("Accession Date")] = int("19" + splitdate[0])
                else:
                    tms_list[new_order.index("Accession Date")] = int(splitdate[0])
            elif old_col == "Date":
                tms_list[new_order.index("Date")] = column_val
                column_val = str(column_val)
                if '-' in column_val:
                    # HANDLE OR CASES
                    splitdate = column_val.split("-")
                    tms_list[new_order.index("Begin ISO Date")] = splitdate[0]
                    tms_list[new_order.index("End ISO Date")] = splitdate[1]
            else:
                new_col = convert.get(old_col)
                if new_col == "Notes":
                    # print(type(notes_str), type(old_col), type(column_val))
                    notes_str = notes_str + old_col + ": "+ str(column_val) + "\n"
                    tms_list[new_order.index(new_col)] = notes_str
                elif new_col == None:
                    break
                else:
                    tms_list[new_order.index(new_col)] = column_val

    tms_nd = np.append(tms_nd, np.array([tms_list]), axis=0)




newer_df = pd.DataFrame(tms_nd, columns=new_order)
# del newer_df[newer_df.columns[0]]
sheetname = sheetname + '.xlsx'
newer_df.to_excel(sheetname)
et = time.time()
elapsed_time = et - st
print("Completed!")
print("Execution time:", elapsed_time, "seconds")
