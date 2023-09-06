import numpy as np
import pandas as pd



df = pd.read_excel('ImporterTester.xlsx', nrows=10)

# make key value pairs for reasonable equivalents from old to new database
convert = {"Accession Number": "Object Number", ""
}

print(df)

df.to_excel("output.xlsx")
