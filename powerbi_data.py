# CSV -> Excel -> Power BI

import pandas as pd

CATEGORIES = ["security", "reliability", "performance-efficiency"]

logs = ""
csvPath = "test.csv"

try:
    df = pd.read_csv(csvPath, 
                     dtype=object, 
                     usecols=["Cloud Provider", 
                              "Check Status", 
                              "Categories",
                              ])
    df.index.name = "index"
    df.dropna(inplace=True)
except:
    print("Can't read csv file.")
    logs = logs + "Can't read csv file.\n"
    pass

providers = df["Cloud Provider"].unique().tolist()
writer = pd.ExcelWriter("conformityyy.xlsx")

for provider in providers:
    prv = df[df["Cloud Provider"] == provider]

    for category in CATEGORIES:
        sheet = provider + "." + category
        tmp = prv[(prv['Categories'].str.contains(category))]
        tmp.to_excel(writer, sheet_name=sheet, columns=["Check Status"])

writer.close()    
