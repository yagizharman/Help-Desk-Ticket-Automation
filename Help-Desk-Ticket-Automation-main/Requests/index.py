import os
import xml.etree.ElementTree as ETree
import pandas as pd
#required modules
#pip install pandas
#pip install xlsxwriter

#give dynamic path
xmlpath = "./data.xml"
excelpath = "./Request.xlsx"

#parse xml file
Tree = ETree.parse(xmlpath)
root = Tree.getroot()
A = list()
for element in root:
    B = dict()
    for i in list(element):
        B.update({i.tag: i.text})
        A.append(B)
df = pd.DataFrame(A)

#remove duplicates
df.drop_duplicates(keep="first", inplace=True)
df.reset_index(drop=True, inplace=True)

#write parsed data to excel file
writer = pd.ExcelWriter(excelpath, engine="xlsxwriter")
df.to_excel(writer, sheet_name="sheet")
worksheet = writer.sheets["sheet"]
worksheet.set_column("B:D", 30) #set 30 chars width for columns [B:D]
writer.save()




