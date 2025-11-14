import pandas as pd

xy=pd.read_excel('P:/Downloads/AR_Report.xlsx', sheet_name='By Sales Channel - Ky')
print(xy.columns)

yy=xy.drop(columns='OrganizationName', axis=1)
print(yy.columns)