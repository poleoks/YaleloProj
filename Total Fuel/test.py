# import yfinance as yf
# import numpy as np
import pandas as pd
# import Jinja2
# import statsmodels.api as sm
data=pd.read_excel("P:\Pertinent Files\Python\scripts\Finance\ky_Report.xlsx")
# data=pd.DataFrame(data)
data.iloc[:,3:]=data.iloc[:,3:].astype('int')
df=pd.DataFrame(data.iloc[:,3:].sum().to_frame().T)
dd=pd.concat([data,df], ignore_index=True)
dd.iloc[-1, 0] = 'Total'  
dd.fillna('',inplace=True)


print(dd.tail())
