#%%
import pandas as pd
import os
os.chdir('P:/Pertinent Files/Python/scripts/')

#%%
fx  = pd.read_excel("Consolidated_tb.xlsx")
fx=fx[["Date","Promo"]].drop_duplicates()
fx.rename(columns={0:"Date"}, inplace=True)
fx['mn']=fx["Date"].astype(str).str.slice(start=5, stop=7)

mth={'01':"Jan",'02':"Feb",'03':"Mar",'04':"Apr",'05':"May",'06':"Jun",'07':"Jul",'08':"Aug",
     '09':"Sep",'10':"Oct",'11':"Nov",'12':"Dec"
     }
fx['mm']=fx["mn"].map(mth)
fx['ay']=fx.apply(lambda x: x.mm in x.Promo, axis=1)
print(fx[fx.Date.astype(str).str.slice(start=8, stop=10)=="01"])
# %%
