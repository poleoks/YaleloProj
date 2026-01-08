save_dir= "C:/Users/Administrator/Documents/Python_Automations/"
import sys
sys.path.append(save_dir)
from credentials import *
from powerbi_sign_in_file import *
from warehouse_email import *
import glob
import time
import os
import pandas as pd
#%%
all_dispatches = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/d4e677323a58b06ae221"
finance_all_dispatches = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/a5e9a4c8f62722575c70"
stock_movement = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/67fa000c96323c50965b"
month_end_stock_movt = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/06b296daf151423218f4"
FG_Sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/4db7a23828ef88aa2214"
net_received = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/a25014801c32e2406317?ctid=683df1d2-484b-4bb5-8ca6-1cdbf2b98d28"
transfers_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/e4e07d5f73507b70b7bd"
productionbal = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/6dab297c5ed36ee40b66"
finance_fuel_recon = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/2f993e0760669650e06d"
all_sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/6ea640eaa18c4ad2ad29"
sales_channel_stock = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/905a173e3afcce906eab"
all_transfers = "http://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/aca5448dcbe1175be2d2"
harvests_aqua = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/7095f1e2-0598-4ec9-ba08-cc4f2cd2f9c6/818e08a4e3e8b2785e2c"
nairobi_received_stock = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/fe0370666195e7dbdbc8"
finance_all_sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/e93c265c-fe76-4358-b056-8f49bc2ed8ad/db31a1f8030ec9954a01"
overall_dispatch = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/48b560bd224dad1d787b"
overall_sale = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/069182dba7e1167bd78a"
otif_qty = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/334b99a8547099358c4a"
otif_tm = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/531890ead8b934643bb5"

#%%
download_address=glob.glob("c:/Users/Administrator/Downloads/data"+"*xlsx") #c:/Users\Administrator\Downloads
download_file_path = "C:/Users/Administrator/Downloads/data.xlsx"
file_path=[]
#%%
for file in download_address:
    try:
        os.remove(file)
        print(f"{file} removed")
    except:
        print("no file found")
# export pbi reports
dispatch = pbi_export(all_dispatches,download_file_path)
# print(dispatch.head(2))

time.sleep(2)
stock_summary = pbi_export(stock_movement,download_file_path)
# print(stock_summary.head(2))

time.sleep(2)

month_end_stock_movt = pbi_export(month_end_stock_movt,download_file_path)


time.sleep(2)
farmgate = pbi_export(FG_Sales,download_file_path)


time.sleep(2)
net_received_stock = pbi_export(net_received,download_file_path)

time.sleep(2)
driploss = pbi_export(transfers_url,download_file_path)

pdnopening =pbi_export(productionbal,download_file_path)
time.sleep(2)

ffuel_recon =pbi_export(finance_fuel_recon,download_file_path)
time.sleep(2)

all_net_sales =pbi_export(all_sales,download_file_path)
time.sleep(2)

sc_stock = pbi_export(sales_channel_stock,download_file_path)
time.sleep(2)

transfers = pbi_export(all_transfers,download_file_path)
time.sleep(2)

harvests_aquam = pbi_export(harvests_aqua,download_file_path)
time.sleep(2)

nai_rstock = pbi_export(nairobi_received_stock,download_file_path)
time.sleep(2)


fin_sales = pbi_export(finance_all_sales,download_file_path)
time.sleep(2)

fin_dispatches = pbi_export(finance_all_dispatches,download_file_path)
time.sleep(2)


overall_dispatches = pbi_export(overall_dispatch,download_file_path)
time.sleep(2)
overall_sales = pbi_export(overall_sale,download_file_path)
time.sleep(2)
otif_q = pbi_export(otif_qty,download_file_path)
time.sleep(2)

otif_tm = pbi_export(otif_tm,download_file_path)
time.sleep(2)
browser.quit()
#%% sys.path.insert(0,"E:/Python_Automations/Pole/Finance/")


with pd.ExcelWriter('stock_flow.xlsx') as wr:
    dispatch.to_excel(wr, sheet_name="dispatch")
    stock_summary.to_excel(wr, sheet_name= "stock_summary")
    month_end_stock_movt.to_excel(wr, sheet_name = "month_end_stock_movt")
    farmgate.to_excel(wr, sheet_name = "farm_gate_sales")
    net_received_stock.to_excel(wr, sheet_name="net_stock_received")
    driploss.to_excel(wr, sheet_name = "drip_loss")
    pdnopening.to_excel(wr, sheet_name = "prod_opening")
    ffuel_recon.to_excel(wr,sheet_name="ffuel_recon")
    all_net_sales.to_excel(wr, sheet_name="all_net_sales")
    sc_stock.to_excel(wr, sheet_name="shipped_sc_stock")
    transfers.to_excel(wr, sheet_name="all_transf")
    harvests_aquam.to_excel(wr, sheet_name="aquamgr_harvests_cogs")
    nai_rstock.to_excel(wr, sheet_name="Received_Nairobi")
    fin_sales.to_excel(wr, sheet_name="fin_sales")
    fin_dispatches.to_excel(wr, sheet_name="fin_dispatches")
    overall_dispatches.to_excel(wr, sheet_name='overall_dispatch')
    overall_sales.to_excel(wr, sheet_name='overall_sale')
    otif_q.to_excel(wr,sheet_name="otif_qty")
    otif_tm.to_excel(wr,sheet_name="otif_time")

#%%
from gmail_sender import gmail_function

gmail_function('pokuttu@yalelo.ug','stock_flow','This is an ETL for all stock origin and destination',
               'stock_flow.xlsx')

try:
    os.remove('stock_flow.xlsx')
    print("stock_flow file deleted")
except:
    print("no file found")
    pass

#%%
kill_browser("chrome")