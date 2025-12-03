#%%
import sys
file_path_r = 'C:/Users/Administrator/Documents/Python_Automations/'
download_path = 'C:/Users/Administrator/Downloads/'



# file_path_r = 'D:/YU/ScriptCodez/'
# download_path = 'C:/Users/Pole Okuttu/Downloads/'

download_file_address = f"{download_path}data.xlsx"
sys.path.append(file_path_r)
from powerbi_sign_in_file import *
today = datetime.datetime.today()
first_day = today.replace(day=1)
first_day=first_day.strftime('%m/%d/%Y')

#%%
active_warehouse = [
    {'WarehouseId' : 'Bulaga', 'Email':'bulagaduuka@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Bunamwaya', 'Email':'bunamwayastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Busia', 'Email':'busiastore@yalelo.ug,  pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Bwaise', 'Email':'bwaisestore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Elegu', 'Email':'eleguborder@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Gulu', 'Email':'gulustore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Jinja V3', 'Email':'jinjastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kafunta', 'Email':'kafuntastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kasangati', 'Email':'kasangatistore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kajjansi', 'Email':'kajjansistore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kasubi', 'Email':'kasubistore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kawempe V2', 'Email':'kawempestore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kibuli', 'Email':'kibulistore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kibuye', 'Email':'kibuyestore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kireka', 'Email':'kirekastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kisoro', 'Email':'kisorostore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kyaliwajal', 'Email':'kyaliwajjalastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kyambogo', 'Email':'kyambogostore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Kyengera-R', 'Email':'kyengerastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Distrbtion', 'Email':'kyengerastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Malaba', 'Email':'malababorder@yalelo.ug,  pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Mpondwe', 'Email':'mpondweborder@yalelo.ug,  pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Mukono', 'Email':'mukonoshop@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Mutungo', 'Email':'mutungostore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Nansana V2', 'Email':'nansanastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Production', 'Email':'nbitsinze@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Natete V2', 'Email':'nateetestore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Ntinda', 'Email':'ntindastore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'KyebandoR', 'Email':'mbiryetega@yalelo.ug, jnanyonjo@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'HighValue', 'Email':'mbiryetega@yalelo.ug, jnanyonjo@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Nyahuka', 'Email':'nyahukaborder@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'Odramacaku', 'Email':'odramacakustore@yalelo.ug, pokuttu@yalelo.ug'},
    {'WarehouseId' : 'KyebandoDC', 'Email':'mbiryetega@yalelo.ug, jnanyonjo@yalelo.ug, pokuttu@yalelo.ug'}
    ]

active_warehouse = pd.DataFrame(active_warehouse)

#%%
gr_transfers_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/e4e07d5f73507b70b7bd"
actual_received_stock = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/a25014801c32e2406317"
overall_sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/069182dba7e1167bd78a"
banking_ = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/91687127-028f-4c1e-8ad9-3f783f724150/374266a9250a727b593e"

file_path=[]
    # load url
#%%
# read files
tr_in = pbi_export(gr_transfers_url,download_file_address)
df_in = pd.pivot_table(tr_in,index=["Date", "ReceivingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
df_in['Total'] = df_in.sum(axis=1, skipna=True)

#%%
#DC transfers from production
df_in_dc_pdn = pd.pivot_table(tr_in[(tr_in['Received (kg)'] > 15) 
                                    & ((tr_in['ShippingWarehouseId'] == 'Production')|(tr_in['ShippingWarehouseId'] == 'HighValue')) 
                                    & (tr_in['ReceivingWarehouseId'] == 'KyebandoDC')],
                              index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], 
                              columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])

df_in_dc_pdn['Total Received'] = df_in_dc_pdn.sum(axis=1, skipna=True)

#%%
#DC transfers from other warehouse
df_in_dc_wh = pd.pivot_table(tr_in[(tr_in['Received (kg)'] > 15) 
                                   & (((tr_in['ShippingWarehouseId'] != 'Production') & (tr_in['ShippingWarehouseId'] != 'HighValue')) ) 
                                   & (tr_in['ReceivingWarehouseId'] == 'KyebandoDC')],
                                    index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], 
                                    columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
df_in_dc_wh['Total Received Frm WHS'] = df_in_dc_wh.sum(axis=1, skipna=True)
#%%
#Driploss transfers from other warehouse
df_dl_dc = pd.pivot_table(tr_in[(tr_in['Received (kg)'] < 15) 
                                & ((tr_in['ShippingWarehouseId'] != 'Production')|(tr_in['ShippingWarehouseId'] != 'HighValue')) 
                                & (tr_in['ReceivingWarehouseId'] == 'KyebandoDC')],
                                index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], 
                                columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')
df_dl_dc['Driploss From Shops'] = df_dl_dc.sum(axis=1, skipna=True)
#%%
# tr_out = pbi_export(gr_transfers_url,download_file_address)
df_out = pd.pivot_table(tr_in,index=["Date", "ShippingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
df_out['Total'] = df_out.sum(axis=1, skipna=True)
#%%
net_received_stock = pbi_export(actual_received_stock,download_file_address)
net_received_stock = pd.pivot_table(net_received_stock,index=["Date","WarehouseId"], columns=["SKU","SizeId22"], values="Received Stock", aggfunc='sum')#.reset_index()
net_received_stock['Total'] = net_received_stock.sum(axis=1, skipna=True)

#%%
sold_stock = pbi_export(overall_sales,download_file_address)
sold_stock = sold_stock[sold_stock['Date'] >= first_day]
sold_stock = pd.pivot_table(sold_stock,index=["Date","WarehouseId"], columns=["ItemNumber","ProductSizeId"], values="Sales", aggfunc='sum')#.reset_index()
sold_stock['Total'] = sold_stock.sum(axis=1, skipna=True)
#%%
# Banking
banked = pbi_export(banking_,download_file_address)
# banked["Date"] = pd.to_datetime(banked['Date'])
banked = banked[banked['Date'] >= first_day]
browser.quit()
#%%

#%%
# log-in to email server

"""
# ######################################################################
# # Email With Attachments Python Script
# ######################################################################
# """

from gmail_sender import gmail_function
print("Succesfully connected to server")
#%%
for i,e in zip(active_warehouse['WarehouseId'].to_list(), active_warehouse['Email'].to_list()):
    email_list = [email.strip() for email in e.split(',')] 
    # cc_list =  ["pokuttu@yalelo.ug","rnabukeera@yalelo.ug"]
    try:
        act_received  = net_received_stock.loc[(slice(None), i),:].dropna(axis=1, how='all')
        act_received.loc['Sub Total', :] = act_received.sum().values
        
    except:
        None
    try:
        sales_stock  = sold_stock.loc[(slice(None), i),:].dropna(axis=1, how='all')
        sales_stock.loc['Sub Total', :] = sales_stock.sum().values
    except:
        None
    try:
        df_in_w = df_in.loc[(slice(None), i),:].dropna(axis=1, how='all')
        df_in_w.loc['Sub Total', :] = df_in_w.sum().values
    except:
        None
    try:
        df_out_w = df_out.loc[(slice(None), i),:].dropna(axis=1, how='all')
        df_out_w.loc['Sub Total', :] = df_out_w.sum().values
    except:
        None
    try:
        df_in_dc_pdn = df_in_dc_pdn.loc[slice(None), i, slice(None)].dropna(axis=1, how='all')
        df_in_dc_pdn.loc['Sub Total', :] = df_in_dc_pdn.sum().values
    except:
        pass
    try:
        df_in_dc_wh = df_in_dc_wh.loc[slice(None), i, slice(None)].dropna(axis=1, how='all')
        df_in_dc_wh.loc['Sub Total', :] = df_in_dc_wh.sum().values
    except:
        None
    try:
        df_dl_dc = df_dl_dc.loc[slice(None), i, slice(None)].dropna(axis=1, how='all')
        df_dl_dc.loc['Sub Total', :] = df_dl_dc.sum().values
    except:
        None
    try:
        bank_tr  = banked[banked['WarehouseId']== i]#.dropna(axis=1, how='all')
        subtotal = bank_tr.select_dtypes(include='number').sum()
        subtotal['WarehouseId'] = 'Sub Total'
        subtotal['Date'] = ''
        subtotal['Risk Status'] = ''
        # Append subtotal
        bank_tr = pd.concat([bank_tr, pd.DataFrame([subtotal])], ignore_index=True)
        # Format numeric columns with commas
        num_cols = bank_tr.select_dtypes(include='number').columns
        bank_tr[num_cols] = bank_tr[num_cols].applymap(lambda x: f"{x:,.0f}")
        
    except:
        None

    if i=='KyebandoDC':
        with pd.ExcelWriter(f"{i} MTD Stock.xlsx") as writer:
                # Write df_in_w to the first sheet
                try:
                    df_in_dc_pdn.to_excel(writer, sheet_name=f'{i} Received(Frm Pdn)')
                except:
                    None
                # Write df_in_w to the first sheet
                try:
                    df_dl_dc.to_excel(writer, sheet_name=f'{i} Driploss(Frm Shops)')
                except:
                    None
                # Write df_in_w to the first sheet
                try:
                    df_in_dc_wh.to_excel(writer, sheet_name=f'{i} Received(WHS)')
                except:
                    None
                # Write df_out_w to the second sheet
                try:
                   df_out_w.to_excel(writer, sheet_name=f'{i} Stock Shipped Out')
                except:
                    None
    else:
        with pd.ExcelWriter(f"{i} MTD Stock.xlsx") as writer:
            # Write df_in_w to the first sheet
            try:
                act_received.to_excel(writer, sheet_name=f'{i} Net Stock Received')
            except:
                None
            try:
                df_in_w.to_excel(writer, sheet_name=f'{i} Gross Stock Received')
            except:
                None
            # Write df_out_w to the second sheet
            try:
                df_out_w.to_excel(writer, sheet_name=f'{i} Stock Shipped Out')
            except:
                print("Stock Shipped Out")
                None
            try:
                sales_stock.to_excel(writer, sheet_name=f'{i} Sold Stock')
            except:
                None
            try:
                bank_tr.to_excel(writer, sheet_name=f'{i} Banking Records', index=False)
            except:
                None
#name the email subject
    subject = f"{i} Transfers Report [Month To Date]"
    attachment_ = f"{i} MTD Stock.xlsx"
    # Prepare the body of the email
    if i == 'KyebandoDC':
        body = f"""
        Hello {i} Centre,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug for followup.

        Regards,
        Pole
        """
    elif i in {'Malaba','Busia','Nyahuka','Mpondwe'}:
        body = f"""
        Hello {i} Border Point,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug and pomara@yalelo.ug in copy for followup.

        Regards,
        Pole
        """
    else:
        body = f"""
        Hello {i} Store,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug and mbiryetega@yalelo.ug in copy for followup.

        Regards,
        Pole
        """
    gmail_function(e,subject_line=subject, email_body=body, attachment_path=attachment_)
#%%
import glob
import os

for k in glob.glob("C:/Users/Administrator/Downloads" + "*.xlsx"):
    os.remove(k)
try:
    for g in glob.glob("C:/Users/Administrator/Documents/Python_Automations/"+ "*MTD Stock*" + "*.xlsx"):
        os.remove(g)
except:
    None
try:
    for g in glob.glob("C:/Users/Administrator/Documents/Python_Automations/Finance/"+ "*MTD Stock*" + "*.xlsx"):
        os.remove(g)
except:
    None

print("All files removed from repository")
