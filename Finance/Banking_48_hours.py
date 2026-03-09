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
today_ = today.strftime('%m/%d/%Y')
first_day=first_day.strftime('%m/%d/%Y')

banking_ = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/91687127-028f-4c1e-8ad9-3f783f724150/374266a9250a727b593e"

# Banking
banked = pbi_export(banking_,download_file_address)
#%%

df_risk = banked.loc[(banked['Date'] >= first_day) & (banked['Date'] < today_)]
df_risk['Daywithoutbanking'] = (pd.Timestamp.now() - pd.to_datetime(df_risk['Date'])).dt.days


df_risk_aggd = df_risk[['WarehouseId', 'Cash Sales (UGX)', 'Deposits (UGX)', 'Variance (UGX)']].groupby('WarehouseId').agg(
    Cash_Sales_ = ('Cash Sales (UGX)', 'sum'),
    Deposits_ = ('Deposits (UGX)', 'sum'),
    Variance_ = ('Variance (UGX)', 'sum')
    ).reset_index()

#risky mtd
df_risky_mtd = df_risk_aggd[df_risk_aggd['Variance_'] < -100000]
num_cols = df_risky_mtd.select_dtypes(include='number').columns
df_risky_mtd[num_cols] = df_risky_mtd[num_cols].applymap(lambda x: f"{x:,.0f}")
print(df_risky_mtd.head(10))

#risky days
df_high_risk = df_risk[(df_risk['Daywithoutbanking'] >= 1) & (df_risk['Variance (UGX)'] < -100000) & (df_risk['WarehouseId'].isin(df_risky_mtd['WarehouseId'].tolist()))] 
num_cols = df_high_risk.select_dtypes(include='number').columns
df_high_risk[num_cols] = df_high_risk[num_cols].applymap(lambda x: f"{x:,.0f}")
print(df_high_risk.head(10))

#daily transactions for the risky

df_risky_daily = df_risk[(df_risk['WarehouseId'].isin(df_risky_mtd['WarehouseId'].tolist()))]
num_cols = df_risky_daily.select_dtypes(include='number').columns
df_risky_daily[num_cols] = df_risky_daily[num_cols].applymap(lambda x: f"{x:,.0f}")
print(df_risky_daily.head(10))
#%%
risk_report = pd.ExcelWriter(f"{file_path_r}Banking_Risk_Report_{today.strftime('%Y%m%d')}.xlsx", engine='xlsxwriter')
df_risky_mtd.to_excel(risk_report, sheet_name='Total_Unbanked_48hrs', index=False)
df_high_risk.to_excel(risk_report, sheet_name='Specific_days_unbanked', index=False)
df_risky_daily.to_excel(risk_report, sheet_name='Daily_Record_For_Risky_Shops', index=False)
risk_report.close()
kill_browser()

time.sleep(5)

if df_risky_mtd.shape[0] == 0:
    print("No risky shops identified for the current month.")
else:
    from gmail_sender import *

    gmail_function(
        subject_line = f"Banking Risk Report for {today.strftime('%Y-%m-%d')}",
        email_body = "Please find attached the Banking Risk Report for the current month.",
        receipient_email = 'pokuttu@yalelo.ug, dmukisa@yalelo.ug',
        attachment_path = f"{file_path_r}Banking_Risk_Report_{today.strftime('%Y%m%d')}.xlsx"
    )
    # delete the report after sending
    time.sleep(5)
    os.remove(f"{file_path_r}Banking_Risk_Report_{today.strftime('%Y%m%d')}.xlsx")
    print("Report sent and deleted successfully.")