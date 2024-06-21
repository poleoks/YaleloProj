import pandas as pd
import datetime as dt
from datetime import datetime#,
date_today=dt.date.today()
print()
date_ar= datetime.today().strftime('%Y-%m-%d')


# country_ar_extract='P:/Downloads/FI_Downloads/Customer aging report.xlsx'
# customer_details_extract='P:/Downloads/FI_Downloads/data.xlsx'

def ug_extract(country_ar_extract,customer_details_extract):
    bb=f"As at: {date_today} [ugx]"
    df=pd.read_excel(country_ar_extract,skiprows=11,skipfooter=1)
    df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
    df.rename(columns={ df.columns[3]:bb,df.columns[4]:"90+ days [ugx]",df.columns[5]:"90 days [ugx]",df.columns[6]:"60 days [ugx]",
                    df.columns[7]:"30 days [ugx]",df.columns[8]:"Current Date [ugx]"}, inplace = True)
    customer_details=df
    #%%
    ar=pd.read_excel(customer_details_extract)
    print(ar.head(2))

    # full_ar=ar[['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode']].join(customer_details, how='left')

    full_ar=pd.merge(ar,customer_details, right_on='Account', left_on='CustomerAccount', how='left')
    print(full_ar.head())


    sales_channel=full_ar.drop(columns=['Account','CustomerAccount','Name','Customer group'])

    print(sales_channel.columns)


    country_ar_summary=sales_channel.groupby(['Description', 'CustomerPriorityClassificationGroupCode']).sum()
    country_ar_summary=country_ar_summary[country_ar_summary.iloc[:,3]>0]

    full_ar=full_ar.groupby(['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode',
                                    'Account','CustomerAccount','Name','Customer group']).sum()
    # print(country_ar_summary.head(),full_ar.head())
    
    # country_ar_summary.sort_values(by=f'As at: {date_ar} [ugx]', ascending=False,inplace=True).drop(columns='OrganizationName', axis=1)
    # full_ar.sort_values(by=f'As at: {date_ar} [ugx]', ascending=False,inplace=True)
# kenya_ar.sort_values(by=f'As at: {date_ar} [kshs]', ascending=False,inplace=True).drop(columns='OrganizationName', axis=1).to_excel(AR_Standing_status, sheet_name='By Sales Channel - KY',index=False)
# customer_ky.sort_values(by=f'As at: {date_ar} [kshs]', ascending=False,inplace=True).drop(columns=['Account','Name','CustomerAccount.1'],axis=1).to_excel(AR_Standing_status, sheet_name='Customer Details KY',index=False)


    country_ar_summary.iloc[:,3:]=country_ar_summary.iloc[:,3:].astype('int')
    df=pd.DataFrame(country_ar_summary.iloc[:,3:].sum().to_frame().T)
    country_ar_summary=pd.concat([country_ar_summary,df], ignore_index=True)

    country_ar_summary.sort_values(by=f'As at: {date_ar} [ugx]', ascending=False,inplace=True)
    country_ar_summary.iloc[-1, 0] = 'Total'  
    country_ar_summary.fillna('',inplace=True)

    full_ar.iloc[:,8:]=full_ar.iloc[:,8:].astype('int')
    df=pd.DataFrame(full_ar.iloc[:,8:].sum().to_frame().T)
    full_ar=pd.concat([full_ar,df], ignore_index=True)

    full_ar.sort_values(by=f'As at: {date_ar} [ugx]', ascending=False,inplace=True)

    full_ar.iloc[-1, 0] = 'Total'  
    full_ar.fillna('',inplace=True)

    return country_ar_summary, full_ar

def ke_extract(country_ar_extract,customer_details_extract):

    bb=f"As at: {date_today} [kshs]"
    df=pd.read_excel(country_ar_extract,skiprows=11,skipfooter=1)
    df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
    df.rename(columns={ df.columns[3]:bb,df.columns[4]:"Current (0-7) days [kshs]",df.columns[5]:"8-14 days [kshs]",df.columns[6]:"15-29 days [kshs]",
                    df.columns[7]:"30-89 days [kshs]",df.columns[8]:"90+ [kshs]"}, inplace = True)
    customer_details=df
    #%%
    ar=pd.read_excel(customer_details_extract)
    # print(df.head(2))
    # full_ar=ar[['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode']].join(customer_details, how='left')

    full_ar=pd.merge(ar,customer_details, right_on='Account', left_on='CustomerAccount', how='left')
    sales_channel=full_ar.drop(columns=['Account','CustomerAccount','Name','Customer group'])

    # print(sales_channel.columns)
    country_ar_summary=sales_channel.groupby(['Description', 'CustomerPriorityClassificationGroupCode']).sum()
    country_ar_summary=country_ar_summary[country_ar_summary.iloc[:,3]>0]

    full_ar=full_ar.groupby(['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode',
                                    'Account','CustomerAccount','Name','Customer group']).sum()
    
    



    country_ar_summary.iloc[:,3:]=country_ar_summary.iloc[:,3:].astype('int')
    df=pd.DataFrame(country_ar_summary.iloc[:,3:].sum().to_frame().T)
    country_ar_summary=pd.concat([country_ar_summary,df], ignore_index=True)

    print(f"cty-ar-cols:{country_ar_summary.columns}")
    country_ar_summary.sort_values(by=f'As at: {date_ar} [kshs]', ascending=False,inplace=True)
# country_ar_summary.sort_values(by=f'As at: {date_ar} [kshs]', ascending=False,inplace=True).drop(columns=['Account','Name','CustomerAccount'],axis=1)
    country_ar_summary.iloc[-1, 0] = 'Total'  
    country_ar_summary.fillna('',inplace=True)

    full_ar.iloc[:,8:]=full_ar.iloc[:,8:].astype('int')
    df=pd.DataFrame(full_ar.iloc[:,8:].sum().to_frame().T)
    full_ar=pd.concat([full_ar,df], ignore_index=True)
    print(f"f_ar-cols{full_ar.columns}")


    full_ar.sort_values(by=f'As at: {date_ar} [kshs]', ascending=False,inplace=True)
    # full_ar.sort_values(by=f'As at: {date_ar} [kshs]', ascending=False,inplace=True).drop(columns='OrganizationName', axis=1)

    full_ar.iloc[-1, 0] = 'Total'  
    full_ar.fillna('',inplace=True)


    return country_ar_summary, full_ar