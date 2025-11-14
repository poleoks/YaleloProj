import pandas as pd
import datetime as dt
date_today=dt.date.today()
print()
<<<<<<< HEAD
country_ar_extract='C:/Users/Administrator/Downloads/Customer aging report.xlsx'
customer_details_extract='C:/Users/Administrator/Downloads/data.xlsx'

# bb=f"As at: {date_today} [ugx]"

# def ug_extract():
#     df=pd.read_excel(country_ar_extract,skiprows=11,skipfooter=1)
#     df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
#     df.rename(columns={ df.columns[3]:bb,df.columns[4]:"90+ days [ugx]",df.columns[5]:"90 days [ugx]",df.columns[6]:"60 days [ugx]",
#                     df.columns[7]:"30 days [ugx]",df.columns[8]:"Current Date [ugx]"}, inplace = True)
#     customer_details=df
#     #%%
#     ar=pd.read_excel(customer_details_extract)
#     print(ar.head(2))


#     # full_ar=ar[['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode']].join(customer_details, how='left')

#     full_ar=pd.merge(ar,customer_details, right_on='Account', left_on='CustomerAccount', how='left')
#     print(full_ar.head())


#     sales_channel=full_ar.drop(columns=['Account','CustomerAccount','Name','Customer group'])

#     print(sales_channel.columns)


#     country_ar_summary=sales_channel.groupby(['Description', 'CustomerPriorityClassificationGroupCode']).sum()
#     country_ar_summary=country_ar_summary[country_ar_summary.iloc[:,3]>0]

#     full_ar=full_ar.groupby(['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode',
#                                     'Account','CustomerAccount','Name','Customer group']).sum()
#     print(country_ar_summary.head(),full_ar.head())

#     return country_ar_summary, full_ar

# def ug_extract():

bb=f"As at: {date_today} [kshs]"
df=pd.read_excel(country_ar_extract,skiprows=11,skipfooter=1)
df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
df.rename(columns={ df.columns[3]:bb,df.columns[4]:"Current (0-7) days [kshs]",df.columns[5]:"8-14 days [kshs]",df.columns[6]:"15-29 days [kshs]",
                df.columns[7]:"30-89 days [kshs]",df.columns[8]:"90+ [kshs]"}, inplace = True)
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
print(country_ar_summary.head(),full_ar.head())

# return country_ar_summary, full_ar
=======


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
    # print(country_ar_summary.head(),full_ar.head())

    return country_ar_summary, full_ar
>>>>>>> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
