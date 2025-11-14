import pandas as pd
# country_ar_extract='C:/Users/Administrator/Downloads/Customer aging report.xlsx'
# customer_details_extract='C:/Users/Administrator/Downloads/data.xlsx'


def receivables(country_ar_extract,customer_details_extract):
    df=pd.read_excel(country_ar_extract, skiprows=11 )
    df = df[df.columns.drop(list(df.filter(regex='Unnamed')))]
    df=df.iloc[:-1,:]
    # df.rename(columns={ df.columns[3]:"Balance as of today" }, inplace = True)

    f_col=pd.read_excel(country_ar_extract, skiprows=10 ).head(1)
    f_col=f_col[f_col.columns.drop(list(f_col.filter(regex='Unnamed')))]


    series1=pd.Series(pd.to_datetime(f_col.iloc[:,1:].columns).astype(str))
    series2=pd.Series(pd.to_datetime(df.iloc[:,5:].columns).astype(str))
    result = series1 + "~" + series2
    result=result.to_list()

    # Create a dictionary mapping old names to new names
    rename_dict = dict(zip(df.columns[-4:], result))

    # Rename the last four columns
    df = df.rename(columns=rename_dict)


    #prefix balance as of
    prefix = 'Balance_as_at:'

    column_names_4th_and_5th = pd.Series(pd.to_datetime(df.columns[3:5]).astype(str))

    # Modify the column names by adding the prefix
    new_column_names = [prefix + column_name for column_name in column_names_4th_and_5th]

    # Create a dictionary mapping old names to new names
    rename_dict = dict(zip(df.columns[3:5], new_column_names))

    # Rename the 4th and 5th columns
    df = df.rename(columns=rename_dict)
    customer_details=df

    # print(customer_details.head(3))

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

    return country_ar_summary, full_ar