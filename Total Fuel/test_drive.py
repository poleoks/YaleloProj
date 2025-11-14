import pandas as pd

# Sample data (same as above)
data = {'date': ['2023-08-15', '2023-09-20', '2023-10-10', '2023-11-25', '2023-12-31',
                 '2024-01-10', '2024-02-15', '2024-03-20', '2024-04-25', '2024-05-31'],
        'month_of_promotion': ['Aug, Sep', 'Sep, Oct, Nov', 'Dec', 'Jan, Feb', 'Mar', '', '', 'Apr, May', '', '']}

df = pd.DataFrame(data)

# Extract month from date
df['month'] = pd.to_datetime(df['date']).dt.month_name().str[:3]

df['month_of_promotion']=df['month_of_promotion'].apply(lambda x: x.split(',') if x else [])
df['month_of_promotion']=df['month_of_promotion'].apply(lambda y: '' if y==[] else y)


df['j']=df.apply(lambda row: row['month'] in row['month_of_promotion'], axis=1)
print(df)
# print(row(df.month))



