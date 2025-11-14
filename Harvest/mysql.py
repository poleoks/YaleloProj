import mysql.connector
import pandas as pd

#%%
conn=mysql.connector.connect(
    host='127.0.0.1',
    user='powerbi',
    password='Y@l3l0@2023',
    # port=3306,
    database='logmanager1'
)

my_cursor=conn.cursor()

# SQL query to select top 5 rows
sql = "SELECT * FROM productionlogsheet LIMIT 5"

# Execute the query
my_cursor.execute(sql)

# Fetch the results
results = my_cursor.fetchall()
# print(results)
col_names=[i[0] for i in my_cursor.description]
# Print the results
# for row in results:
#     print(row)
harvest_data = pd.DataFrame(results, columns=col_names)
print(harvest_data)
# conn.commit()
my_cursor.close()
conn.close()

print("Successfully connected")