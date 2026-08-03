
import pandas as pd
import mysql.connector

server = "y360db.mysql.database.azure.com"
database = "y360"
username = "moweigh"
password = "6Million$"
#%%


conn = mysql.connector.connect(
    host="y360db.mysql.database.azure.com",
    user="moweigh",
    password="6Million$",
    database="y360", 
    connect_timeout=10
)
# h.CreatedOn, b.BatchCode, h.Weight, h.Pieces, p.ProductName, p.Size
cursor = conn.cursor()

query = """
SELECT

h.CreatedOn, h.HarvestId,
b.BatchCode, h.Weight, h.Pieces, p.ProductName,
p.Size, c.CrateCode, cn.CrateCode as NewCrateCode,
d.DestinationName, dn.DestinationName as NewDestinationName

FROM y360.harvest h
LEFT JOIN y360.Batch b ON h.BatchId = b.BatchId
LEFT JOIN y360.Product p ON h.ProductId = p.ProductId
LEFT JOIN y360.Crate c ON h.CrateId = c.CrateId
LEFT JOIN y360.Crate cn on h.NewCrateId = cn.CrateId
LEFT JOIN y360.Destination d ON h.DestinationId = d.DestinationId
LEFT JOIN y360.Destination dn on h.NewDestinationId = dn.DestinationId
where DATE(h.CreatedOn) = CURDATE() and h.Ispacked = 1 and dn.DestinationName ilike 'Gulu%'
"""
cursor.execute(query)
rows = cursor.fetchall()
columns = [desc[0] for desc in cursor.description]

df = pd.DataFrame(rows, columns=columns)

cursor.close()
conn.close()

df.to_csv('harvest_data.csv', index=False)
print(df.head())

# print(columns)
#%%