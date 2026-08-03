#%%##
## navigate the report and take screenshot
import sys
sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
#%%##
import pandas as pd
import mysql.connector
import warnings
warnings.filterwarnings('ignore')
import mysql.connector
import pandas as pd
import time
from requests.auth import HTTPBasicAuth
import sys
import datetime
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
import glob
import os


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
    connect_timeout=20
)
# h.CreatedOn, b.BatchCode, h.Weight, h.Pieces, p.ProductName, p.Size
cursor = conn.cursor()

query = """
SELECT
CONVERT_TZ(h.CreatedOn, '+00:00', '+03:00') as datetime, 
b.BatchCode as batch_number, h.Weight as netweight, h.Pieces as number_of_pieces,
p.Size as material

FROM y360.harvest h
LEFT JOIN y360.Batch b ON h.BatchId = b.BatchId
LEFT JOIN y360.Product p ON h.ProductId = p.ProductId
LEFT JOIN y360.Crate c ON h.CrateId = c.CrateId
LEFT JOIN y360.Crate cn on h.NewCrateId = cn.CrateId
LEFT JOIN y360.Destination d ON h.DestinationId = d.DestinationId
LEFT JOIN y360.Destination dn on h.NewDestinationId = dn.DestinationId
where DATE(CONVERT_TZ(h.CreatedOn, '+00:00', '+03:00')) = CURDATE()
"""
cursor.execute(query)
rows = cursor.fetchall()
columns = [desc[0] for desc in cursor.description]

df = pd.DataFrame(rows, columns=columns)

cursor.close()
conn.close()


df.dropna(subset=['datetime'], inplace=True)
df.sort_values(by='datetime', inplace=True)


from whatsapp_file_sign_in import *
#%%
today = datetime.datetime.today()# - timedelta(days=3)
currentdatetime = (datetime.datetime.today() - timedelta(days=19)).strftime('%Y-%m-%d %H:%M:%S')
currentdatetime = today  #.strftime('%Y-%m-%d %H:%M:%S')
currentdate = today.strftime('%Y-%m-%d')
# df['datetime'] = pd.to_datetime(df['datetime'])
# df['netweight'] = df['netweight'].str.replace('kg','').astype('float')
df['batch_number'] = df['batch_number'].str[:4] + "("+ df['batch_number'].str[7:] +")"
# df['number_of_pieces'] = df['number_of_pieces'].astype('int')
df['date'] = df['datetime'].dt.strftime('%Y-%m-%d')
df['timedifference_mins'] = (abs(df['datetime'] - currentdatetime).dt.total_seconds() / 60).fillna(0).astype('int')
last_date =  df.tail(1)['date'].min()
last_time = df.tail(1)['datetime'].dt.strftime('%H:%M').min()
start_time = df.head(1)['datetime'].dt.strftime('%H:%M').min()
print(f"Harvest_start: {start_time}, Harvest end: {last_time}")


# Total sum of netweight
df['datetime'] = pd.to_datetime(df['datetime'])
df['netweight'] = df['netweight'].fillna(0).astype(float)
df['number_of_pieces'] = df['number_of_pieces'].fillna(0).astype(int)
total_weight = df['netweight'].astype(float).sum() / 1000  # Convert to tons

# Total duration in hours
time_span = df['datetime'].max() - df['datetime'].min()
total_hours = time_span.total_seconds() / 3600 
total_minutes = time_span.total_seconds() / 60

# 2. Split into total hours and remaining minutes
hours_t, minutes_t = divmod(total_minutes, 60)

# 3. Cast to integers
hours_t = int(hours_t)
minutes_t = int(minutes_t)


# Calculation
if total_hours > 0:
    avg_weight_per_hour = total_weight / (total_hours)
else:
    avg_weight_per_hour = 0

print(f"Total Weight: {total_weight:.2f}T")
print(f"Total Hours: {total_hours:.2f}")
print(f"Ton/Hr: {avg_weight_per_hour:.2f}")

(xrows,ycols) = df.shape
print(df.shape)

#%%
# Define the custom order for materials
material_order = {
    "Below 150":9,
    "Carcass-HSXXL":21,
    "Carcass-SXXL":22,
    "Culled Fish":10,
    "Extra Large":4,
    "Extra Small":8,
    "Fillet Skin On SXXL":11,
    "Fillet- Skinless -L":18,
    "Fillet-Skinless HSXXL":12,
    "Fillet-Skinless HXXL":13,
    "Fillet-Skinless M":19,
    "Fillet-Skinless Medium":20,
    "Fillet-Skinless SXXL":15,
    "Fillet-Skinless XL":17,
    "Fillet-Skinless XXL":16,
    "Fillets-Skinless HXL":14,
    "Fresh G&S Large":26,
    "Fresh G&S Medium":27,
    "Fresh G&S Small":28,
    "Fresh G&S SXXL":23,
    "Fresh G&S XL":25,
    "Fresh G&S XXL":24,
    "HExtra Large":31,
    "HLarge":32,
    "HMD":35,
    "HMedium":33,
    "HSmall":34,
    "HSXXL":29,
    "HXXL":30,
    "Large":5,
    "Large-Gutted Scales On":40,
    "Medium":6,
    "Mortality":36,
    "RedMeat-SXXL":37,
    "Salted-CulledFish":38,
    "Skins-SXXL":39,
    "Small":7,
    "SXXL":1,
    "XXL":2,
    "XXL-Gutted Scales On":41,
    "XXXL":3
}

# batch_number_order = {
#     batch: idx for idx, batch in enumerate(sorted(df['batch_number'].unique()), start=1)
# }
# Convert the unique values to strings before sorting
unique_batches = sorted(df['batch_number'].astype(str).unique())

batch_number_order = {batch: idx for idx, batch in enumerate(unique_batches, start=1)}
    
#%%
if (df.sort_values(by='datetime').tail(1)['timedifference_mins'].min() < 120):# and xrows > 10:
    # Group by and aggregate
    df2 = df[df['date'] == currentdate][['batch_number', 'material', 'netweight', 'number_of_pieces']].groupby(['batch_number', 'material']).agg(
        total_weight_kg=('netweight', 'sum'),
        total_pieces=('number_of_pieces', 'sum'),
        total_crates=('batch_number', 'count')
    )

    s = df2.groupby(level=0).sum()
    s.index = pd.MultiIndex.from_product([s.index, ['**BATCH SUMMARY**'], ['']])

    # Concatenate and apply custom sorting
    dd = pd.concat([df2, s])

    # Custom sorting based on both batch_number and material
    dd = dd.sort_index(level=[0, 1], key=lambda idx: (
        idx.map(batch_number_order) if idx.name == 'batch_number' else idx.map(material_order)
    ))

    # Add 'Grand Total' row
    dd.loc['Grand Total', :] = df2.sum().values

    # Additional formatting
    # dd['ABW (g)'] = (dd.total_weight_kg * 1000 / dd.total_pieces).astype('int').apply(lambda x: '{:,}'.format(x))
    dd['total_crates'] = dd['total_crates'].astype('int').apply(lambda x: '{:,}'.format(x))
    # dd['total_pieces'] = dd['total_pieces'].astype('int').apply(lambda x: '{:,}'.format(x))
    # dd['total_weight_kg'] = dd['total_weight_kg'].astype('float').round(2).apply(lambda x: '{:,}'.format(x)) 
    
    
    dd['total_weight_kg'] = pd.to_numeric(
        dd['total_weight_kg'].astype(str).str.replace(',', ''),
        errors='coerce'
    ).round(2)

    dd['total_pieces'] = pd.to_numeric(
        dd['total_pieces'].astype(str).str.replace(',', ''),
        errors='coerce'
    ).astype('Int64')

    dd['ABW (g)'] = (
        (dd['total_weight_kg'] * 1000) / dd['total_pieces']
    ).round(0).astype('Int64')
    
    print(f"dd.head():\n{dd.head()}")
    
    dd['total_weight_kg'] = dd['total_weight_kg'].round(2).apply(lambda x: f"{x:,.2f}")
    dd['total_pieces'] = dd['total_pieces'].apply(lambda x: f"{x:,}")
    dd['total_crates'] = dd['total_crates'].fillna(0).astype('int').apply(lambda x: f"{x:,}")
    dd['ABW (g)'] = dd['ABW (g)'].apply(lambda x: f"{x:,}")

    # Reset index and adjust appearance
    # dd[['ABW (g)', 'total_crates', 'total_pieces', 'total_weight_kg']] = dd[['ABW (g)', 'total_crates', 'total_pieces', 'total_weight_kg']].applymap(lambda x: '{:,}'.format(x) if pd.notnull(x) else '')
    
    dd = dd.reset_index()
    time.sleep(5)

    # Replace duplicate 'batch_number' values with empty strings
    dd['batch_number'] = dd['batch_number'].mask(dd['batch_number'].duplicated(), '')
    dd.rename(columns={'level_1': 'SKU'}, inplace=True)

    #----------Material----------
    #____________________________________________________________________________________________________________________________________________________________________________________
    df2 = df[df['date'] == currentdate]

    
    dd2 = df2.groupby('material', as_index=False).agg(
            total_weight_kg=('netweight', 'sum'),
            total_pieces=('number_of_pieces', 'sum'),
            total_crates=('batch_number', 'count')
        )
    
    print(dd2.head())
    
    
    
    dd2.info()
    
    dd2['ABW (g)'] = (dd2.total_weight_kg * 1000 / dd2.total_pieces).astype(int)
    dd2 = dd2.sort_values(by='total_weight_kg', ascending=False)

    dd2['total_weight_kg'] = dd2['total_weight_kg'].round(2).apply(lambda x: f"{x:,.2f}")
    dd2['total_pieces'] = dd2['total_pieces'].apply(lambda x: f"{x:,}")
    dd2['total_crates'] = dd2['total_crates'].apply(lambda x: f"{x:,}")
    dd2['ABW (g)'] = dd2['ABW (g)'].apply(lambda x: f"{x:,}")

    dd2.loc[len(dd2.index)] = [
        'Grand Total',
        f"{df2['netweight'].sum():,.2f}",
        f"{df2['number_of_pieces'].sum():,}",
        f"{len(df2):,}",
        f"{(df2.netweight.sum() * 1000/df2.number_of_pieces.sum()).astype(int): ,}"
    ]
    
    # Style and export
    
    df_styled = (dd.style.hide(axis='index').set_caption("Batch-Level Harvest Summary").set_properties(**{'background-color': "#67C7F3"}).set_table_styles([{'selector': 'th', 'props': [('background-color', "#0C0B0B"), ('color', 'white')]}] ) )
    df2_styled = (dd2.style.hide(axis='index').set_caption("SKU-Level Summary").set_properties(**{'background-color': "#F3D173"}).set_table_styles([{'selector': 'th', 'props': [('background-color', '#404040'), ('color', 'white')]}] ))

    
    # Export the DataFrame as an image
    batch_img = 'C:/Users/Administrator/Documents/Python_Automations/Harvest/harvest_batch.png'
    mat_img ='C:/Users/Administrator/Documents/Python_Automations/Harvest/harvest_sku.png'
    final_img = 'C:/Users/Administrator/Documents/Python_Automations/Harvest/harvest.png'

    dfi.export(df_styled, batch_img, table_conversion='matplotlib', max_rows=None, max_cols=None, fontsize=12)
    dfi.export(df2_styled, mat_img, table_conversion='matplotlib', max_rows=None, max_cols=None, fontsize=12)

    img1 = Image.open(batch_img)
    img2 = Image.open(mat_img)

    combined_height = img1.height + img2.height + 60
    combined_width = max(img1.width, img2.width)
    new_img = Image.new('RGB', (combined_width, combined_height), color='white')

    new_img.paste(img1, (0, 0))
    new_img.paste(img2, (0, img1.height + 60))
    new_img.save(final_img)
    time.sleep(5)

    #INSTANTIATE WHATSAPP
    files_t =['harvest.png']
    groups_t = ['YU S&OP Planning Cell']
    messages_t = [f"Harvest Report\nStart-End: {start_time}-{last_time}\nTotal Weight: {total_weight:.2f}T, \nTotal Time: {hours_t}h {minutes_t}m \nT/H: {avg_weight_per_hour:.2f}"]
    directory_t = "C:/Users/Administrator/Documents/Python_Automations/Harvest/"

    whatsapp_share(groups_t, messages_t,files_t, directory_t, Pole)

    
else:
    print("No Latest Harvest Data!")
    pass
#%%
try:
    for i in glob.glob('C:/Users/Administrator/Documents/Python_Automations/Harvest/*.png') or glob.glob('C:/Users/Administrator/Documents/Python_Automations/Harvest/*.xlsx'):
        os.remove(i)
        print(f"Deleted file: {i}") 
except Exception as e:
    print(f'No file to delete or error occurred: {e}')
    
time.sleep(5)
#%%
kill_browser("chrome")