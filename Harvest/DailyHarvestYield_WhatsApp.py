#%%##
## navigate the report and take screenshot
import sys
sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")

# from harvest_data_logmanager_posts import *
from powerbi_sign_in_file import *
try:
    for i in glob.glob('C:/Users/Administrator/Documents/Python_Automations/Harvest/*.png') or glob.glob('C:/Users/Administrator/Documents/Python_Automations/Harvest/*.xlsx') or glob.glob('C:/Users/Administrator/Downloads/*.xlsx')  :
        os.remove(i)
        print(f"Deleted file: {i}") 
except Exception as e:
    print(f'No file to delete or error occurred: {e}')
#%%
download_path = "C:/Users/Administrator/Downloads/data.xlsx"
df = pbi_export("https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/e3d968fc-5351-4d04-ab9c-0e7c241a98ac/bd4fa43b35e81444b15e",
           download_path)

print(f"Data exported to: {download_path}")
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
total_weight = df['netweight'].astype(float).sum() / 1000  # Convert to tons

# Total duration in hours
time_span = df['datetime'].max() - df['datetime'].min()
total_hours = time_span.total_seconds() / 3600 

# Calculation
if total_hours > 0:
    avg_weight_per_hour = total_weight / (total_hours)
else:
    avg_weight_per_hour = 0

print(f"Total Weight: {total_weight:.2f} kg")
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
try: 
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
        dd['ABW (g)'] = (dd.total_weight_kg * 1000 / dd.total_pieces).astype('int').apply(lambda x: '{:,}'.format(x))
        dd['total_crates'] = dd['total_crates'].astype('int').apply(lambda x: '{:,}'.format(x))
        dd['total_pieces'] = dd['total_pieces'].astype('int').apply(lambda x: '{:,}'.format(x))
        dd['total_weight_kg'] = dd['total_weight_kg'].astype('float').round(2).apply(lambda x: '{:,}'.format(x))

        # Reset index and adjust appearance
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
        messages_t = [f"Harvest Report\nStart-End: {start_time}-{last_time}\nTotal Wt: {total_weight:.2f}T, Total Hours: {total_hours:.0f} \nT/H: {avg_weight_per_hour:.2f}"]
        directory_t = "C:/Users/Administrator/Documents/Python_Automations/Harvest/"

        whatsapp_share(groups_t, messages_t,files_t, directory_t, Pole)

        
    else:
        print("No Latest Harvest Data!")
        pass
except:
    from gmail_sender import *
    gmail_function('pokuttu@yalelo.ug','Harvest Update On Whatsapp Has Failed',
               'Hello, the harvest subscription for Whatsapp has failed to go. Please look into it.',''
               )
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