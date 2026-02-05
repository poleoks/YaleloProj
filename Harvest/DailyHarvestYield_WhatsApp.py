#%%##

from harvest_data_logmanager_posts import *
## navigate the report and take screenshot
sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
from whatsapp_file_sign_in import *


#%%
try:
    for i in glob.glob('C:/Users/Administrator/Documents/Python_Automations/Harvest/*.png'):
        os.remove(i)
        print(f"Deleted file: {i}") 
except Exception as e:
    print(f'No file to delete or error occurred: {e}')

#%%
try:
    today = datetime.datetime.today()# - timedelta(days=3)
    currentdatetime = (datetime.datetime.today() - timedelta(days=19)).strftime('%Y-%m-%d %H:%M:%S')
    currentdatetime = today  #.strftime('%Y-%m-%d %H:%M:%S')
    currentdate = today.strftime('%Y-%m-%d')
    # df['datetime'] = pd.to_datetime(df['datetime'])
    df['netweight'] = df['netweight'].str.replace('kg','').astype('float')
    df['batch_number'] = df['batch_number'].str[:4] + "("+ df['batch_number'].str[7:] +")"
    df['number_of_pieces'] = df['number_of_pieces'].astype('int')
    df['date'] = df['datetime'].dt.strftime('%Y-%m-%d')
    df['timedifference_mins'] = (abs(df['datetime'] - currentdatetime).dt.total_seconds() / 60).astype('int')
    last_date =  df.tail(1)['date'].min()
    last_time = df.tail(1)['datetime'].dt.strftime('%H:%M:%S').min()
    start_time = df.head(1)['datetime'].dt.strftime('%H:%M:%S').min()
    print(f"Harvest_start: {start_time}, Harvest end: {last_time}")
    
    (xrows,ycols) = df.shape
    print(df.shape)

    #%%
    # Define the custom order for materials
    material_order = {
        'SXXL': 1,
        'XXL': 2,
        'Extra Large': 3,
        'Large': 4,
        'Medium': 5,
        'Small': 6,
        'Extra Small': 7,
        'HSXXL': 8,
        'HXXL': 9,
        'HExtra Large': 10,
        'HLarge': 11,
        'HMedium': 12,
        'HSmall': 13,
        'Below 150': 14,
        'Mortality': 15,
        'Sick Fish': 16,
        'Culled': 17
    }

    batch_number_order = {
        batch: idx for idx, batch in enumerate(sorted(df['batch_number'].unique()), start=1)
    }

        
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
    for i in glob.glob('C:/Users/Administrator/Documents/Python_Automations/Harvest/*.png'):
        os.remove(i)
        print(f"Deleted file: {i}") 
except Exception as e:
    print(f'No file to delete or error occurred: {e}')
    
time.sleep(5)
#%%
kill_browser("chrome")