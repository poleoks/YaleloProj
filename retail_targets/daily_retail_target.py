import pandas as pd

# Replace 'YOUR_EXCEL_URL' with the actual URL of your Excel file
excel_url = 'https://firstwavegroup.sharepoint.com/:x:/s/FirstWaveDA/Ed75NOr9xL9Blpbyb58eTGIBWSjZ2Ovl-Sb3esCrxajnhw?e=lUclNB'
import requests
import pandas as pd



# Replace 'YOUR_USERNAME' and 'YOUR_PASSWORD' with your authentication credentials
username = 'pokuttu@yalelo.ug'
password = 'Aligator@1'

# Set up authentication
auth = (username, password)

# Make a request to the Excel file URL with authentication
response = requests.get(excel_url, auth=auth)

# Check if the request was successful (status code 200)
if response.status_code == 200:
    # Read Excel file from the response content
    df = pd.read_excel(pd.BytesIO(response.content))
    
    # Display the DataFrame
    print(df)
else:
    print(f"Failed to retrieve data. Status code: {response.status_code}")
