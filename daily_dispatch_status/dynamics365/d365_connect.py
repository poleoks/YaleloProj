import pyodata
import requests

SERVICE_URL = 'http://services.odata.org/V2/Northwind/Northwind.svc/'
SERVICE_URL = 'https://fw-d365-prod.operations.dynamics.com/data'
HTTP_LIB = requests.Session()

northwind = pyodata.Client(SERVICE_URL, HTTP_LIB)

for customer in northwind.entity_sets.Customers.get_entities().execute():
    print(customer.CustomerID, customer.CompanyName)
