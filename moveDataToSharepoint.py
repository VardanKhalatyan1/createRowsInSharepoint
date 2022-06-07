import pandas as pd
import pyodbc
from datetime import datetime
from requests_ntlm import HttpNtlmAuth
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import urllib3



def get_data():
    server = 'server'
    database = 'database'
    table = 'table'
    sql_conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER={ server };DATABASE={ database };Trusted_Connection=yes'.format(server=server , database=database))
    query=r"SELECT create_time,card_no,state,firstname,lastname,id FROM { table } where create_time > { date }".format(table=table , date=datetime.today() - timedelta(days = 1 ))
    df = pd.read_sql(query, sql_conn)
    data=df.values
    upload_data(data)


def upload_data(row_data):
    my_data=[]
    username = "user"
    password = "pass"
    urllib3.disable_warnings()
    authcookie = Office365('https://bysmartclick.sharepoint.com/', username=username, password=password).GetCookies()
    site = Site('https://bysmartclick.sharepoint.com/sites/leaverequest', version=Version.v365, authcookie=authcookie)
    sp_list = site.List('in-out')
    for i in row_data:
        name=i[3] + " " + i[4]
        my_data.append({'Title':i[5],'create_time':i[0],'card_no':i[1],'state':i[2], 'Full Name':name})

    sp_list.UpdateListItems(data=my_data, kind='New')


get_data()