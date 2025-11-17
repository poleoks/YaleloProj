import mysql.connector

conn=mysql.connector.connector(
    host:'127.0.0.1',
    user:'powerbi',
    password:'Y@l3l0@2023',
    port:3306,
    database:'logmanager1',
    auth_plugin='mysql_native_password'
)

my_cursor=conn.cursor()
conn.commit()
conn.close()