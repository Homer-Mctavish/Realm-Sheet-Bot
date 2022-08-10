import pymssql 
import pygsheets
import pandas as pd

conn = pymssql.connect("192.168.250.245:1433", "realmsheets", "Realm140", "RealmTEST")
# cursor = conn.cursor()
# cursor.execute("SELECT TT1.[ItemCode],MAX(TT1.[ItemName]) as Description, MAX(TT2.Cost) as Cost, MAX(TT1.SalePrice) as SalePrice FROM(SELECT T1.[ItemCode], T1.[ItemName],[Price] as SalePrice FROM [RealmTEST].[dbo].[ITM1] T0 JOIN [RealmTEST].[dbo].[OITM] T1 ON T0.[ItemCode] = T1.[ItemCode] WHERE T0.[PriceList] = '1') TT1 JOIN (SELECT [ItemCode],[Price] as Cost FROM [RealmTEST].[dbo].[ITM1] T0 WHERE T0.[PriceList] = '3') TT2 ON TT1.ItemCode = TT2.ItemCode group by TT1.[ItemCode]")
gc = pygsheets.authorize('C:\\Users\\Intern\\Documents\\client_secret_137777125265-n8hl4bi41mbvvm26svs6ph9g93bj5hsp.apps.googleusercontent.com.json')
df = pd.read_sql_query("SELECT TT1.[ItemCode],MAX(TT1.[ItemName]) as Description, MAX(TT2.Cost) as Cost, MAX(TT1.SalePrice) as SalePrice FROM(SELECT T1.[ItemCode], T1.[ItemName],[Price] as SalePrice FROM [RealmTEST].[dbo].[ITM1] T0 JOIN [RealmTEST].[dbo].[OITM] T1 ON T0.[ItemCode] = T1.[ItemCode] WHERE T0.[PriceList] = '1') TT1 JOIN (SELECT [ItemCode],[Price] as Cost FROM [RealmTEST].[dbo].[ITM1] T0 WHERE T0.[PriceList] = '3') TT2 ON TT1.ItemCode = TT2.ItemCode group by TT1.[ItemCode]", conn)
# df = pd.DataFrame()
sh = gc.open('google calendar test')
wks = sh[0]
wks.set_dataframe(df,(1,1))


