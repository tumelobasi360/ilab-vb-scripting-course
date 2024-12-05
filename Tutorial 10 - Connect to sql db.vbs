'********************************************
'       Tutorial 10 - Connect To DB
'********************************************

Option Explicit

' Declare
DIM strConnection
DIM objConnecting
DIM rs ' Recordset
DIM strSELECT, strResults

' Assign
strConnection = "Provider=sqloledb;Server=ILAB-008\SQLEXPRESS;Database=Northwind;Integrated Security=SSPI"
SET objConnecting = CreateObject("ADODB.Connection")
SET rs = CreateObject("ADODB.Recordset")

' Initialize variables
strResults = ""

' Open connection
objConnecting.Open strConnection

' SQL Query
strSELECT = "SELECT * FROM Shippers ORDER BY ShipperID"

' Open recordset
rs.Open strSELECT, objConnecting

' Loop through records
DO WHILE NOT rs.EOF
    strResults = strResults & rs.Fields.Item("ShipperID").Value & " - " & rs.Fields.Item("CompanyName").Value & vbNewLine
    rs.MoveNext
LOOP

' Display results
MsgBox strResults

' Cleanup
rs.Close
objConnecting.Close
SET rs = Nothing
SET objConnecting = Nothing
