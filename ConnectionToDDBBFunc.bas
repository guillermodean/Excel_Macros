'Define the constants

Public Const sqlserver = "SQLSERVER"
Public Const DataBase = sqlserver
Public Conn As ADODB.Connection

Function connectDB() As Boolean
'Function to estalish and check the Database connection
connectDB = False

Set Conn = New ADODB.Connection
If Conn.State = 0 Then
      'Define the connection with provider MSOLEDBSQL string in this case with integrated security and DDBB as database name
      sConn = "Provider=MSOLEDBSQL;Data Source=" & DataBase & ";Initial Catalog=DDBB;Integrated Security=SSPI"
    With Conn
        'Pass the conection parameters to Conn
        .ConnectionString = sConn
        .ConnectionTimeout = 25
        .CommandTimeout = 35
        
    End With
    On Error Resume Next
    'Start connection
    Conn.Open
    On Error GoTo 0
End If
If Conn.State = 0 Then
    'Define the connection with provider SQLOLEDB string in this case with integrated security and DDBB as database name
    sConn = "Provider=SQLOLEDB;Data Source=" & DataBase & ";Initial Catalog=PM_EPC;Integrated Security=SSPI"
    With Conn
        .ConnectionString = sConn
        .ConnectionTimeout = 25
        .CommandTimeout = 35
        
    End With
    Conn.Open
End If

If Conn.State = adStateOpen Then connectDB = True

End Function
  
  'Extra info
  'MSOLEDBSQL = The OLE DB Driver for SQL Server is a stand-alone data access application programming interface (API), used for OLE DB, that was introduced in SQL Server 2005
  'SQLOLEDB = The Microsoft OLE DB Provider for SQL Server, SQLOLEDB, allows ADO to access Microsoft SQL Server.
  ' Find the differences between drivers in https://learn.microsoft.com/en-us/sql/connect/oledb/major-version-differences?view=sql-server-ver16
