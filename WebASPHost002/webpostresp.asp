﻿<!DOCTYPE html>
<html>
    <head>
        <title></title>
    </head>
<body>
    <% 
'declare the variables 
Dim Connection
Dim Recordset
Dim SQL

'declare the SQL statement that will query the database
SQL = "SELECT * FROM customer"
'create an instance of the ADO connection and recordset objects
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

'define the connection string, specify database driver
ConnString="DRIVER={SQL Server};SERVER=sql.freeasphost.net\MSSQL2016;UID=eddyko00_SampleDB;PWD=DBSamplePW;DATABASE=eddyko00_SampleDB"

'Open the connection to the database
Connection.Open ConnString

'Open the recordset object executing the SQL statement and return records 
Recordset.Open SQL,Connection

'first of all determine whether there are any records 
If Recordset.EOF Then 
Response.Write("No records returned.") 
Else 
'if there are records then loop through the fields 
Do While NOT Recordset.Eof   
Response.write Recordset("username")
Response.write Recordset("password")
Response.write "<br>"    
Recordset.MoveNext     
Loop
End If

'close the connection and recordset objects to free up resources
Recordset.Close
Set Recordset=nothing
Connection.Close
Set Connection=nothing

%>

Welcome
<%
response.write(request.form("fname"))
response.write(" " & request.form("lname"))
%>


</body>
</html>