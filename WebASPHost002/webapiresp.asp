<!DOCTYPE html>
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
cmd = request.form("cmd")
Dim sqlCmd as string = request.form("sql")

If Len(sqlCmd) > 0 Then
SQL = sqlCmd
End If
Response.Write(SQL) 

'create an instance of the ADO connection and recordset objects
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

'define the connection string, specify database driver
'remember to add a ; at the end
'remember to add a ; at the end
ConnString="DRIVER={SQL Server};SERVER=sql.freeasphost.net\MSSQL2016;UID=eddyko00_SampleDB;PWD=DBSamplePW;DATABASE=eddyko00_SampleDB;"

'Open the connection to the database
Connection.Open ConnString

'Open the recordset object executing the SQL statement and return records 
Recordset.Open SQL,Connection

'first of all determine whether there are any records 
If Recordset.EOF Then 
Response.Write("No records returned.") 
Else 
Response.write("[")

Do While NOT Recordset.Eof  
Response.write("{")
'if there are records then loop through the fields 
first=0
for each x in Recordset.fields
If first > 0 Then
Response.write(",")
End If
first = first + 1
Response.write("""")
  Response.write(x.name)
Response.write(""":")
  Response.write("""")
  Response.write(x.value)
  Response.write("""")
  Response.write(" ")
next

Recordset.MoveNext
If Recordset.EOF Then 
Response.write("}") 
Else 
Response.write("},") 
End If        
Loop
Response.write("]")
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