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
Dim SQLResp

'declare the SQL statement that will query the database
SQL = "SELECT * FROM customer"
cmd  = request.form("cmd")
req  = request.form("req")
resp = request.form("resp")


If isNULL(req) Then
    SQL = req
End If

Response.Write(cmd) 
Response.write "<br>" 
Response.Write(req) 
Response.write "<br>" 
Response.Write(resp) 
Response.write "<br>" 
Response.Write(SQL) 
Response.write "<br>" 

'create an instance of the ADO connection and recordset objects
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

'define the connection string, specify database driver
'remember to add a ; at the end
'remember to add a ; at the end
ConnString="DRIVER={SQL Server};SERVER=sql.freeasphost.net\MSSQL2016;UID=eddyko00_SampleDB;PWD=DBSamplePW;DATABASE=eddyko00_SampleDB;"

If cmd = "1" Then

    'Open the connection to the database
    Connection.Open ConnString

    'Open the recordset object executing the SQL statement and return records 
    Recordset.Open SQL,Connection

    'first of all determine whether there are any records 
    If Recordset.EOF Then 
        Response.Write("No records returned.") 
    Else 
        SQLResp = SQLResp & "["

        Do While NOT Recordset.Eof  
            SQLResp = SQLResp & "{"
            'if there are records then loop through the fields 
            first=0
            for each x in Recordset.fields
                If first > 0 Then
                    SQLResp = SQLResp & ","
                End If
                first = first + 1
                SQLResp = SQLResp & """"
                SQLResp = SQLResp & x.name
                SQLResp = SQLResp & """:"
                SQLResp = SQLResp & """"
                SQLResp = SQLResp & x.value
                SQLResp = SQLResp & """"
                
            next

            Recordset.MoveNext
            If Recordset.EOF Then 
                SQLResp = SQLResp & "}"
            Else 
                SQLResp = SQLResp & "},"
            End If        
        Loop
        SQLResp = SQLResp &"]"
    End If
    Response.write ("~~ ") 
    Response.write (SQLResp)
    Response.write (" ~~")         
    'close the connection and recordset objects to free up resources
    Recordset.Close
    Set Recordset=nothing
    Connection.Close
    Set Connection=nothing

End If
%>

</body>
</html>