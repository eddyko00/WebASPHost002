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
Dim cmd
Dim cmdreq
Dim cmdpost
Dim req
Dim SQL
Dim SQLResp

'declare the SQL statement that will query the database
'setRequestHeader "Content-type", "application/x-www-form-urlencoded";
SQL = "SELECT * FROM account"
cmdreq = request.QueryString("cmd")
cmdpost = request.form("cmdpost")
req  = request.form("req")
        
cmd =0
If not isnull(cmdreq) Then
    if len(cmdreq) > 0 Then
    cmd = cmdreq
    End if
End if

If not isnull(cmdpost) Then
    if len(cmdpost) > 0 Then
    cmd = cmdpost
    End if
End if

SQL = req

Response.Write("cmd=" & cmd) 
Response.write "<br>" 
Response.Write("cmdpost=" & cmdpost) 
Response.write "<br>" 

If cmd = 0 Then
    Response.Write("SQL=" & SQL) 
    Response.write "<br>" 
End If

'create an instance of the ADO connection and recordset objects
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

'define the connection string, specify database driver
'remember to add a ; at the end
'remember to add a ; at the end
ConnString="DRIVER={SQL Server};SERVER=sql.freeasphost.net\MSSQL2016;UID=eddyko00_SampleDB;PWD=DBSamplePW;DATABASE=eddyko00_SampleDB;"

If cmd = "3" Then
    'Open the connection to the database
    Connection.Open ConnString

    If Len(SQL) > 0 Then
        Dim updateParamArray
        updateParamArray = Split(SQL , "~")
       
        If UBound(updateParamArray ) > -1 Then
          Dim i
          Dim name, param

          For i = 0 To UBound(updateParamArray )
            param = updateParamArray(i)
            name = ""
            name = param
            If name <> "" Then
                'Response.write (i & " " & name)
                'Response.write ("<br>")      
                Connection.Execute name,adExecuteNoRecords  
            End If
          Next
        End If
    End If 
    Response.write ("~~ ") 
    Response.write (i-1)
    Response.write (" ~~")  

    ' close the connection
    Connection.Close
    Set Connection=nothing
End If  

If cmd = "2" Then
    'Open the connection to the database
    Connection.Open ConnString

    Connection.Execute SQL,adExecuteNoRecords  
    Response.write ("~~ ") 
    Response.write (adExecuteNoRecords)
    Response.write (" ~~")  
    ' close the connection
    Connection.Close
    Set Connection=nothing
End If  
        
If cmd = "1" Then

    'Open the connection to the database
    Connection.Open ConnString

    'Open the recordset object executing the SQL statement and return records 
    Recordset.Open SQL,Connection

    'first of all determine whether there are any records 
    If Recordset.EOF Then 
        Response.Write(" ") 
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