<!DOCTYPE html>
<html>
    <head>
        <title></title>
    </head>
<body>
Welcome
<%
response.write(request.querystring("fname"))
response.write(" " & request.querystring("lname"))
%>
</body>
</html>