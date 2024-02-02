<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim name, position
    name = Request.Form("name")
    position = Request.Form("position")

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=SQLite3 Datasource;Database=C:/src/vb-script-sql/company.db;"

    conn.Execute "INSERT INTO Employees (Name, Position) VALUES ('" & name & "', '" & position & "')"
    new_id = conn.Execute("SELECT last_insert_rowid()")


    conn.Close
    Set conn = Nothing

    Response.Redirect "default.asp"
Else
%>
    <p>Error: Invalid request</p>
<%
End If
%>
