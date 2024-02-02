<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DRIVER=SQLite3;Database=C:/src/vb-script-sql/company.db;"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM Employees", conn

Do Until rs.EOF
    Response.Write("ID: " & rs("ID") & "<br>")
    Response.Write("Name: " & rs("Name") & "<br>")
    Response.Write("Position: " & rs("Position") & "<br>")
    Response.Write("<a href='edit.asp?id=" & rs("ID") & "'>Edit</a><br>")
    Response.Write("<a href='delete.asp?id=" & rs("ID") & "'>Delete</a><br>")
    rs.MoveNext
Loop

rs.Close
conn.Close
%>
