<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DRIVER=SQLite3;Database=C:/src/vb-script-sql/company.db;"

If Request("id") <> "" Then
    conn.Execute("DELETE FROM Employees WHERE ID = " & Request("id"))
End If

Response.Redirect("index.asp")

conn.Close
%>
