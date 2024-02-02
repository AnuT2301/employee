<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=SQLite3 Datasource;Database=C:/src/vb-script-sql/company.db;"

If Request("id") <> "" Then
    conn.Execute("DELETE FROM Employees WHERE ID = " & Request("id"))
End If

Response.Redirect("default.asp")

conn.Close
%>
