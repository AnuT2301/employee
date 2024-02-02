<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DRIVER=SQLite3;Database=C:/src/vb-script-sql/company.db;"

If Request("id") <> "" Then
    If Request("Name") <> "" And Request("Position") <> "" Then
        conn.Execute("UPDATE Employees SET Name = '" & Request("Name") & "', Position = '" & Request("Position") & "' WHERE ID = " & Request("id"))
        Response.Redirect("index.asp")
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM Employees WHERE ID = " & Request("id"), conn
        If Not rs.EOF Then
            Response.Write("<form method='post'>")
            Response.Write("Name: <input type='text' name='Name' value='" & rs("Name") & "'><br>")
            Response.Write("Position: <input type='text' name='Position' value='" & rs("Position") & "'><br>")
            Response.Write("<input type='submit' value='Update'>")
            Response.Write("</form>")
        End If
        rs.Close
    End If
End If

conn.Close
%>
