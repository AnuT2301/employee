<%
Dim sort
sort = Request.QueryString("sort")

If sort = "" Then
  sort = "ID" 
End If

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=SQLite3 Datasource;Database=C:/src/vb-script-sql/company.db;"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM Employees ORDER BY " & sort, conn
%>

<head>
  <link rel="stylesheet" type="text/css" href="styles.css">
</head>

<div class="container">
  <form action="add.asp" method="post">
    <label for="name">Name:</label>
    <input type="text" id="name" name="name"><br><br>
    <label for="position">Position:</label>
    <input type="text" id="position" name="position"><br><br>
    <input type="submit" value="Add Employee">
  </form>
</div>

<table>
  <tr>
    <th><a href="?sort=ID">ID</a></th>
    <th><a href="?sort=Name">Name</a></th>
    <th><a href="?sort=Position">Position</a></th>
    <th>Edit</th>
    <th>Delete</th>
  </tr>

<%
Do Until rs.EOF
%>
  <tr>
    <td><%=rs("ID")%></td>
    <td><%=rs("Name")%></td>
    <td><%=rs("Position")%></td>
    <td><a href='edit.asp?id=<%=rs("ID")%>'>Edit</a></td>
    <td><a href='delete.asp?id=<%=rs("ID")%>'>Delete</a></td>
  </tr>
<%
  rs.MoveNext
Loop
rs.Close
conn.Close
%>

</table>
