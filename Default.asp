<%
Dim sort
sort = Request.QueryString("sort")

If sort = "" Then
  sort = "ID" ' Default sort column
End If

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=SQLite3 Datasource;Database=C:/src/vb-script-sql/company.db;"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM Employees ORDER BY " & sort, conn
%>

<style>
body {
  font-family: Arial, sans-serif;
}

form {
  margin-bottom: 1em;
}

label {
  display: block;
  margin-top: 1em;
}

input[type="text"] {
  width: 100%;
  padding: 0.5em;
  margin: 0.5em 0;
  box-sizing: border-box;
}

input[type="submit"] {
  background-color: #4CAF50;
  color: white;
  padding: 0.5em 1em;
  margin: 1em 0;
  border: none;
  cursor: pointer;
}

input[type="submit"]:hover {
  background-color: #45a049;
}

table {
  width: 100%;
  border-collapse: collapse;
}

th, td {
  text-align: left;
  padding: 0.5em;
  border-bottom: 1px solid #ddd;
}

tr:hover {background-color: #f5f5f5;}

.container {
  width: 50%; 
  margin: auto; 
  padding: 1em;
  border: 1px solid #ddd;
  border-radius: 5px;
}

input[type="text"] {
  width: 100%; 
}
</style>

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
