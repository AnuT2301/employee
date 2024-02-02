# VB Script, ASP Classic, and SQL

### Technologies Used
- VB Script: A scripting language primarily used for client-side web development.
- ASP Classic: Active Server Pages (ASP) Classic is a server-side scripting environment for building dynamic web applications.
- SQL: Structured Query Language is used for managing and manipulating relational databases.


Issues Encountered: 
- IIS Startup Issues: Starting up the web server using IIS has been challenging.
    - Install ODBC Driver:
        - http://www.ch-werner.de/sqliteodbc/sqliteodbc_w64.exe 
    - Working DSN connection string:  
        - `conn.Open "DSN=SQLite3 Datasource;Database=C:/src/vb-script-sql/company.db;"`