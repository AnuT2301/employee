# VB Script, ASP Classic, and SQL

### Usage: 
- Seed Database: 
```
cd ../vb-script-sql
sqlite3 < init.sql
```
- Start via IIS 

### Issues Encountered: 
- IIS Startup Issues: Starting up the web server using IIS has been challenging.
- #### Resolution: 
    - Install ODBC Driver:
        - http://www.ch-werner.de/sqliteodbc/sqliteodbc_w64.exe 
    - Working DSN connection string:  
        - `conn.Open "DSN=SQLite3 Datasource;Database=C:/path/to/vb-script-sql/company.db;"`

### Notes: 
Check System DSN:
    - `run | control admintools`
    - ODBC Data Sources (64-bit) 
    - System DSN
