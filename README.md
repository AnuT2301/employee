# CRUD Application: VB Script, ASP Classic, and SQLite

## Demo: 

![demo](https://github.com/neon-nomad/vb-script-sql/assets/58483497/95f5d171-d2fd-45c8-9c46-eb04fa36997a)

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
