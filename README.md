# CRUD Application: ASP Classic and SQLite Project 
## Overview

This project is a basic web application that demonstrates the integration of VBScript with SQL. It includes ASP pages for CRUD operations and a SQLite database.

## Demo: 

![demo](https://github.com/neon-nomad/vb-script-sql/assets/58483497/95f5d171-d2fd-45c8-9c46-eb04fa36997a)


## File Structure

- **add.asp**: Handles adding new entries to the database.
- **company.db**: SQLite database file.
- **default.asp**: Main page displaying the data.
- **delete.asp**: Handles deletion of entries.
- **edit.asp**: Handles editing existing entries.
- **init.sql**: SQL script to initialize the database schema.
- **README.md**: Basic readme file.
- **styles.css**: Stylesheet for the web application.
- **web.config**: Configuration file for the web application.

## Getting Started

1. **Setup Environment**:
   - Ensure you have a web server supporting ASP (e.g., IIS).
   - SQLite should be installed if not included in your environment.

2. **Database Initialization**:
   - Run `init.sql` to set up the initial database schema. This script creates the necessary tables.

3. **Deploy Application**:
   - Place the files in the web root directory of your server.
   - Ensure the database file `company.db` is writable by the web server process.

## CRUD Operations

- **Add (add.asp)**: Form to input new data. Submits to `add.asp` for processing.
- **Edit (edit.asp)**: Form pre-filled with existing data. Submits to `edit.asp` for updating the entry.
- **Delete (delete.asp)**: Processes deletion of entries.
- **Display (default.asp)**: Lists all data entries from the database.

## Configuration

- **web.config**: Contains necessary configurations for running the ASP application. Ensure this file is correctly set up to handle your environment specifics.

## Styles and Presentation

- **styles.css**: Customize the appearance of the web application. Modify this file to change the look and feel.

## Notes
- ODBC Driver Installation:
    - http://www.ch-werner.de/sqliteodbc/sqliteodbc_w64.exe 
    - Working DSN connection string:  
        - `conn.Open "DSN=SQLite3 Datasource;Database=C:/path/to/vb-script-sql/company.db;"`
 
- Check System DSN:
    - `run | control admintools`
    - ODBC Data Sources (64-bit) 
    - System DSN
