📊 Excel to Oracle SQL Migrator
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Python-based utility that automates the transformation of multiple .xlsx files into SQL tables within an Oracle database. It intelligently detects column names and data types, generates CREATE TABLE statements, and populates the tables using INSERT INTO.

⚙️ Features
- Scans multiple Excel files from the Organizador/ folder
- Infers data types from sample rows
- Dynamically creates Oracle SQL tables
- Inserts real records safely, escaping special characters
- Converts .csv files to .xlsx format (optional step)
🚀 Requirements
- Python 3.9+
- Oracle Instant Client (e.g., instantclient_23_8)
- Access to a local or remote Oracle database
- Dependencies:
    pandas
    openpyxl
    oracledb


Install with:
pip install pandas openpyxl oracledb


📁 Expected Structure
Each Excel file should start with a header row, followed by actual data:
| name | email | age | 
| Leila Howe | leila@correo.com | 28 | 
| Kevin Smith | kevin@ejemplo.cl | 34 | 


All files must be placed in the Organizador/ folder.
🔧 Setup
Edit your database connection details in the script:
connection = oracledb.connect(
    user="YOUR_USER",
    password="YOUR_PASSWORD",
    dsn="localhost:1521/XE"  # Update based on your Oracle config
)


Also set the correct path to your Oracle Instant Client:
oracledb.init_oracle_client(lib_dir="C:/path/to/instantclient")


🛠 How to Run
- Place .xlsx files inside Organizador/ if you have another format put it on Excels (only supports csv)
- Execute the script:
    python main.py

- The console will show each CREATE TABLE execution and inserted records
✅ Data Safety
The script escapes single quotes (') inside text fields to avoid SQL syntax errors. If your files contain complex types like dates, booleans, or formatted currency, you can extend the generate_insert() function accordingly.
