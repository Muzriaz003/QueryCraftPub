QueryCraft: SQL to Excel Report Generator
QueryCraft is a Streamlit-based web application that allows users to execute SQL queries against a Microsoft SQL Server (MSSQL) database and export the results into a professionally formatted Excel workbook. It supports multi-sheet generation using simple SQL comments.

✨ Features
Custom Multi-Sheet Support: Use -- sheet: SheetName comments in your SQL to automatically split results into different Excel tabs.

Professional Formatting: Automatically converts query results into Excel tables with auto-fitted columns and banded rows for readability.

Flexible Authentication: Supports both Windows Authentication (Trusted Connection) and standard SQL Server Authentication.

Dynamic UI: Built with Streamlit for a lightweight, interactive user experience.

Efficient Processing: Handles multiple queries per sheet and combines them into a single continuous report.

🛠️ Tech Stack
Frontend: Streamlit

Database Connection: SQLAlchemy & pyodbc

Data Handling: Pandas

Excel Export: XlsxWriter

🚀 Getting Started
Prerequisites
Python 3.8+

Microsoft ODBC Driver for SQL Server

Installation
Clone the repository:

Bash

git clone https://github.com/Muzriaz003/QueryCraftPub.git
cd QueryCraftPub
Install dependencies:

Bash

pip install streamlit pandas sqlalchemy pyodbc xlsxwriter
Running the App
Bash

streamlit run streamlit_app.py
📝 How to Use
Enter your Server and Database details in the sidebar.

In the query editor, define your sheets using the following syntax:

SQL

-- sheet: Summary
SELECT COUNT(*) as TotalSales FROM Sales.SalesOrderHeader;

-- sheet: Details
SELECT TOP 100 * FROM Sales.SalesOrderHeader;
Enter a filename and click Generate Report.

Download your formatted .xlsx file!.
