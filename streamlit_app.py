import streamlit as st
import re
import io
from report_generation import main  
from config import get_connection_string
from sqlalchemy import create_engine

#input sidebar ui for MSM server connection
server = st.sidebar.text_input("Server", value="localhost\\SQLEXPRESS")
database = st.sidebar.text_input("Database", value="AdventureWorks2019")
trusted = st.sidebar.checkbox("Use Windows Auth", value=True)

#if no Windows Auth then use username and password for sql server connection
username = password = None
if not trusted:
    username = st.sidebar.text_input("SQL Username")
    password = st.sidebar.text_input("SQL Password", type="password")

# Build the connection string
try:
    connection_string = get_connection_string(server, database, trusted=trusted, username=username, password=password)
    engine = create_engine(connection_string)
    st.success("Connected to database successfully.")
except Exception as e:
    st.error(f"Connection failed: {e}")
    st.stop()

st.title("SQL Report Generator")

st.markdown("""
            write your SQL like this:
            -- sheet: Summary
SELECT 'Summary' AS Info;
SELECT COUNT(*) AS Total FROM Sales;

-- sheet: Detail
SELECT * FROM Sales;
""")

query = st.text_area("Enter your SQL with --sheet: Labels", height=200)
file_name = st.text_input("Enter desired file name (without .xlsx):")

if st.button("Generate Report"):
    if not query.strip():
        st.warning("Please enter a SQL query.")
    elif not file_name.strip():
        st.warning("Please enter a valid file name.")
    else:
        try:
            st.info("Generating report... please wait.")
            output = f"{file_name}.xlsx"
            main(query, output, connection_string)

            with open(output, "rb") as f:
                report_bytes = f.read()
            st.success(f"Report generated as '{output}'")
            st.download_button("Download Excel Report", data=report_bytes, file_name=output)
        except Exception as e:
            st.error(f"Error generating report: {e}")

