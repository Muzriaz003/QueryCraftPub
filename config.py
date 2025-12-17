
from urllib.parse import quote_plus


def get_connection_string(server, database="master", driver="ODBC Driver 17 for SQL server", trusted=True, username=None, password=None):
    base = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
    if trusted:
        base += "Trusted_Connection=yes;"
    else:
        base += f"UID={username};PWD={password};"

    return f"mssql+pyodbc:///?odbc_connect={quote_plus(base)}"
    

