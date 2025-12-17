from xlsxwriter.utility import xl_rowcol_to_cell

# Optimized format_worksheet using Excel table feature

def format_worksheet(writer, df, sheet_name):
    """
    Formats an entire sheet by converting it into an Excel table and auto-sizing columns.
    Args:
        writer: pandas ExcelWriter with xlsxwriter engine
        df: pandas DataFrame with all data for the sheet
        sheet_name: name of the sheet to format
    """
    # Access workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Number of rows and columns
    n_rows = len(df) + 1  # +1 for header
    n_cols = len(df.columns)

    # Auto-fit columns once based on header and data
    for i, column in enumerate(df.columns):
        # Compute maximum length among header and column data
        max_len = max(
            len(str(column)),
            df[column].astype(str).map(len).max()
        )
        worksheet.set_column(i, i, max_len + 2)

    # Define Excel table range (A1 to last cell)
    first_cell = xl_rowcol_to_cell(0, 0)
    last_cell = xl_rowcol_to_cell(n_rows, n_cols - 1)

    # Add Excel table for built-in styling, filters, and banded rows
    worksheet.add_table(f"{first_cell}:{last_cell}", {
        'columns': [{'header': col} for col in df.columns],
        'style': 'None'
    })
