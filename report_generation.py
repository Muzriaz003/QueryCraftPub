import pandas as pd
from sqlalchemy import create_engine, text
from formatting import format_worksheet
import time
import re


def parse_query_blocks(query):
    parts = re.split(r"--\s*sheet:\s*(.+)", query)
    blocks = []
    for i in range(1, len(parts), 2):
        sheet_name = parts[i].strip()
        sql_block = parts[i + 1].strip()
        queries = [q.strip() for q in sql_block.split(";") if q.strip()]
        blocks.append((sheet_name, queries))
    return blocks


def generate_dynamic_reports(engine, writer, sheet_query_blocks, chunksize=100000):
    #start time for debugging
    start_time = time.time()
    #engine for SSMS Connection
    with engine.connect() as conn:
        for sheet_name, queries in sheet_query_blocks:
            print(f"[{round(time.time() - start_time, 2)}s] Processing sheet: {sheet_name}")
            offset = 0
            first_chunk = True
            df_chunks = []

            for query in queries: 
                try:
                    df = pd.read_sql(text(query), conn)
                    df.fillna('NULL')
                    df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False,
                        header=first_chunk,
                        startrow=offset
                    )
                    print(f"[{round(time.time() - start_time, 2)}s] Wrote query to {sheet_name} (rows: {len(df)})")
                    df_chunks.append(df)
                    offset += len(df) + (1 if first_chunk else 0)
                    first_chunk = False
                except Exception as e:
                    print(f"Error in query for sheet {sheet_name}: {e}")
                    continue
        
        
            # Combine chunks into full DataFrame)
            try:
                if df_chunks: 
                  df_all = pd.concat(df_chunks, ignore_index=True)
                  print(f"[{round(time.time() - start_time, 2)}s] Formatting entire sheet: {sheet_name}")
                  format_worksheet(writer, df_all, sheet_name)
                  print(f"[{round(time.time() - start_time, 2)}s] Formatted: {sheet_name}")
            except Exception as e:
                print(f"Error while formatting {sheet_name}: {e}")

    print(f"[{round(time.time() - start_time, 2)}s] All tables processed.")

def main(query, file_path, connection_string):
    overall_start = time.time()
    engine = create_engine(connection_string)
    sheet_query_blocks = parse_query_blocks(query)


    if not sheet_query_blocks:
        print("No sheets detected, Use -- sheet: SheetName in Your Query Window")
        return

    with pd.ExcelWriter(
        file_path,
        engine='xlsxwriter',
    ) as writer:
        generate_dynamic_reports(
            engine,
            writer,
            sheet_query_blocks,
            chunksize=100000
        )

    print(f"Excel file '{file_path}' created with the selected tables in separate sheets.")
    print(f"[{round(time.time() - overall_start, 2)}s] Excel file '{file_path}' created successfully.")
