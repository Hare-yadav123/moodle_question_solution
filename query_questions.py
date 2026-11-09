import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
print("base dir",BASE_DIR)

DB_PATH = BASE_DIR / "output" / "questions.db"
SQL_PATH = BASE_DIR / "queries.sql"

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

# Read SQL file
with open(SQL_PATH, "r", encoding="utf-8") as file:
    sql_queries = file.read()

# Split and execute individual SELECT queries
queries = [
    query.strip()
    for query in sql_queries.split(";")
    if query.strip()
]

for i, query in enumerate(queries, start=1):

    print(f"\n========== Query {i} Result ==========")

    cursor.execute(query)

    # Fetch result
    rows = cursor.fetchall()

    # Print column names
    if cursor.description:
        columns = [column[0] for column in cursor.description]
        print(columns)

        for row in rows:
            print(row)
    else:
        print("Query executed successfully.")

conn.close()