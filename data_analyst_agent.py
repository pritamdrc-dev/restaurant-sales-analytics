
import sqlite3
import pandas as pd

DB_PATH = "/storage/emulated/0/Download/Restaurant-Sales-Analytics/restaurant_analytics.db"

def main():
    try:
        conn = sqlite3.connect(DB_PATH)

        tables = pd.read_sql_query(
            "SELECT name FROM sqlite_master WHERE type='table'",
            conn
        )["name"].tolist()

        if not tables:
            print("No tables found in database.")
            return

        print("\nAvailable tables:", tables)

        table = "sales" if "sales" in tables else tables[0]
        columns = pd.read_sql_query(
            f'PRAGMA table_info("{table}")', conn
        )["name"].tolist()

        print("Using table:", table)
        print("Columns:", columns)

        if "Net_Sales" in columns:
            revenue = '"Net_Sales"'
        elif all(c in columns for c in
                 ["Quantity", "Unit_Price", "Discount"]):
            revenue = '"Quantity" * "Unit_Price" - "Discount"'
        else:
            print("Cannot identify the revenue calculation.")
            return

        queries = {
            "total sales":
                f"SELECT ROUND(SUM({revenue}), 2) AS total_sales FROM \"{table}\"",

            "best branch":
                f"""SELECT Branch, ROUND(SUM({revenue}), 2) AS revenue
                    FROM "{table}" GROUP BY Branch
                    ORDER BY revenue DESC LIMIT 1""",

            "top product":
                f"""SELECT Product, ROUND(SUM({revenue}), 2) AS revenue
                    FROM "{table}" GROUP BY Product
                    ORDER BY revenue DESC LIMIT 1""",

            "monthly revenue":
                f"""SELECT strftime('%Y-%m', Order_Date) AS month,
                    ROUND(SUM({revenue}), 2) AS revenue
                    FROM "{table}" GROUP BY month ORDER BY month""",

            "payment analysis":
                f"""SELECT Payment_Method,
                    COUNT(*) AS orders,
                    ROUND(SUM({revenue}), 2) AS revenue
                    FROM "{table}" GROUP BY Payment_Method
                    ORDER BY revenue DESC"""
        }

        print("\nRESTAURANT DATA ANALYST AGENT")
        print("Ask: total sales, best branch, top product,")
        print("     monthly revenue, payment analysis")
        print("Type help or exit.")

        while True:
            question = input("\nAsk your analyst: ").strip().lower()

            if question == "exit":
                break

            if question == "help":
                print("\n".join(queries.keys()))
                continue

            if question not in queries:
                print("Unknown question. Type help.")
                continue

            result = pd.read_sql_query(queries[question], conn)
            print(result.to_string(index=False))

        conn.close()

    except Exception as error:
        print("Error:", error)
        print("Check the database path, table and column names.")

if __name__ == "__main__":
    main()
      
