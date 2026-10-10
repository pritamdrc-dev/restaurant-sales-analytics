
import os
import sqlite3
import pandas as pd

BASE = (
    "/storage/emulated/0/Download/"
    "Restaurant-Sales-Analytics/"
)

DB = os.path.join(BASE, "restaurant_analytics.db")
OUTPUT = os.path.join(BASE, "Restaurant_Analytics_Report.xlsx")

REPORTS = {
    "Branch Performance": """
        SELECT Branch,
               COUNT(DISTINCT Order_ID) AS Orders,
               SUM(Net_Sales) AS Revenue
        FROM sales
        GROUP BY Branch ORDER BY Revenue DESC
    """,
    "Product Performance": """
        SELECT Product,
               SUM(Quantity) AS Units_Sold,
               SUM(Net_Sales) AS Revenue
        FROM sales
        GROUP BY Product ORDER BY Revenue DESC
    """,
    "Monthly Revenue": """
        SELECT strftime('%Y-%m', Order_Date) AS Month,
               SUM(Net_Sales) AS Revenue
        FROM sales
        GROUP BY Month ORDER BY Month
    """,
    "Payment Analysis": """
        SELECT Payment_Method,
               COUNT(DISTINCT Order_ID) AS Orders,
               SUM(Net_Sales) AS Revenue
        FROM sales
        GROUP BY Payment_Method ORDER BY Revenue DESC
    """,
    "Category Analysis": """
        SELECT Category,
               SUM(Quantity) AS Units_Sold,
               SUM(Net_Sales) AS Revenue
        FROM sales
        GROUP BY Category ORDER BY Revenue DESC
    """
}

def main():
    if not os.path.exists(DB):
        print("Database not found:", DB)
        return

    with sqlite3.connect(DB) as conn:
        summary = pd.read_sql_query("""
            SELECT
                SUM(Net_Sales) AS Total_Net_Sales,
                COUNT(DISTINCT Order_ID) AS Total_Orders,
                ROUND(
                    SUM(Net_Sales) * 1.0 /
                    COUNT(DISTINCT Order_ID), 2
                ) AS Average_Order_Value
            FROM sales
        """, conn)

        best_branch = pd.read_sql_query("""
            SELECT Branch, SUM(Net_Sales) AS Revenue
            FROM sales
            GROUP BY Branch
            ORDER BY Revenue DESC LIMIT 1
        """, conn)

        top_product = pd.read_sql_query("""
            SELECT Product, SUM(Net_Sales) AS Revenue
            FROM sales
            GROUP BY Product
            ORDER BY Revenue DESC LIMIT 1
        """, conn)

        with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
            summary.to_excel(
                writer, sheet_name="Executive Summary", index=False
            )
            best_branch.to_excel(
                writer, sheet_name="Best Branch", index=False
            )
            top_product.to_excel(
                writer, sheet_name="Top Product", index=False
            )

            for sheet, query in REPORTS.items():
                df = pd.read_sql_query(query, conn)
                df.to_excel(
                    writer, sheet_name=sheet, index=False
                )

            clean_data = pd.read_sql_query(
                "SELECT * FROM sales", conn
            )
            clean_data.to_excel(
                writer, sheet_name="Sales Data", index=False
            )

    print("EXCEL REPORT CREATED SUCCESSFULLY!")
    print("Saved to:", OUTPUT)
    print("Open the file from your Download folder.")

if __name__ == "__main__":
    main()
