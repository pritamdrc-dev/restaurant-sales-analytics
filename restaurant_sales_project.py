
import pandas as pd
import random
from datetime import date, timedelta

random.seed(42)

products = {
    "Pizza": ("Main Course", 299),
    "Burger": ("Main Course", 149),
    "Pasta": ("Main Course", 229),
    "Garlic Bread": ("Sides", 119),
    "Cold Coffee": ("Beverages", 99),
    "Brownie": ("Dessert", 129)
}

branches = ["Kolkata", "Dankuni", "Howrah", "Salt Lake"]
payments = ["UPI", "Cash", "Card"]

rows = []
start_date = date(2026, 1, 1)

for i in range(1, 121):
    product = random.choice(list(products.keys()))
    category, price = products[product]
    quantity = random.randint(1, 5)
    discount = random.choice([0, 0, 0, 20, 50])

    order_date = start_date + timedelta(
        days=random.randint(0, 89)
    )

    rows.append({
        "Order_ID": f"ORD{i:04d}",
        "Order_Date": order_date.isoformat(),
        "Branch": random.choice(branches),
        "Product": product,
        "Category": category,
        "Quantity": quantity,
        "Unit_Price": price,
        "Discount": discount,
        "Payment_Method": random.choice(payments)
    })

df = pd.DataFrame(rows)

# Introduce two data-quality issues for practice
df.loc[5, "Quantity"] = None
df = pd.concat([df, df.iloc[[0]]], ignore_index=True)

path = "/storage/emulated/0/Download/restaurant_sales_raw.csv"
df.to_csv(path, index=False)

print("RESTAURANT SALES DATASET CREATED!")
print("Total rows:", len(df))
print("Columns:", list(df.columns))
print("Saved to:", path)
print("\nFirst 5 rows:")
print(df.head().to_string(index=False))
