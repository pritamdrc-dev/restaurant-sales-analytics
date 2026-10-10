-- Restaurant Sales Analytics
-- SQLite queries

-- 1. Total net sales
SELECT ROUND(SUM(Quantity * Unit_Price - Discount), 2) AS total_net_sales
FROM sales;

-- 2. Revenue by branch
SELECT Branch,
       ROUND(SUM(Quantity * Unit_Price - Discount), 2) AS revenue
FROM sales
GROUP BY Branch
ORDER BY revenue DESC;

-- 3. Revenue by product
SELECT Product,
       SUM(Quantity) AS units_sold,
       ROUND(SUM(Quantity * Unit_Price - Discount), 2) AS revenue
FROM sales
GROUP BY Product
ORDER BY revenue DESC;

-- 4. Revenue by payment method
SELECT Payment_Method,
       COUNT(*) AS orders,
       ROUND(SUM(Quantity * Unit_Price - Discount), 2) AS revenue
FROM sales
GROUP BY Payment_Method
ORDER BY revenue DESC;

-- 5. Monthly revenue
SELECT strftime('%Y-%m', Order_Date) AS month,
       ROUND(SUM(Quantity * Unit_Price - Discount), 2) AS revenue
FROM sales
GROUP BY month
ORDER BY month;
