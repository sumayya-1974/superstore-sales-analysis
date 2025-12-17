import pandas as pd
import sqlite3
import matplotlib.pyplot as plt

# File paths
csv_path = r"E:\superstore sales dataset\train.csv"
db_path = r"E:\superstore sales dataset\superstore.db"

# Load CSV
df = pd.read_csv(csv_path, encoding='utf-8')

print("First 5 rows of the dataset:")
print(df.head())

# Connect to SQLite
conn = sqlite3.connect(db_path)

# --- Initial Save to DB ---
df.to_sql("sales", conn, if_exists="replace", index=False)
print("\nDatabase 'superstore.db' created and 'sales' table loaded successfully!")

# Column info
print("\nColumn names and datatypes:")
print(df.dtypes)

# Missing Values
print("\nMissing values in each column:")
print(df.isnull().sum())

# --- DATA CLEANING ---
df['Postal Code'] = df['Postal Code'].fillna(0)

df['Order Date'] = pd.to_datetime(df['Order Date'], dayfirst=True)
df['Ship Date'] = pd.to_datetime(df['Ship Date'], dayfirst=True)

print(df['Order Date'].head())
print(df['Ship Date'].head())

df['Delivery_Days'] = (df['Ship Date'] - df['Order Date']).dt.days
print(df[['Order Date', 'Ship Date', 'Delivery_Days']].head())

print("\nMissing values in each column:")
print(df.isnull().sum())

print("Average delivery days:", round(df['Delivery_Days'].mean()))
print("Minimum delivery days:", df['Delivery_Days'].min())
print("Maximum delivery days:", df['Delivery_Days'].max())

print(df[df['Delivery_Days'] == 0][['Ship Mode', 'Order Date', 'Ship Date']].head())

# --- VERY IMPORTANT: UPDATE CLEAN DATA BACK INTO DB ---
df.to_sql("sales", conn, if_exists="replace", index=False)

# --- CATEGORY SALES ---
query = """
SELECT Category, SUM(Sales) AS TotalSales
FROM sales
GROUP BY Category;
"""

result = pd.read_sql_query(query, conn)
print(result)
# Save category table
result.to_excel(r"E:\superstore sales dataset\result.xlsx", index=False)
# Pie chart
plt.figure(figsize=(6,6))
plt.pie(result['TotalSales'], labels=result['Category'], autopct='%1.1f%%')
plt.title("Sales Share by Category")
plt.savefig(r"E:\superstore sales dataset\category_sales_plot.png")
plt.close()

# --- SUB-CATEGORY SALES ---
query_sub = """
SELECT [Sub-Category], SUM(Sales) AS TotalSales
FROM sales
GROUP BY [Sub-Category]
ORDER BY TotalSales DESC
"""
sub_category_sales = pd.read_sql_query(query_sub, conn)
print(sub_category_sales)
sub_category_sales.to_excel(r"E:\superstore sales dataset\result.xlsx", index=False)

plt.figure(figsize=(12,6))
plt.bar(sub_category_sales['Sub-Category'], sub_category_sales['TotalSales'])
plt.title("Total Sales by Sub-Category")
plt.xlabel("Sub-Category")
plt.ylabel("Total Sales")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig(r"E:\superstore sales dataset\sub_category_sales_plot.png")
plt.close()

sub_category_sales.to_excel(r"E:\superstore sales dataset\sub_category_sales.xlsx", index=False)
# Query: Total Sales by Region
query_region = """
SELECT Region, SUM(Sales) AS TotalSales
FROM sales
GROUP BY Region
ORDER BY TotalSales DESC
"""
region_sales = pd.read_sql_query(query_region, conn)
# Display table
print(region_sales)
# Save to Excel
region_sales.to_excel(r"E:\superstore sales dataset\region_sales.xlsx", index=False)
plt.figure(figsize=(8,5))
plt.bar(region_sales['Region'], region_sales['TotalSales'], color=['skyblue','orange','green','red'])
plt.title("Total Sales by Region")
plt.xlabel("Region")
plt.ylabel("Total Sales")
plt.tight_layout()
# Save chart
plt.savefig(r"E:\superstore sales dataset\region_sales.png", dpi=300)
plt.close()

#segment wise sales
query_segment = """
SELECT Segment, SUM(Sales) AS TotalSales
FROM sales
GROUP BY Segment
ORDER BY TotalSales DESC
"""
segment_sales = pd.read_sql_query(query_segment, conn)
print("\nSales by Segment:")
print(segment_sales)
segment_sales.to_excel(r"E:\superstore sales dataset\sales_by_segment.xlsx", index=False)

#plotting bar graph
plt.figure(figsize=(8,5))
plt.bar(segment_sales['Segment'], segment_sales['TotalSales'])
plt.title("Total Sales by Segment")
plt.xlabel("Segment")
plt.ylabel("Total Sales")
plt.tight_layout()
plt.savefig(r"E:\superstore sales dataset\sales_by_segment.png")   # SAVE PLOT
plt.close()
#sales by shipmode
query_ship = """
SELECT [Ship Mode], SUM(Sales) AS TotalSales
FROM sales
GROUP BY [Ship Mode]
ORDER BY TotalSales DESC
"""
ship_sales = pd.read_sql_query(query_ship, conn)
print("\nSales by Ship Mode:")
print(ship_sales)
# Save Ship Mode sales table
ship_sales.to_excel(r"E:\superstore sales dataset\ship_sales.xlsx", index=False)
#plotting bar graph
plt.figure(figsize=(8,5))
plt.bar(ship_sales['Ship Mode'], ship_sales['TotalSales'])
plt.title("Total Sales by Ship Mode")
plt.xlabel("Ship Mode")
plt.ylabel("Total Sales")
plt.tight_layout()
plt.savefig(r"E:\superstore sales dataset\ship_sales..png",dpi=300)
plt.close()
#statewise sales
query_state = """
SELECT State, SUM(Sales) AS TotalSales
FROM sales
GROUP BY State
ORDER BY TotalSales DESC
"""
state_sales = pd.read_sql_query(query_state, conn)
print("\nTop States by Sales:")
print(state_sales.head(10))
state_sales.to_excel(r"E:\superstore sales dataset\state_sales.xlsx", index=False)
#bar chart
plt.figure(figsize=(10,6))
plt.bar(state_sales.head(10)['State'], state_sales.head(10)['TotalSales'])
plt.title("Top 10 States by Total Sales")
plt.xlabel("State")
plt.ylabel("Total Sales")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig(r"E:\superstore sales dataset\state_sales.png", dpi=300)
plt.close()
#%%
query_delivery = """
SELECT [Ship Mode],
       ROUND(AVG(Delivery_Days), 2) AS Avg_Delivery_Days
FROM sales
GROUP BY [Ship Mode]
ORDER BY Avg_Delivery_Days
"""
delivery_by_ship = pd.read_sql_query(query_delivery, conn)

print("\nAverage Delivery Days by Ship Mode:")
print(delivery_by_ship)
delivery_by_ship.to_excel(
    r"E:\superstore sales dataset\delivery_days_by_ship_mode.xlsx",
    index=False
)
plt.figure(figsize=(8,5))
plt.bar(delivery_by_ship['Ship Mode'],
        delivery_by_ship['Avg_Delivery_Days'])

plt.title("Average Delivery Days by Ship Mode")
plt.xlabel("Ship Mode")
plt.ylabel("Average Delivery Days")
plt.tight_layout()
plt.savefig(
    r"E:\superstore sales dataset\delivery_days_by_ship_mode.png",
    dpi=300)
plt.close()
# Create delivery speed categories
df['Delivery_Speed'] = pd.cut(
    df['Delivery_Days'],
    bins=[-1, 2, 5, 10, 100],
    labels=['Fast (0–2 days)', 'Medium (3–5 days)', 'Slow (6–10 days)', 'Very Slow (10+ days)']
)

print(df[['Delivery_Days', 'Delivery_Speed']].head())
df.to_sql("sales", conn, if_exists="replace", index=False)
print("Updated sales table with Delivery_Speed column")
query_speed_sales = """
SELECT [Delivery_Speed],
       ROUND(SUM(Sales), 2) AS TotalSales,
       ROUND(AVG(Sales), 2) AS AvgSales
FROM sales
GROUP BY [Delivery_Speed]
ORDER BY TotalSales DESC
"""
speed_sales = pd.read_sql_query(query_speed_sales, conn)
print("\nSales vs Delivery Speed:")
print(speed_sales)
speed_sales.to_excel(
    r"E:\superstore sales dataset\sales_vs_delivery_speed.xlsx",
    index=False
)

plt.figure(figsize=(9,5))
plt.bar(speed_sales['Delivery_Speed'], speed_sales['TotalSales'])
plt.title("Total Sales by Delivery Speed")
plt.xlabel("Delivery Speed")
plt.ylabel("Total Sales")
plt.xticks(rotation=20)
plt.tight_layout()
plt.savefig(
    r"E:\superstore sales dataset\sales_vs_delivery_speed.png",
    dpi=300
)
plt.close()
# Extract Year and Month from Order Date
df['Order_Year'] = df['Order Date'].dt.year
df['Order_Month'] = df['Order Date'].dt.month
df['Order_Month_Name'] = df['Order Date'].dt.month_name()

print(df[['Order Date', 'Order_Year', 'Order_Month_Name']].head())
#update sql table for newly added columns
df.to_sql("sales", conn, if_exists="replace", index=False)
print("Sales table updated with Year and Month columns")
#yearly sales trend
query_yearly = """
SELECT Order_Year, ROUND(SUM(Sales), 2) AS TotalSales
FROM sales
GROUP BY Order_Year
ORDER BY Order_Year
"""
yearly_sales = pd.read_sql_query(query_yearly, conn)

print("\nYearly Sales Trend:")
print(yearly_sales)
yearly_sales.to_excel(r"E:\superstore sales dataset\yearly_sales_trend.xlsx", index=False)
#visualisation 
plt.figure(figsize=(8,5))
plt.plot(yearly_sales['Order_Year'], yearly_sales['TotalSales'], marker='o')
plt.title("Yearly Sales Trend")
plt.xlabel("Year")
plt.ylabel("Total Sales")
plt.grid(True)
plt.tight_layout()
plt.savefig(r"E:\superstore sales dataset\yearly_sales_trend.png", dpi=300)
plt.close()
#monthly sales trend
query_monthly = """
SELECT Order_Month_Name,
       ROUND(SUM(Sales), 2) AS TotalSales
FROM sales
GROUP BY Order_Month, Order_Month_Name
ORDER BY Order_Month
"""
monthly_sales = pd.read_sql_query(query_monthly, conn)
print("\nMonthly Sales Trend:")
print(monthly_sales)
monthly_sales.to_excel(r"E:\superstore sales dataset\monthly_sales_trend.xlsx", index=False)
#visualisation
plt.figure(figsize=(10,5))
plt.plot(monthly_sales['Order_Month_Name'], monthly_sales['TotalSales'], marker='o')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()
plt.savefig(r"E:\superstore sales dataset\monthly_sales_trend.png", dpi=300)
plt.close()
#top 10 products
query_top_products = """
SELECT [Product Name],
       ROUND(SUM(Sales), 2) AS TotalSales
FROM sales
GROUP BY [Product Name]
ORDER BY TotalSales DESC
LIMIT 10
"""
top_products = pd.read_sql_query(query_top_products, conn)
print("\nTop 10 Products by Sales:")
print(top_products)
top_products.to_excel(
    r"E:\superstore sales dataset\top_10_products.xlsx",
    index=False)
#visualisation
plt.figure(figsize=(10,6))
plt.barh(top_products['Product Name'], top_products['TotalSales'])
plt.title("Top 10 Products by Total Sales")
plt.xlabel("Total Sales")
plt.ylabel("Product Name")
plt.gca().invert_yaxis()  # Highest sales on top
plt.tight_layout()
plt.savefig(
    r"E:\superstore sales dataset\top_10_products.png",
    dpi=300)
plt.close()

















# %%
