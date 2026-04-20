import pandas as pd

# 1. Load data
df = pd.read_csv("online_retail.csv")

# 2. Quick check
print("Shape:", df.shape)
print(df.head())
print(df.info())

# 3. Drop unnecessary column
if "index" in df.columns:
    df = df.drop(columns=["index"])

# 4. Convert date column
df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="coerce")

# 5. Remove missing important values
df = df.dropna(subset=["InvoiceDate", "CustomerID", "Description"])

# 6. Remove invalid business rows
df = df[df["Quantity"] > 0]
df = df[df["UnitPrice"] > 0]

# 7. Convert CustomerID to integer
df["CustomerID"] = df["CustomerID"].astype(int)

# 8. Create new columns
df["TotalPrice"] = df["Quantity"] * df["UnitPrice"]
df["Month"] = df["InvoiceDate"].dt.to_period("M").astype(str)

# 9. KPI calculations
total_revenue = df["TotalPrice"].sum()
total_orders = df["InvoiceNo"].nunique()
unique_customers = df["CustomerID"].nunique()
avg_order_value = total_revenue / total_orders

print("\n--- KPI ---")
print("Total Revenue:", round(total_revenue, 2))
print("Total Orders:", total_orders)
print("Unique Customers:", unique_customers)
print("Average Order Value:", round(avg_order_value, 2))

# 10. Monthly summary
monthly_summary = (
    df.groupby("Month", as_index=False)
    .agg(
        Revenue=("TotalPrice", "sum"),
        Orders=("InvoiceNo", "nunique"),
        Customers=("CustomerID", "nunique")
    )
    .sort_values("Month")
)

monthly_summary["Growth_%"] = monthly_summary["Revenue"].pct_change()

# 11. Country summary
country_summary = (
    df.groupby("Country", as_index=False)
    .agg(
        Revenue=("TotalPrice", "sum"),
        Orders=("InvoiceNo", "nunique"),
        Customers=("CustomerID", "nunique")
    )
    .sort_values("Revenue", ascending=False)
)

# 12. Top products
product_summary = (
    df.groupby("Description", as_index=False)
    .agg(
        Revenue=("TotalPrice", "sum"),
        Quantity_Sold=("Quantity", "sum")
    )
    .sort_values("Revenue", ascending=False)
    .head(10)
)

# 13. Top customers
customer_summary = (
    df.groupby("CustomerID", as_index=False)
    .agg(
        Revenue=("TotalPrice", "sum"),
        Orders=("InvoiceNo", "nunique")
    )
    .sort_values("Revenue", ascending=False)
    .head(10)
)

# 14. Export sample cleaned data + summaries to Excel
cleaned_sample = df.head(5000).copy()

with pd.ExcelWriter("retail_report.xlsx", engine="openpyxl") as writer:
    cleaned_sample.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    monthly_summary.to_excel(writer, sheet_name="Monthly_Summary", index=False)
    country_summary.to_excel(writer, sheet_name="Country_Summary", index=False)
    product_summary.to_excel(writer, sheet_name="Top_Products", index=False)
    customer_summary.to_excel(writer, sheet_name="Top_Customers", index=False)

print("\n Excel report created successfully: retail_report.xlsx")