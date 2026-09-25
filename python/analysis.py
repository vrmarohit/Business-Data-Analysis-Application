# ============================================================
# BUSINESS DATA ANALYSIS APPLICATION
# Complete Analysis + Visualization + Forecasting
# + AI-Assisted Automated Business Insights
# ============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os

# ============================================================
# 1. LOAD CLEANED DATA
# ============================================================

file_path = "data/cleaned_business_sales.csv"

df = pd.read_csv(file_path)

# Convert Date column
df["Date"] = pd.to_datetime(df["Date"])

print("\n========================================")
print("BUSINESS DATA ANALYSIS APPLICATION")
print("========================================")

print("\nDataset Shape:")
print(df.shape)

print("\nFirst 5 Records:")
print(df.head())


# ============================================================
# 2. BASIC STATISTICAL ANALYSIS
# ============================================================

print("\n========================================")
print("SALES STATISTICS")
print("========================================")

print("Total Sales:", df["Sales"].sum())
print("Average Sales:", df["Sales"].mean())
print("Median Sales:", df["Sales"].median())
print("Minimum Sales:", df["Sales"].min())
print("Maximum Sales:", df["Sales"].max())
print("Standard Deviation:", df["Sales"].std())

print("\n========================================")
print("PROFIT STATISTICS")
print("========================================")

print("Total Profit:", df["Profit"].sum())
print("Average Profit:", df["Profit"].mean())
print("Median Profit:", df["Profit"].median())
print("Minimum Profit:", df["Profit"].min())
print("Maximum Profit:", df["Profit"].max())
print("Standard Deviation:", df["Profit"].std())

print("\n========================================")
print("QUANTITY STATISTICS")
print("========================================")

print("Total Quantity:", df["Quantity"].sum())
print("Average Quantity:", df["Quantity"].mean())
print("Median Quantity:", df["Quantity"].median())
print("Minimum Quantity:", df["Quantity"].min())
print("Maximum Quantity:", df["Quantity"].max())

print("\n========================================")
print("SALES QUARTILES")
print("========================================")

print(df["Sales"].quantile([0.25, 0.50, 0.75]))


# ============================================================
# 3. CATEGORY-WISE ANALYSIS
# ============================================================

category_sales = df.groupby("Category")["Sales"].sum()

print("\n========================================")
print("CATEGORY-WISE SALES")
print("========================================")

print(category_sales)


# ============================================================
# 4. REGION-WISE ANALYSIS
# ============================================================

region_sales = df.groupby("Region")["Sales"].sum()

print("\n========================================")
print("REGION-WISE SALES")
print("========================================")

print(region_sales)


# ============================================================
# 5. CORRELATION ANALYSIS
# ============================================================

print("\n========================================")
print("CORRELATION")
print("========================================")

correlation = df[
    ["Quantity", "Sales", "Cost", "Profit"]
].corr()

print(correlation)


# ============================================================
# 6. CREATE SCREENSHOTS FOLDER
# ============================================================

os.makedirs("screenshots", exist_ok=True)


# ============================================================
# 7. MONTHLY SALES TREND
# ============================================================

monthly_sales = (
    df.groupby(df["Date"].dt.to_period("M"))["Sales"]
    .sum()
)

monthly_sales.index = monthly_sales.index.astype(str)

plt.figure(figsize=(12, 6))

plt.plot(
    monthly_sales.index,
    monthly_sales.values,
    marker="o"
)

plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig(
    "screenshots/monthly_sales_trend.png",
    dpi=300
)

plt.close()


# ============================================================
# 8. SALES BY CATEGORY
# ============================================================

plt.figure(figsize=(8, 6))

category_sales.plot(
    kind="bar"
)

plt.title("Sales by Category")
plt.xlabel("Category")
plt.ylabel("Sales")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig(
    "screenshots/sales_by_category.png",
    dpi=300
)

plt.close()


# ============================================================
# 9. SALES BY REGION
# ============================================================

plt.figure(figsize=(8, 6))

region_sales.plot(
    kind="bar"
)

plt.title("Sales by Region")
plt.xlabel("Region")
plt.ylabel("Sales")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig(
    "screenshots/sales_by_region.png",
    dpi=300
)

plt.close()


# ============================================================
# 10. TOP PRODUCTS BY SALES
# ============================================================

product_sales = (
    df.groupby("Product")["Sales"]
    .sum()
    .sort_values(ascending=False)
)

plt.figure(figsize=(10, 6))

product_sales.plot(
    kind="bar"
)

plt.title("Top Products by Sales")
plt.xlabel("Product")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig(
    "screenshots/top_products.png",
    dpi=300
)

plt.close()


# ============================================================
# 11. PROFIT BY CATEGORY
# ============================================================

category_profit = (
    df.groupby("Category")["Profit"]
    .sum()
)

plt.figure(figsize=(8, 6))

category_profit.plot(
    kind="bar"
)

plt.title("Profit by Category")
plt.xlabel("Category")
plt.ylabel("Profit")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig(
    "screenshots/profit_by_category.png",
    dpi=300
)

plt.close()


# ============================================================
# 12. QUANTITY BY CATEGORY
# ============================================================

category_quantity = (
    df.groupby("Category")["Quantity"]
    .sum()
)

plt.figure(figsize=(8, 6))

category_quantity.plot(
    kind="bar"
)

plt.title("Quantity Sold by Category")
plt.xlabel("Category")
plt.ylabel("Quantity")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig(
    "screenshots/quantity_by_category.png",
    dpi=300
)

plt.close()


# ============================================================
# 13. SALES DISTRIBUTION
# ============================================================

plt.figure(figsize=(10, 6))

sns.histplot(
    df["Sales"],
    bins=30,
    kde=True
)

plt.title("Sales Distribution")
plt.xlabel("Sales")
plt.ylabel("Frequency")
plt.tight_layout()

plt.savefig(
    "screenshots/sales_distribution.png",
    dpi=300
)

plt.close()


# ============================================================
# 14. SALES VS PROFIT
# ============================================================

plt.figure(figsize=(10, 6))

sns.scatterplot(
    data=df,
    x="Sales",
    y="Profit"
)

plt.title("Sales vs Profit Relationship")
plt.xlabel("Sales")
plt.ylabel("Profit")
plt.tight_layout()

plt.savefig(
    "screenshots/sales_vs_profit.png",
    dpi=300
)

plt.close()


print("\n========================================")
print("ALL VISUALIZATIONS CREATED SUCCESSFULLY")
print("Charts saved in: screenshots/")
print("========================================")


# ============================================================
# 15. CUSTOMER TYPE ANALYSIS
# ============================================================

customer_analysis = (
    df.groupby("Customer_Type")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Average_Sales=("Sales", "mean"),
        Average_Profit=("Profit", "mean"),
        Total_Quantity=("Quantity", "sum"),
        Total_Orders=("Order_ID", "count")
    )
    .reset_index()
)

print("\n========================================")
print("CUSTOMER TYPE ANALYSIS")
print("========================================")

print(customer_analysis)

customer_analysis.to_csv(
    "data/customer_type_analysis.csv",
    index=False
)

print("\nCustomer analysis saved successfully.")


# ============================================================
# 16. DISCOUNT VS PROFIT ANALYSIS
# ============================================================

discount_analysis = (
    df.groupby("Discount")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Average_Profit=("Profit", "mean")
    )
    .reset_index()
)

print("\n========================================")
print("DISCOUNT VS PROFIT ANALYSIS")
print("========================================")

print(discount_analysis)

discount_analysis.to_csv(
    "data/discount_profit_analysis.csv",
    index=False
)

print("\nDiscount analysis saved successfully.")


# ============================================================
# 17. SALES FORECASTING
# ============================================================

print("\n========================================")
print("SALES FORECASTING")
print("========================================")

# Monthly sales data
forecast_data = (
    df.groupby(df["Date"].dt.to_period("M"))["Sales"]
    .sum()
    .reset_index()
)

forecast_data.columns = ["Date", "Actual_Sales"]

forecast_data["Date"] = forecast_data["Date"].dt.to_timestamp()

# Numeric time index
forecast_data["Time_Index"] = np.arange(
    len(forecast_data)
)

# Linear regression using numpy
x = forecast_data["Time_Index"].values
y = forecast_data["Actual_Sales"].values

slope, intercept = np.polyfit(x, y, 1)

# Forecast next 6 months
future_count = 6

future_indices = np.arange(
    len(forecast_data),
    len(forecast_data) + future_count
)

future_dates = pd.date_range(
    start=forecast_data["Date"].max()
    + pd.DateOffset(months=1),
    periods=future_count,
    freq="MS"
)

future_forecast = (
    intercept + slope * future_indices
)

future_df = pd.DataFrame({
    "Date": future_dates,
    "Actual_Sales": np.nan,
    "Forecast_Sales": future_forecast
})

# Historical rows
forecast_data["Forecast_Sales"] = np.nan

# Combine historical + future
sales_forecast = pd.concat(
    [
        forecast_data[
            ["Date", "Actual_Sales", "Forecast_Sales"]
        ],
        future_df[
            ["Date", "Actual_Sales", "Forecast_Sales"]
        ]
    ],
    ignore_index=True
)

print("\nForecast Results:")
print(sales_forecast.tail(10))

sales_forecast.to_csv(
    "data/sales_forecast.csv",
    index=False
)

print("\nSales forecast saved successfully.")


# ============================================================
# 18. AI-ASSISTED AUTOMATED BUSINESS INSIGHTS
# ============================================================

print("\n========================================")
print("AI-ASSISTED BUSINESS INSIGHTS")
print("========================================")


# ---------- Overall Performance ----------

total_sales = df["Sales"].sum()
total_profit = df["Profit"].sum()
total_orders = len(df)

profit_margin = (
    total_profit / total_sales
) * 100


# ---------- Product Analysis ----------

product_sales = (
    df.groupby("Product")["Sales"]
    .sum()
)

product_profit = (
    df.groupby("Product")["Profit"]
    .sum()
)

product_quantity = (
    df.groupby("Product")["Quantity"]
    .sum()
)

top_product_sales = product_sales.idxmax()
top_product_sales_value = product_sales.max()

top_product_profit = product_profit.idxmax()
top_product_profit_value = product_profit.max()

top_product_quantity = product_quantity.idxmax()
top_product_quantity_value = product_quantity.max()


# ---------- Region Analysis ----------

region_sales = (
    df.groupby("Region")["Sales"]
    .sum()
)

region_profit = (
    df.groupby("Region")["Profit"]
    .sum()
)

best_region_sales = region_sales.idxmax()
best_region_sales_value = region_sales.max()

best_region_profit = region_profit.idxmax()
best_region_profit_value = region_profit.max()

lowest_region_sales = region_sales.idxmin()
lowest_region_sales_value = region_sales.min()


# ---------- Customer Analysis ----------

customer_sales = (
    df.groupby("Customer_Type")["Sales"]
    .sum()
)

customer_profit = (
    df.groupby("Customer_Type")["Profit"]
    .sum()
)

best_customer_type = customer_sales.idxmax()
best_customer_sales = customer_sales.max()


# ---------- Discount Analysis ----------

discount_median = df["Discount"].median()

low_discount_profit = (
    df[df["Discount"] <= discount_median]["Profit"]
    .mean()
)

high_discount_profit = (
    df[df["Discount"] > discount_median]["Profit"]
    .mean()
)


# ---------- Monthly Analysis ----------

df["Month"] = (
    df["Date"]
    .dt.to_period("M")
    .astype(str)
)

monthly_sales_ai = (
    df.groupby("Month")["Sales"]
    .sum()
)

best_month = monthly_sales_ai.idxmax()
best_month_sales = monthly_sales_ai.max()

lowest_month = monthly_sales_ai.idxmin()
lowest_month_sales = monthly_sales_ai.min()


# ============================================================
# 19. GENERATE AUTOMATED INSIGHTS
# ============================================================

insights = []

insights.append(
    f"1. Overall Performance: The business generated total "
    f"sales of Rs. {total_sales:,.2f} from {total_orders:,} "
    f"orders, with total profit of Rs. {total_profit:,.2f}."
)

insights.append(
    f"2. Profitability: The overall profit margin is "
    f"approximately {profit_margin:.1f}%."
)

insights.append(
    f"3. Product Sales: {top_product_sales} generated the "
    f"highest total sales of Rs. {top_product_sales_value:,.2f}."
)

insights.append(
    f"4. Product Profit: {top_product_profit} generated the "
    f"highest total profit of Rs. {top_product_profit_value:,.2f}."
)

insights.append(
    f"5. Product Quantity: {top_product_quantity} had the "
    f"highest quantity sold with "
    f"{top_product_quantity_value:,.0f} units."
)

insights.append(
    f"6. Regional Performance: {best_region_sales} generated "
    f"the highest regional sales of "
    f"Rs. {best_region_sales_value:,.2f}."
)

insights.append(
    f"7. Regional Profit: {best_region_profit} generated the "
    f"highest regional profit of "
    f"Rs. {best_region_profit_value:,.2f}."
)

insights.append(
    f"8. Attention Area: {lowest_region_sales} recorded the "
    f"lowest regional sales of "
    f"Rs. {lowest_region_sales_value:,.2f}."
)

insights.append(
    f"9. Customer Analysis: {best_customer_type} customers "
    f"generated the highest total sales of "
    f"Rs. {best_customer_sales:,.2f}."
)


# ---------- Discount Insight ----------

if high_discount_profit < low_discount_profit:

    insights.append(
        "10. Discount Observation: Higher discount levels "
        "are associated with lower average profit in the "
        "current dataset. Discount usage should therefore "
        "be monitored carefully."
    )

else:

    insights.append(
        "10. Discount Observation: Higher discount levels "
        "do not show lower average profit in the current "
        "dataset. Discount impact should still be monitored."
    )


insights.append(
    f"11. Monthly Performance: {best_month} recorded the "
    f"highest monthly sales of "
    f"Rs. {best_month_sales:,.2f}."
)

insights.append(
    f"12. Lowest Monthly Sales: {lowest_month} recorded the "
    f"lowest monthly sales of "
    f"Rs. {lowest_month_sales:,.2f}."
)


# ============================================================
# 20. AUTOMATED BUSINESS RECOMMENDATIONS
# ============================================================

recommendations = []

recommendations.append(
    f"- Focus on {top_product_sales}, which currently has "
    f"the highest sales."
)

recommendations.append(
    f"- Monitor {lowest_region_sales} and identify reasons "
    f"for its lower sales."
)

recommendations.append(
    f"- Study the purchasing behaviour of "
    f"{best_customer_type} customers and encourage "
    f"repeat purchases."
)

recommendations.append(
    "- Review discount strategies regularly to maintain "
    "healthy profit margins."
)

recommendations.append(
    f"- Use the monthly sales trend, especially the "
    f"performance of {best_month}, for future business "
    f"planning."
)

recommendations.append(
    "- Continue monitoring sales, profit, orders and "
    "profit margin through the Power BI dashboard."
)


# ============================================================
# 21. DISPLAY AI INSIGHTS
# ============================================================

for insight in insights:
    print(insight)

print("\n========================================")
print("AI-ASSISTED RECOMMENDATIONS")
print("========================================")

for recommendation in recommendations:
    print(recommendation)


# ============================================================
# 22. SAVE AI INSIGHTS TO TEXT FILE
# ============================================================

with open(
    "data/ai_business_insights.txt",
    "w",
    encoding="utf-8"
) as file:

    file.write(
        "AI-ASSISTED BUSINESS INSIGHTS\n"
    )

    file.write(
        "=" * 60 + "\n\n"
    )

    for insight in insights:
        file.write(
            insight + "\n\n"
        )

    file.write(
        "\nAI-ASSISTED BUSINESS RECOMMENDATIONS\n"
    )

    file.write(
        "=" * 60 + "\n\n"
    )

    for recommendation in recommendations:
        file.write(
            recommendation + "\n"
        )


# ============================================================
# 23. FINAL MESSAGE
# ============================================================

print("\n========================================")
print("ANALYSIS COMPLETED SUCCESSFULLY")
print("========================================")

print("\nGenerated files:")

print("1. data/customer_type_analysis.csv")
print("2. data/discount_profit_analysis.csv")
print("3. data/sales_forecast.csv")
print("4. data/ai_business_insights.txt")

print("\nGenerated charts:")
print("Charts saved inside: screenshots/")

print("\nAI BUSINESS INSIGHTS SAVED SUCCESSFULLY.")
print("========================================")