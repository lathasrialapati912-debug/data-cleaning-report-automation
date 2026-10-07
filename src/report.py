import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference

# --------------------------------
# 1. Load cleaned data
# --------------------------------

input_file = "data/cleaned_sales_data.csv"

df = pd.read_csv(input_file)

# --------------------------------
# 2. Create report folder
# --------------------------------

os.makedirs("reports", exist_ok=True)

# --------------------------------
# 3. Calculate KPIs
# --------------------------------

total_revenue = df["Total_Sales"].sum()
total_quantity = df["Quantity"].sum()
average_sales = df["Total_Sales"].mean()

best_product = (
    df.groupby("Product")["Total_Sales"]
    .sum()
    .sort_values(ascending=False)
)

best_city = (
    df.groupby("City")["Total_Sales"]
    .sum()
    .sort_values(ascending=False)
)

monthly_sales = (
    df.groupby("Sales_Date")["Total_Sales"]
    .sum()
)

# --------------------------------
# 4. Create summary
# --------------------------------

summary = pd.DataFrame({
    "KPI": [
        "Total Revenue",
        "Total Quantity Sold",
        "Average Transaction Value",
        "Best Product",
        "Best City"
    ],
    "Value": [
        total_revenue,
        total_quantity,
        round(average_sales, 2),
        best_product.index[0],
        best_city.index[0]
    ]
})

# --------------------------------
# 5. Create Excel report
# --------------------------------

report_file = "reports/automated_sales_report.xlsx"

with pd.ExcelWriter(report_file, engine="openpyxl") as writer:

    summary.to_excel(
        writer,
        sheet_name="Dashboard",
        index=False
    )

    best_product.to_excel(
        writer,
        sheet_name="Product Analysis"
    )

    best_city.to_excel(
        writer,
        sheet_name="City Analysis"
    )

    monthly_sales.to_excel(
        writer,
        sheet_name="Monthly Sales"
    )

    df.to_excel(
        writer,
        sheet_name="Cleaned Data",
        index=False
    )

# --------------------------------
# 6. Format Excel workbook
# --------------------------------

workbook = load_workbook(report_file)

# Dashboard formatting
dashboard = workbook["Dashboard"]

dashboard["A1"] = "DATA CLEANING & SALES REPORT"
dashboard["A1"].font = Font(
    bold=True,
    size=16
)

dashboard["A1"].alignment = Alignment(
    horizontal="center"
)

dashboard.merge_cells("A1:B1")

# Header formatting
for cell in dashboard[2]:
    cell.font = Font(bold=True)

# Column widths
dashboard.column_dimensions["A"].width = 30
dashboard.column_dimensions["B"].width = 25

# Product sheet
product_sheet = workbook["Product Analysis"]

product_sheet.column_dimensions["A"].width = 20
product_sheet.column_dimensions["B"].width = 20

# City sheet
city_sheet = workbook["City Analysis"]

city_sheet.column_dimensions["A"].width = 20
city_sheet.column_dimensions["B"].width = 20

# Monthly sales sheet
monthly_sheet = workbook["Monthly Sales"]

monthly_sheet.column_dimensions["A"].width = 20
monthly_sheet.column_dimensions["B"].width = 20

# Cleaned data sheet
cleaned_sheet = workbook["Cleaned Data"]

for column in cleaned_sheet.columns:
    max_length = 0
    column_letter = column[0].column_letter

    for cell in column:
        if cell.value is not None:
            max_length = max(
                max_length,
                len(str(cell.value))
            )

    cleaned_sheet.column_dimensions[
        column_letter
    ].width = min(max_length + 2, 25)

# --------------------------------
# 7. Add Product Sales Chart
# --------------------------------

chart = BarChart()

chart.title = "Sales by Product"
chart.y_axis.title = "Total Sales"
chart.x_axis.title = "Product"

data = Reference(
    product_sheet,
    min_col=2,
    min_row=1,
    max_row=product_sheet.max_row
)

categories = Reference(
    product_sheet,
    min_col=1,
    min_row=2,
    max_row=product_sheet.max_row
)

chart.add_data(
    data,
    titles_from_data=True
)

chart.set_categories(categories)

dashboard.add_chart(
    chart,
    "D3"
)

# --------------------------------
# 8. Add City Sales Chart
# --------------------------------

city_chart = BarChart()

city_chart.title = "Sales by City"
city_chart.y_axis.title = "Total Sales"
city_chart.x_axis.title = "City"

city_data = Reference(
    city_sheet,
    min_col=2,
    min_row=1,
    max_row=city_sheet.max_row
)

city_categories = Reference(
    city_sheet,
    min_col=1,
    min_row=2,
    max_row=city_sheet.max_row
)

city_chart.add_data(
    city_data,
    titles_from_data=True
)

city_chart.set_categories(
    city_categories
)

dashboard.add_chart(
    city_chart,
    "D20"
)

# --------------------------------
# 9. Save workbook
# --------------------------------

workbook.save(report_file)

print("================================")
print("PROFESSIONAL REPORT CREATED")
print("================================")
print("Total Revenue:", total_revenue)
print("Total Quantity:", total_quantity)
print(
    "Average Transaction:",
    round(average_sales, 2)
)
print("Best Product:", best_product.index[0])
print("Best City:", best_city.index[0])
print("--------------------------------")
print("Report saved to:")
print(report_file)