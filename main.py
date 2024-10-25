import pandas as pd

# Load the Excel file
file_path = "data.xlsx"  # Ensure 'data.xlsx' is in the same directory as the script
excel_data = pd.read_excel(file_path)

# Convert 'Order Date' to datetime
excel_data['Order Date'] = pd.to_datetime(excel_data['Order Date'])

# 1. Overall Sales Trend
sales_trend = excel_data.groupby(excel_data['Order Date'].dt.to_period('M'))['Sales'].sum()

# 2. Sales by Category
sales_by_category = excel_data.groupby('Category')['Sales'].sum().sort_values(ascending=False)

# 3. Sales by Sub-Category
sales_by_subcategory = excel_data.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=False)

# 4. Profit Analysis
profit_trend = excel_data.groupby(excel_data['Order Date'].dt.to_period('M'))['Profit'].sum()

# 5. Profit by Customer Segment
profit_by_segment = excel_data.groupby('Segment')['Profit'].sum().sort_values(ascending=False)

# 6. Top 10 Products by Sales
top_10_products = excel_data.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10)

# 7. Most Selling Products
most_selling_products = excel_data.groupby('Product Name')['Quantity'].sum().sort_values(ascending=False).head(10)

# 8. Most Preferred Ship Mode
preferred_ship_mode = excel_data.groupby('Ship Mode')['Sales'].sum().sort_values(ascending=False)

# 9. Most Profitable Category and Sub-Category
most_profitable_category = excel_data.groupby('Category')['Profit'].sum().sort_values(ascending=False)
most_profitable_subcategory = excel_data.groupby('Sub-Category')['Profit'].sum().sort_values(ascending=False)

# Save results to an Excel file
output_file_path = "sales_analysis_output.xlsx"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    sales_trend.to_frame(name='Sales Trend').to_excel(writer, sheet_name='Sales Trend')
    sales_by_category.to_frame(name='Sales by Category').to_excel(writer, sheet_name='Sales by Category')
    sales_by_subcategory.to_frame(name='Sales by Sub-Category').to_excel(writer, sheet_name='Sales by Sub-Category')
    profit_trend.to_frame(name='Profit Trend').to_excel(writer, sheet_name='Profit Trend')
    profit_by_segment.to_frame(name='Profit by Segment').to_excel(writer, sheet_name='Profit by Segment')
    top_10_products.to_frame(name='Top 10 Products by Sales').to_excel(writer, sheet_name='Top 10 Products')
    most_selling_products.to_frame(name='Most Selling Products').to_excel(writer, sheet_name='Most Selling Products')
    preferred_ship_mode.to_frame(name='Preferred Ship Mode').to_excel(writer, sheet_name='Preferred Ship Mode')
    most_profitable_category.to_frame(name='Most Profitable Category').to_excel(writer, sheet_name='Profitable Category')
    most_profitable_subcategory.to_frame(name='Most Profitable Sub-Category').to_excel(writer, sheet_name='Profitable Sub-Category')

print(f"Analysis results saved to {output_file_path}")
