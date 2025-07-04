import pandas as pd
import numpy as np

# Create sample data for multiple sheets
data1 = {
    'ID': [1, 2, 3, 4, 5],
    'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
    'Age': [25, 30, 35, 40, 45],
    'City': ['New York', 'London', 'Paris', 'Tokyo', 'Sydney']
}

data2 = {
    'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones'],
    'Price': [999.99, 25.50, 75.00, 299.99, 150.00],
    'Stock': [50, 200, 100, 75, 120]
}

data3 = {
    'Date': pd.date_range('2024-01-01', periods=5, freq='D'),
    'Sales': [1000, 1500, 1200, 1800, 2000],
    'Region': ['North', 'South', 'East', 'West', 'Central']
}

# Create Excel file with multiple sheets
with pd.ExcelWriter('examples/multi_sheet_test.xlsx', engine='openpyxl') as writer:
    pd.DataFrame(data1).to_excel(writer, sheet_name='Employees', index=False)
    pd.DataFrame(data2).to_excel(writer, sheet_name='Products', index=False)
    pd.DataFrame(data3).to_excel(writer, sheet_name='Sales', index=False)

print("Multi-sheet Excel file created successfully!")
