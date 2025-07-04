import pandas as pd
import numpy as np

# Create similar data for comparison
data1 = {
    'ID': [1, 2, 3, 4, 6],  # Different ID 6 instead of 5
    'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Frank'],  # Different name Frank instead of Eve
    'Age': [25, 30, 35, 42, 45],  # Different age for David
    'City': ['New York', 'London', 'Paris', 'Tokyo', 'Berlin']  # Different city for Frank
}

data2 = {
    'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Speakers'],  # Different product
    'Price': [999.99, 25.50, 80.00, 299.99, 175.00],  # Different prices
    'Stock': [45, 200, 95, 75, 130]  # Different stock levels
}

data3 = {
    'Date': pd.date_range('2024-01-01', periods=5, freq='D'),
    'Sales': [1000, 1600, 1200, 1800, 2100],  # Different sales figures
    'Region': ['North', 'South', 'East', 'West', 'Central']
}

# Create Excel file with multiple sheets (same structure, different data)
with pd.ExcelWriter('examples/multi_sheet_test_2.xlsx', engine='openpyxl') as writer:
    pd.DataFrame(data1).to_excel(writer, sheet_name='Employees', index=False)
    pd.DataFrame(data2).to_excel(writer, sheet_name='Products', index=False)
    pd.DataFrame(data3).to_excel(writer, sheet_name='Sales', index=False)

print("Second multi-sheet Excel file created successfully!")
