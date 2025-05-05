import pandas as pd
import yaml
import matplotlib.pyplot as plt
from datetime import datetime
import xlrd
import numpy as np

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)
    
def read_excel_with_colors(file_path):
    # Read the Excel file with xlrd
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)
    
    # Create lists to store data
    dates = []
    concepts = []
    amounts = []
    descriptions = []
    
    # Skip header row
    for row_idx in range(1, sheet.nrows):
        # Get date from first column
        date = sheet.cell(row_idx, 0).value
        if date:
            # Convert Excel serial date to datetime
            try:
                date = datetime(*xlrd.xldate_as_tuple(date, wb.datemode))
                dates.append(date)
            except:
                dates.append(date)  # Keep original value if conversion fails
            
        # Get description from second column
        description = sheet.cell(row_idx, 1).value
        if description:
            descriptions.append(description)
            
        # Get concept from third column
        concept = sheet.cell(row_idx, 2).value
        if concept:
            concepts.append(concept)
            
        # Get amount from fourth column
        amount = sheet.cell(row_idx, 3).value
        if amount:
            # Convert to float and handle negative numbers
            try:
                amount = float(amount)
                amounts.append(amount)
            except ValueError:
                continue
    
    # Create DataFrame
    df = pd.DataFrame({
        'Date': dates,
        'Description': descriptions,
        'Concept': concepts,
        'Amount': amounts
    })
    
    # Print first few rows
    print("\nProcessed DataFrame:")
    print(df.head())
    
    return df

def calculate_yearly_water_expenses(df):
    # Filter by water expenses
    water_df = df[df['Concept'].isin(config['concepts']['water'])]
    print(water_df)
    # Convert date to datetime
    water_df['Date'] = pd.to_datetime(water_df['Date'])
    
    # Extract year from date
    water_df['Year'] = water_df['Date'].dt.year
    
    # Calculate yearly totals
    yearly_totals = water_df.groupby('Year')['Amount'].sum()
    
    # Plot the data
    plt.figure(figsize=(12, 6))
    yearly_totals.plot(kind='bar', color='blue')
    
    plt.title('Yearly Water Expenses')
    plt.xlabel('Year')
    plt.ylabel('Amount')
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig('yearly_water_expenses.png')
    plt.close()
    
    # Print summary
    print("\nYearly Water Expenses Summary:")
    print(yearly_totals)

# Main execution
file = 'Movimientos.xls'
df = read_excel_with_colors(file)
calculate_yearly_water_expenses(df)
print("\nAnalysis complete! Check 'yearly_water_expenses.png' for the visualization.")
