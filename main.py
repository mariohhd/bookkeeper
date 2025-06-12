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

def calculate_yearly_expenses_by_type(df, type):
    concepts = config['concepts'][type]
    # Filter by type expenses
    filtered_df = df[df['Concept'].str.contains('|'.join(concepts))]
  
    # Convert date to datetime
    filtered_df['Date'] = pd.to_datetime(filtered_df['Date'])

    # Extract year from date
    filtered_df['Year'] = filtered_df['Date'].dt.year

    # Calculate yearly totals
    yearly_totals = filtered_df.groupby('Year')['Amount'].sum()

    # Plot the data
    plt.figure(figsize=(12, 6))
    yearly_totals.plot(kind='bar', color='blue')

    plt.title('Yearly ' + type + ' Expenses')
    plt.xlabel('Year')
    plt.ylabel('Amount')
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig('yearly_' + type + '_expenses.png')
    plt.close()
    
    # Print summary
    print('\nYearly ' + type + ' Expenses Summary:')
    print(yearly_totals)

def calculate_yearly_contribution_by_neighbour(df, neighbour, year):
    current_contribution = config['current_contribution'][neighbour]
    concepts = config['concepts']['neighbours'][neighbour]
    
    # First convert date to datetime in the original DataFrame
    df['Date'] = pd.to_datetime(df['Date'])
    
    # Then filter by type expenses
    filtered_df = df[df['Concept'].str.contains('|'.join(concepts))]
    
    # Filter by specified year
    yearly_df = filtered_df[filtered_df['Date'].dt.year == year]

    # Calculate yearly total for this neighbor
    yearly_total = yearly_df['Amount'].sum()

    # Calculate percentage of total contribution
    percentage = (yearly_total / current_contribution * 100) if current_contribution > 0 else 0

    # Print results
    print(f"\nContributions for {neighbour} in {year}:")
    print(f"- Total contribution: €{yearly_total:.2f}")
    print(f"- Percentage of total: {percentage:.2f}%")
    
    return yearly_total, percentage, neighbour
    

# Main execution
file = 'Movimientos.xls'
df = read_excel_with_colors(file)
calculate_yearly_expenses_by_type(df, 'water')
calculate_yearly_expenses_by_type(df, 'cleaning')
calculate_yearly_expenses_by_type(df, 'electricity')
calculate_yearly_expenses_by_type(df, 'bank')
calculate_yearly_expenses_by_type(df, 'insurance')
calculate_yearly_contribution_by_neighbour(df, '1A', 2024)
calculate_yearly_contribution_by_neighbour(df, '1B', 2024)
calculate_yearly_contribution_by_neighbour(df, '1C', 2024)
calculate_yearly_contribution_by_neighbour(df, '2A', 2024)
calculate_yearly_contribution_by_neighbour(df, '2B', 2024)
calculate_yearly_contribution_by_neighbour(df, '2C', 2024)
calculate_yearly_contribution_by_neighbour(df, 'LOCAL', 2024)
        

print("\nAnalysis complete! Check 'yearly_*.png' for the visualization.")
