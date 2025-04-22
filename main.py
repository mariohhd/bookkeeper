import pandas as pd
file = 'test.xlsx'
df = pd.read_excel(file)
print(df.head())
