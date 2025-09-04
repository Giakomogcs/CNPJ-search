import pandas as pd

try:
    df = pd.read_excel('Empresas Dados - 72 empresas.xlsx')
    print("Columns in 'Empresas Dados - 72 empresas.xlsx':")
    print(df.columns.tolist())
except Exception as e:
    print(f"Error reading file: {e}")
