import pandas as pd
from openpyxl import load_workbook
from unicodedata import normalize

def get_sheets(filename):
    wb = load_workbook(filename, read_only=True)
    sheets = wb.sheetnames
    return sheets

def get_dataframe(filename, sheet):
    df = pd.read_excel(filename, sheet_name=sheet, dtype=str)

    # Columns to lower case and unicode
    df.columns = [normalize('NFKD', col.lower()).encode('ASCII', 'ignore').decode('ASCII') for col in df.columns]
    
    missing_columns = []
    if 'individuo' not in df.columns: missing_columns.append('individuo')
    if 'pai' not in df.columns: missing_columns.append('pai')
    if 'mae' not in df.columns: missing_columns.append('mae')
    if 'sexo' not in df.columns: missing_columns.append('sexo')
    if len(missing_columns) > 0:
        raise Exception(f'Planilha {sheet} não contém as colunas: {", ".join(missing_columns)}')

    df = df.rename(columns={
        'Indivíduo': 'individuo',
        'Pai': 'pai',
        'Mae': 'mae',
        'SEXO': 'sexo',
    })
    df = df[['individuo', 'pai', 'mae', 'sexo']]

    df.replace('x', 'X', inplace=True)
    df.replace('X', '', inplace=True)
    df['sexo'] = df['sexo'].str.upper()

    df['individuo'] = df['individuo'].apply(lambda x: x.rjust(10, ' ') if x.isnumeric() else x)
    df.sort_values(by=['individuo'], inplace=True)
    df['individuo'] = df['individuo'].str.strip()

    return df

class GrauParentesco:
    def __init__(self, filename, sheet):
        self.dfCoelhos = get_dataframe(filename, sheet)

    @property
    def get_entradadados(self):
        return self.dfCoelhos.to_string()
    