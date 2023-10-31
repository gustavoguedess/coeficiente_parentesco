import pandas as pd
from openpyxl import load_workbook
from unicodedata import normalize
from collections import defaultdict

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

class Coelho:
    def __init__(self, nome, sexo):
        self.nome = nome
        self.sexo = sexo
        self.pai = None
        self.mae = None
        self.filhos = []

    def add_pai(self, pai):
        if pai.get_sexo() != 'M':
            raise Exception(f'Pai do coelho {self.nome}, coelho {pai.nome}, é fêmea')
        self.pai = pai
        pai.add_filho(self)
    
    def add_mae(self, mae):
        if mae.get_sexo() != 'F':
            raise Exception(f'Mãe do coelho {self.nome}, coelho {mae.nome}, é macho')
        self.mae = mae
        mae.add_filho(self)

    def add_filho(self, filho):
        self.filhos.append(filho)

    def get_arvore(self, nivel=0):
        arvore = f'{"  "*nivel}{self.nome} ({self.sexo})\n'
        for filho in self.filhos:
            arvore += filho.get_arvore(nivel+1)
        return arvore
    def get_pai(self):
        return self.pai
    def get_mae(self):
        return self.mae
    def get_sexo(self):
        return self.sexo

class GrauParentesco:
    def __init__(self, filename, sheet):
        self.dfCoelhos = get_dataframe(filename, sheet)
        self.coelhos = self.criar_coelhos()

    @property
    def get_entradadados(self):
        return self.dfCoelhos.to_string()
    
    def criar_coelhos(self):
        coelhos = {}
        for index, row in self.dfCoelhos.iterrows():
            nome = row['individuo']
            sexo = row['sexo']
            coelhos[nome] = Coelho(nome, sexo)

        for index, row in self.dfCoelhos.iterrows():
            nome = row['individuo']
            pai = row['pai']
            mae = row['mae']
            if pai:
                coelhos[nome].add_pai(coelhos[pai])
            if mae:
                coelhos[nome].add_mae(coelhos[mae])
        return coelhos

    def calc_coeficiente_individual(self, coelho_inicial, coelho_atual=None, grau=0, descendente=False):
        if coelho_atual == None: coelho_atual = coelho_inicial
        self.visitados[coelho_atual.nome] = True

        # Se o coelho atual for do sexo oposto ao coelho inicial
        if coelho_atual.get_sexo() != coelho_inicial.get_sexo():
            self.coeficiente_parentesco[coelho_inicial.nome][coelho_atual.nome] += 0.5**grau
        
        for filho in coelho_atual.filhos:
            if not self.visitados[filho.nome]:
                self.calc_coeficiente_macho(coelho_inicial, filho, grau+1, descendente=True)

        if coelho_atual.get_pai() and not descendente:
            self.calc_coeficiente_macho(coelho_inicial, coelho_atual.get_pai(), grau+1)
        if coelho_atual.get_mae() and not descendente:
            self.calc_coeficiente_macho(coelho_inicial, coelho_atual.get_mae(), grau+1)
        
        self.visitados[coelho_atual.nome] = False

    def calcular(self):
        femeas = [coelho.nome for coelho in self.coelhos.values() if coelho.get_sexo() == 'F']
        machos = [coelho.nome for coelho in self.coelhos.values() if coelho.get_sexo() == 'M']
        print(f'Fêmeas: {femeas}')
        print(f'Machos: {machos}')

        # Inicializa o mapa de coeficientes de parentesco
        self.coeficiente_parentesco = defaultdict(dict)
        for macho in machos:
            self.coeficiente_parentesco[macho] = {femea: 0 for femea in femeas}
        # Para cada coelho macho, calcula o coeficiente de parentesco
        for coelho in self.coelhos.values():
            if coelho.get_sexo() == 'M':
                self.visitados = defaultdict(bool)
                print(f'Calculando coeficientes de parentesco para o coelho {coelho.nome}')
                self.calc_coeficiente_macho(coelho)
        