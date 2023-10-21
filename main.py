#%% Abrir arquivo de entrada e ajuste de dados
import pandas as pd

df = pd.read_excel('Dados coelhos- ascendentes.xlsx', dtype=str)
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

df['individuo'].tolist()
# %% Classe Coelho
from collections import defaultdict

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
    

class CoeficienteParentesco:
    def __init__(self, coelhos):
        self.coelhos = coelhos
        self.coeficientes_parentesco = []
    
    def calc_coeficiente_coelho(self, coelho, grau=0):
        if self.visitados[coelho.nome]:
            return
        self.visitados[coelho.nome] = True

        if coelho.get_sexo() == 'F':
            self.coeficiente_coelho[coelho.nome]+=(1/2)**grau
        
        if coelho.get_pai():
            self.calc_coeficiente_coelho(coelho.get_pai(), grau+1)
        if coelho.get_mae():
            self.calc_coeficiente_coelho(coelho.get_mae(), grau+1)
        for filho in coelho.filhos:
            self.calc_coeficiente_coelho(filho, grau+1)

    def calcular(self):
        visitados = set()
        for coelho in self.coelhos.values():
            if coelho.get_sexo() == 'M':
                self.coeficiente_coelho = defaultdict(int)
                self.visitados = defaultdict(bool)
                self.calc_coeficiente_coelho(coelho)
                
                row = {
                    'coelhos': coelho.nome,
                    **self.coeficiente_coelho
                }
                self.coeficientes_parentesco.append(row)
        
    def get_coeficientes_parentesco(self):
        return self.coeficientes_parentesco
#%% Criação da árvore genealógica
coelhos = {}
for index, row in df.iterrows():
    nome = row['individuo']
    sexo = row['sexo']
    coelhos[nome] = Coelho(nome, sexo)

for index, row in df.iterrows():
    nome = row['individuo']
    pai = row['pai']
    mae = row['mae']
    if pai:
        coelhos[nome].add_pai(coelhos[pai])
    if mae:
        coelhos[nome].add_mae(coelhos[mae])

for coelho in coelhos.values():
    if coelho.get_pai() == None and coelho.get_mae() == None:
        print(coelho.get_arvore())

# %% Cálculo dos coeficientes de parentesco
import pandas as pd

cp = CoeficienteParentesco(coelhos)
cp.calcular()
coeficientes_parentesco = cp.get_coeficientes_parentesco()
df = pd.DataFrame(coeficientes_parentesco)
df = df.reindex(sorted(df.columns), axis=1)
df.insert(0, 'coelhos', df.pop('coelhos'))  
df

#%%
d = defaultdict(int)
d['a'] += 1
d
# %%
