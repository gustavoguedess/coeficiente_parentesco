# %% Cálculo dos coeficientes de parentesco
from coelhos import GrauParentesco
import pandas as pd

EXCEL_FILE_NAME = 'coelhos.xlsx'

gp = GrauParentesco(EXCEL_FILE_NAME, 'Planilha1')
# cp.calcular()

coelho_origem = '388'
coelho_destino = '504'
print(gp.coelhos[coelho_origem])
print(gp.coelhos[coelho_destino])
# print(gp.coeficiente_parentesco[coelho])
for parentesco in gp.coeficientes_detalhados:
    if parentesco['origem'] == coelho_origem and parentesco['destino'] == coelho_destino:
        print(parentesco)


#%%
gp.salvar_coeficientes()
gp.salvar_coeficientes_detalhados()
