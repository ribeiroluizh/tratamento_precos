import pandas as pd
from fuzzywuzzy import process

path_to_bautech = r'C:\Users\ribeiroluizh\Documents\python\atualização de preços\PREÇOS TOP 14 FEVEREIRO.xlsx'
df_bautech = pd.read_excel(path_to_bautech, sheet_name=1, header=None)

df_bautech.columns = ['Cod Produto','Produto','Emb','Especial','Normal','Percentual','Obs']
df_bautech = df_bautech.drop([0, 1, 2], axis=0).reset_index(drop=True)
df_bautech = df_bautech.astype({'Emb': 'int64','Especial': 'float32', 'Normal': 'float32'})

path_to_syscom = r'C:\Users\ribeiroluizh\Documents\python\atualização de preços\bautech.xlsx'
df_syscom = pd.read_excel(path_to_syscom, header=None)

df_syscom = df_syscom.drop([0,2,5,6,7,9,11,12,13,14,15,23,24,25,26,27,28,29,30,31,32,10,17,18,19,20,21,22], axis=1)
df_syscom = df_syscom.drop([0,1,2], axis=0)
df_syscom = df_syscom.reset_index(drop=True)
df_syscom.columns = ['REF','PRODUTOS','COD','TAB3','DIV']
df_syscom = df_syscom.astype({'REF': 'int64','TAB3': 'float32', 'DIV': 'int64'})

def preprocessar_texto(texto):
    return texto.upper().strip()
df_bautech['Cod Produto'] = df_bautech['Cod Produto'].astype(str).str.strip()
df_syscom['COD'] = df_syscom['COD'].astype(str).str.strip()
df_bautech['Produto'] = df_bautech['Produto'].astype(str).apply(preprocessar_texto)
df_syscom['PRODUTOS'] = df_syscom['PRODUTOS'].astype(str).apply(preprocessar_texto)


def encontrar_correspondencias(nome_produto, lista_produtos, limite=90):
    correspondencia = process.extractOne(nome_produto, lista_produtos, score_cutoff=limite)
    return correspondencia if correspondencia else ('Indefinido', 0)

lista_produtos_bautech = df_bautech['Produto'].unique()
df_syscom['Melhor Correspondência'] = df_syscom['PRODUTOS'].apply(lambda x: encontrar_correspondencias(x, lista_produtos_bautech))
df_syscom['Nome Bautech'] = df_syscom['Melhor Correspondência'].apply(lambda x: x[0])

df_syscom = pd.merge(
    df_syscom[['REF', 'PRODUTOS', 'COD', 'TAB3', 'Nome Bautech']],
    df_bautech[['Cod Produto', 'Especial', 'Normal']],
    left_on='COD', 
    right_on='Cod Produto', 
    how='left'
)

mascara = (df_syscom['Nome Bautech']  != 'Indefinido') & (df_syscom['Normal'].isna())
sem_preço = df_syscom.loc[mascara, ['REF', 'PRODUTOS', 'Nome Bautech']].copy()
sem_preço = sem_preço.reset_index(drop=True)
sem_preço = sem_preço.drop([17], axis=0).reset_index(drop=True)

sem_preço = pd.merge(
    sem_preço[['REF', 'PRODUTOS', 'Nome Bautech']],
    df_bautech[['Produto', 'Especial', 'Normal']],
    left_on='Nome Bautech', 
    right_on='Produto', 
    how='left'
)

df_syscom.loc[mascara, :] = pd.merge(
    df_syscom[['REF', 'PRODUTOS', 'COD', 'TAB3', 'Nome Bautech']],
    sem_preço[['REF', 'Especial', 'Normal']],
    left_on='REF', 
    right_on='REF', 
    how='left'
)

df_syscom = df_syscom.drop(['Nome Bautech','Cod Produto'], axis=1)

path_df_syscom = r'C:\Users\ribeiroluizh\Documents\python\atualização de preços\planilha_final_bautech.xlsx'
df_syscom.to_excel(path_df_syscom, index=False)

print(f'A Planilha foi salva com sucesso em {path_df_syscom}')