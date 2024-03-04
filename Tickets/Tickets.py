import pandas as pd

def atualizar_planilha_usuarios(importacao_pessoas_path, usuarios_path):
    importacao_pessoas_df = pd.read_excel(importacao_pessoas_path)
    
    importacao_pessoas_df['E-mail'] = importacao_pessoas_df['E-mail'].astype(str)
    
    usuarios_df = pd.read_excel(usuarios_path)
    
    usuarios_df['Email Domain'] = usuarios_df['Email'].apply(lambda x: x.split('@')[-1].split('.')[0].lower())
    usuarios_df['Organização Normalizada'] = usuarios_df['Organização'].str.lower()
    usuarios_filtrados = usuarios_df[usuarios_df.apply(lambda x: x['Organização Normalizada'] in x['Email Domain'], axis=1)]
    
    for org in usuarios_filtrados['Organização'].unique():
        emails = usuarios_filtrados[usuarios_filtrados['Organização'] == org]['Email'].tolist()
        linha = importacao_pessoas_df[importacao_pessoas_df['Organizações'] == org].index
        if not linha.empty:
            importacao_pessoas_df.at[linha[0], 'E-mail'] = ', '.join(emails)
    
    importacao_pessoas_original_df = pd.read_excel(importacao_pessoas_path)
    importacao_pessoas_df['Observações'] = importacao_pessoas_original_df['Observações']
    
    caminho_atualizado = importacao_pessoas_path.replace('.xlsx', '_Atualizada_Organizacoes.xlsx')
    importacao_pessoas_df.to_excel(caminho_atualizado, index=False)
    print(f"Planilha atualizada salva em: {caminho_atualizado}")

importacao_pessoas_path = r''
usuarios_path = r''
atualizar_planilha_usuarios(importacao_pessoas_path, usuarios_path)
