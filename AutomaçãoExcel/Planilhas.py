import pandas as pd

erick_arturia_df = pd.read_excel("")
chamados_df = pd.read_excel("")

mapeamento_email = erick_arturia_df.set_index('Razão social')['E-mail'].to_dict()

chamados_df['Email'] = chamados_df['Organização'].map(mapeamento_email)

chamados_df.to_excel("chamados_atualizados.xlsx", index=False)

print("Planilha chamados atualizada com sucesso e salva como 'chamados_Atualizados.xlsx'.")
