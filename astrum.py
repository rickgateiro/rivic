import pandas as pd
from IPython.display import display, HTML

# Ler o arquivo CSV com o pandas
dataframe = pd.read_csv('astrumgov.csv')

# Agrupar os dados por situação e operador
agrupado = dataframe.groupby(['Situação', 'Nome operador (avaliação)']).size().reset_index(name='Número de documentos')

# Verificar a situação 'ANÁLISE DE AVALIAÇÃO' e criar a coluna 'STATUS' com a mensagem 'aprovada'
agrupado['STATUS'] = ''
agrupado.loc[agrupado['Situação'] == 'ANÁLISE DE AVALIAÇÃO', 'STATUS'] = 'Enviada para o Analista Rivic'
agrupado.loc[agrupado['Situação'] == 'EM AVALIAÇÃO', 'STATUS'] = 'Pendente de envio'
agrupado.loc[agrupado['Situação'] == 'APROVAÇÃO DE AVALIAÇÃO', 'STATUS'] = 'Analisada pela Equipe RIVIC'

# Mostrar os usuários atribuídos a cada situação
for index, row in agrupado.iterrows():
    situacao = row['Situação']
    operador = row['Nome operador (avaliação)']
    contagem = row['Número de documentos']
    status = row['STATUS']
    print(f"Situação: {situacao}, Operador: {operador}, Contagem: {contagem}, STATUS: {status}")

# Salvar a listagem em um arquivo Excel
agrupado.to_excel('listagem.xlsx', index=False)