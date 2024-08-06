import pandas as pd

# Carregue o arquivo Excel
df = pd.read_excel('teste.xlsx')

# Inicialize listas para armazenar os termos repetidos e não repetidos
termos_repetidos = []
termos_nao_repetidos = []

# Itere pelas linhas do DataFrame
for index, row in df.iterrows():
    texto_a = str(row['A'])  # Obtém o conteúdo da coluna A como texto
    texto_b = str(row['B'])  # Obtém o conteúdo da coluna B como texto

    # Verifique se o texto de A está presente em B e vice-versa
    if texto_a in texto_b and texto_a != '':
        termos_repetidos.append(texto_a)
        termos_nao_repetidos.append('')  # Adicione uma célula em branco para manter o mesmo número de elementos
    elif texto_b in texto_a and texto_b != '':
        termos_repetidos.append(texto_b)
        termos_nao_repetidos.append('')  # Adicione uma célula em branco para manter o mesmo número de elementos
    else:
        termos_repetidos.append('')  # Adicione uma célula em branco se não houver termos repetidos
        termos_nao_repetidos.append(texto_a)  # Adicione o texto da coluna A como termo não repetido

# Crie uma nova coluna "C" com os termos repetidos
df['C'] = termos_repetidos

# Crie uma nova coluna "D" com os termos não repetidos
df['D'] = termos_nao_repetidos

# Salve o DataFrame de volta no arquivo Excel
df.to_excel('indexadores.xlsx', index=False)
