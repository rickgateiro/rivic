import pandas as pd

# Carregue o arquivo Excel
df = pd.read_excel('teste2.xlsx')

# Função para separar os termos em uma lista e ignorar os caracteres separadores
def split_and_strip(text):
    if pd.notna(text):
        return [term.strip() for term in text.split(';')]
    return []

# Aplica a função às colunas A e B
df['A'] = df['A'].apply(split_and_strip)
df['B'] = df['B'].apply(split_and_strip)

# Cria a coluna C com os termos em A que não estão em B
df['C'] = df.apply(lambda row: '; '.join([term for term in row['A'] if term not in row['B']]), axis=1)

# Salve o DataFrame de volta no arquivo Excel
df.to_excel('diferencas.xlsx', index=False)
