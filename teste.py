import xml.etree.ElementTree as ET
import pandas as pd

# Carregando o arquivo XML (ajuste o caminho se necessário)
tree = ET.parse('bilhetagem.xml')
root = tree.getroot()

# Extraindo os dados
data = []
for op in root.findall('OP'):
    n = op.find('N').text
    t = op.find('T')
    linha = t.find('L').text
    inicio = t.find('I').text
    fim = t.find('F').text
    duracao = t.find('D').text
    ci = t.find('CI').text
    cf = t.find('CF').text
    psg = t.find('PSG').text
    p = t.find('P').text
    ta = t.find('TA').text
    of = t.find('OF').text
    data.append([n, linha, inicio, fim, duracao, ci, cf, psg, p, ta, of])

# Criando o DataFrame
df = pd.DataFrame(data, columns=['Número da Operação', 'Linha', 'Início', 'Fim', 'Duração', 'Código Inicial', 'Código Final', 'PSG', 'P', 'TA', 'Operador'])

df[['numero', 'nome', 'mat.Cob']] = df['Número da Operação'].str.split('-', expand=True)
df[['numero', 'nome', 'mat.Mot']] = df['Operador'].str.split('-', expand=True)
df[['Data', 'h.Inicial']] = df['Início'].str.split(' ', expand=True)
df[['Data2', 'h.Final']] = df['Fim'].str.split(' ', expand=True)

# df.drop(['Número da Operação', 'numero', 'nome', 'Operador','TA', 'PSG', 'Duração', 'Código Inicial', 'Código Final', 'Início'], axis=1, inplace=True)

novaOrdem = ['mat.Cob', 'Linha', 'Data', 'h.Inicial', 'h.Final', 'P', 'mat.Mot']
df = df.reindex(columns=novaOrdem)

# Salvando em um arquivo Excel
df.to_excel('resultado.xlsx', index=False)