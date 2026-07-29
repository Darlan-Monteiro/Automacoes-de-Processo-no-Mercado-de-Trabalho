# bloco para importar bibliotecas
import pandas as pd
import pathlib
import win32com.client as win32
#importar bases de dados
emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', sep=';', encoding='latin-1')


# adicionando nome das lojas em vendas
vendas = vendas.merge(lojas, on='ID Loja')

# criei um dicionário de lojas
dict_lojas = {}
for loja in lojas['Loja']:
    dict_lojas[loja] = vendas.loc[vendas['Loja']== loja,:]

# para pegar a data mais recente da base de vendas
dia_indicador = vendas['Data'].max()

# identificando se as pastas já existem
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

# verificando se a pasta existe, caso não exista, cria a pasta
aquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in aquivos_pasta_backup]

for loja in dict_lojas:
    if loja not in lista_nomes_backup:
        caminho_backup_loja = caminho_backup / loja
        caminho_backup_loja.mkdir()
        
    # criando o nome do arquivo e salvando o arquivo na pasta da loja
    nome_arquivo = '{}_{}_{} - {}.xlsx'.format(dia_indicador.day, dia_indicador.month, dia_indicador.year, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dict_lojas[loja].to_excel(local_arquivo)
    
loja = 'Norte Shopping'
vendas_loja = dict_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

faturamento_ano = vendas_loja['Valor Final'].sum()
faturamento_dia = vendas_loja_dia['Valor Final'].sum()

qtde_produtos_ano = len(vendas_loja['Produto'].unique())
qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

valor_vendas = vendas_loja.groupby('Código Venda').sum()
ticket_media_ano = valor_vendas['Valor Final'].mean()
valor_vendas_dia = vendas_loja_dia.groupby('Código Venda').sum()
ticket_media_dia = valor_vendas_dia['Valor Final'].mean()


#definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

outlook = win32.Dispatch('outlook.application')
nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
mail.Body = 'Texto do E-mail'



attachment  = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
mail.Attachments.Add(str(attachment))