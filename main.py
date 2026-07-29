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

# pra pegar a data mais recente da base de vendas
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
    
vendas_loja = dict_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

faturamento_ano = vendas_loja['Valor Final'].sum()
faturamento_dia = vendas_loja_dia['Valor Final'].sum()

qtde_produtos_ano = len(vendas_loja['Produto'].unique())
qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

valor_vendas = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
ticket_media_ano = valor_vendas['Valor Final'].mean()
valor_vendas_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
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
mail.Subject = f'Indicadores: Dia {dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year} - Loja {loja}'


# bloco para enviar e-mail para cada loja

for loja in dict_lojas:

    vendas_loja = dict_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # calculando valor da venda e ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    
    
    # bloco para enviar e-mail para cada loja
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f'Indicadores: Dia {dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year} - Loja {loja}'

    # bloco para definir cores dos indicadores
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    # bloco para enviar e-mail para cada loja
    mail.HTMLBody = f'''
    <p>Boa tarde, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Darlan</p>
    '''


    # bloco para anexar o arquivo da loja
    nome_arquivo_anexo = f'{dia_indicador.day}_{dia_indicador.month}_{dia_indicador.year} - {loja}.xlsx'

    attachment = pathlib.Path.cwd() / caminho_backup / loja / nome_arquivo_anexo
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('E-mail da Loja {} enviado'.format(loja))
    
    
# criação do ranking anual
faturamento_lojas = vendas.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

# criação do ranking diário
vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

#enviar o e-mail
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, boa tarde

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Darlan
'''


# Anexos
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))
mail.Send()

print('E-mail da Diretoria enviado')