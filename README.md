# Automação de Indicadores de Rede Varejista (One Page)

## 📌 Objetivo do Projeto
Este projeto consiste em uma automação desenvolvida em Python para calcular e enviar relatórios diários de indicadores (conhecidos como *One Pages*) para os gerentes de 25 lojas de uma grande rede de varejo, além de um relatório consolidado para a diretoria. 

O objetivo principal é eliminar o trabalho manual de análise de dados, segmentação de planilhas e envio de e-mails, transformando um processo demorado em uma execução rápida e 100% automatizada.

---

## ⚙️ Funcionalidades

*   **Processamento de Dados:** Lê e cruza dados de vendas diárias, histórico anual e cadastro de lojas/gerentes.
*   **Cálculo de Indicadores:** Calcula três métricas principais para cada loja (no Dia e no Ano corrente):
    *   Faturamento (Metas: Dia R$ 1.000,00 | Ano R$ 1.650.000,00)
    *   Diversidade de Produtos Vendidos (Metas: Dia 4 | Ano 120)
    *   Ticket Médio por Venda (Metas: Dia R$ 500,00 | Ano R$ 500,00)
*   **Geração de Backups:** Cria automaticamente pastas individuais para cada loja e salva um arquivo Excel apenas com os dados correspondentes àquela unidade.
*   **Envio de E-mails para Gerentes:** Integração com o Microsoft Outlook para enviar um e-mail formatado em HTML com o resumo dos indicadores (sinalizando metas batidas ou perdidas com indicadores visuais) e a planilha de backup em anexo.
*   **Relatório da Diretoria:** Calcula e exporta rankings de faturamento (Diário e Anual) de todas as lojas e envia para a diretoria, destacando no corpo do e-mail a melhor e a pior loja de cada período.

---

## 🛠️ Tecnologias Utilizadas

*   **Python:** Linguagem base do projeto.
*   **Pandas:** Para leitura, manipulação, cruzamento e análise dos dados (arquivos CSV e Excel).
*   **Pathlib:** Para manipulação de diretórios, criação de pastas e caminhos de arquivos, garantindo compatibilidade no Windows.
*   **Pywin32 (`win32com.client`):** Para integração do Python com o Microsoft Outlook (protocolo COM), permitindo a automação do envio de e-mails.

---

## 📂 Estrutura de Arquivos Necessária

Para que o script funcione corretamente, o projeto exige uma pasta raiz contendo o arquivo principal `main.py` e uma subpasta `Bases de Dados` com os seguintes arquivos:

*   `Emails.xlsx`: Contém o nome da loja, nome do gerente e o e-mail de destino (incluindo o contato da Diretoria).
*   `Vendas.xlsx`: Base de dados completa com todas as vendas registradas (Data, Loja, Produto, Valor, Código da Venda).
*   `Lojas.csv`: Tabela auxiliar contendo o ID e o Nome de cada loja.

O script criará automaticamente a pasta `Backup Arquivos Lojas` para armazenar os relatórios gerados.

---

## 🚀 Como Executar

1.  Clone este repositório para a sua máquina local.
2.  Certifique-se de ter o Python instalado.
3.  Instale as bibliotecas necessárias rodando o comando no terminal:
    ```bash
    pip install pandas openpyxl pywin32
    ```
4.  **Requisito Importante:** Por utilizar a biblioteca `pywin32`, este projeto deve ser executado em um ambiente **Windows** e requer que o aplicativo desktop do **Microsoft Outlook** esteja instalado, configurado com uma conta de e-mail ativa e aberto durante a execução.
5.  Execute o arquivo `main.py`.

*(Nota: Para fins de teste, recomenda-se alterar os endereços na base `Emails.xlsx` para o seu e-mail pessoal, evitando o disparo acidental para contatos reais).*

---
> Projeto desenvolvido como aplicação prática de automação de processos e análise de dados com Python.
