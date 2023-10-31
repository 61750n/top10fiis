# PROGRAMA QUE IMPORTA A TABELA DE DADOS DE FIIS DO SITE FUNDAMENTUS, TRATA E CLASSIFICA AS EMPRESAS E EXPORTA PARA EXCEL
 
# 1) INSTALAÇÕES INICIAIS NO TERMINAL
# pip install pandas
# pip install selenium
# pip install webdriver-manager
# pip install lxml
# pip install datetime

# 2) IMPORTAÇÕES DAS BIBLIOTECAS
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import datetime

# 3) ABRE NAVEGADOR --> ACESSA SITE
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

url = 'https://www.fundamentus.com.br/fii_resultado.php'

driver.get(url)

# 4) LÊ A TABELA DE AÇÕES --> FECHA SITE
local_tabela = '/html/body/div[1]/div[2]/table'

elemento = driver.find_element("xpath", local_tabela)

html_tabela = elemento.get_attribute('outerHTML')

tabela = pd.read_html(str(html_tabela), thousands = '.', decimal = ',')[0]

# EXTRA: EXIBE TABELA IMPORTADA PARA VERIFICAÇÃO SE A IMPORTAÇÃO E LEITURA DERAM CERTO
#print(tabela)

#EXTRA: EXIBE O TIPO DE DADO QUE TEM NAS COLUNAS DA TABELA. INT É NÚMERO INTEIRO, FLOAT64 É NÚMERO DECIMAL, 
#tabela.info()

# 5) TRATAMENTO DO TIPO DE DADO DAS COLUNAS. DE OBJECT PARA FLOAT
#5.1) COLUNA Dividend Yield
tabela['Dividend Yield'] = tabela['Dividend Yield'].str.replace("%", "")
tabela['Dividend Yield'] = tabela['Dividend Yield'].str.replace(".", "")
tabela['Dividend Yield'] = tabela['Dividend Yield'].str.replace(",", ".")
tabela['Dividend Yield'] = tabela['Dividend Yield'].astype(float)

#5.2) COLUNA FFO Yield
tabela['FFO Yield'] = tabela['FFO Yield'].str.replace("%", "")
tabela['FFO Yield'] = tabela['FFO Yield'].str.replace(".", "")
tabela['FFO Yield'] = tabela['FFO Yield'].str.replace(",", ".")
tabela['FFO Yield'] = tabela['FFO Yield'].astype(float)

#5.3) COLUNA Mrg Ebit
tabela['Vacância Média'] = tabela['Vacância Média'].str.replace("%", "")
tabela['Vacância Média'] = tabela['Vacância Média'].str.replace(".", "")
tabela['Vacância Média'] = tabela['Vacância Média'].str.replace(",", ".")
tabela['Vacância Média'] = tabela['Vacância Média'].astype(float)

#5.4) COLUNA Cap Rate
tabela['Cap Rate'] = tabela['Cap Rate'].str.replace("%", "")
tabela['Cap Rate'] = tabela['Cap Rate'].str.replace(".", "")
tabela['Cap Rate'] = tabela['Cap Rate'].str.replace(",", ".")
tabela['Cap Rate'] = tabela['Cap Rate'].astype(float)

#6) ELIMINANDO EMPRESAS COM LIQUIDEZ DIÁRIA ABAIXO DE R$500.000,00. 
# ISTO SIGNIFICA QUE É RÁPIDO E FÁCIL COMPRAR E VENDER ESSAS AÇÕES
tabela = tabela[tabela['Liquidez'] > 500000]

#EXTRA: EXIBE TABELA IMPORTADA PARA VERIFICAÇÃO SE A IMPORTAÇÃO E LEITURA DERAM CERTO
#print(tabela)

#EXTRA: EXIBE O TIPO DE DADO QUE TEM NAS COLUNAS DA TABELA. INT É NÚMERO INTEIRO, FLOAT64 É NÚMERO DECIMAL, 
#tabela.info()

#7) ELIMINANDO EMPRESAS COM Vacância Média ABAIXO DE 21%
tabela = tabela[tabela['Vacância Média'] < 20]

#8) ELIMINANDO EMPRESAS COM P/VP MAIOR QUE 1
tabela = tabela[tabela['P/VP'] < 1]

#9) CLASSIFICA TABELA PELO DIVIDENDO EM ORDEM DECRESCENTE (MAIOR PARA O MENOR)
tabela['ranking_dividendos'] = tabela['Dividend Yield'].rank(ascending = False)

#10) CLASSIFICA TABELA PELO FFO Yield EM ORDEM DECRESCENTE (MAIOR PARA O MENOR)
tabela['ranking_ffo'] = tabela['FFO Yield'].rank(ascending = False)

#11) CONSTRUÇÃO DO RANKING
tabela['ranking_total'] = tabela['ranking_dividendos'] + tabela['ranking_ffo']

tabela = tabela.sort_values('ranking_total')

#12) EXIBE AS 10 MELHORES AÇÕES:
print(tabela.head(10))

#13) OBTÉM A DATA ATUAL
data_atual = datetime.datetime.today()

#14) PREPARA O NOME DO ARQUIVO
nome_arquivo = f"top_10_fiis_{data_atual.strftime('%Y-%m-%d')}.xlsx"

#15) EXPORTA PARA EXCEL
tabela.to_excel(f'{nome_arquivo}', index=False)
