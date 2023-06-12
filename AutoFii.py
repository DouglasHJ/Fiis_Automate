##Importando Bibliotecas que irei utilizar, neste caso utilizei o Selenium e Pandas para ajustar os dados
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
from datetime import date

##Comandos para abertura do navegador com o link do site
driver = webdriver.Chrome()
driver.get("https://www.fundamentus.com.br/fii_buscaavancada.php")

##Validando se o título do site é o correto
assert "FUNDAMENTUS - Invista consciente" in driver.title

##Encontrando o botão de pesquisa no site e clicando
elem = driver.find_element(By.CLASS_NAME, "buscar")
elem.click()

##Pegando os dados necessários da tabela HTML e colocando nas variaveis
tabela_papel = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[1]")
tabela_segm = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[2]")
tabela_cot = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[3]")
tabela_ffoyld = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[4]")
tabela_divyld = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[5]")
tabela_pvp = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[6]")
tabela_valormrc = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[7]")
tabela_lqdz = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[8]")
tabela_vacanc = driver.find_elements(By.XPATH,"//table[@id='tabelaResultado']/tbody/tr/td[13]")

##Criando uma lista vazia para adicionar os dados coletados
tabelafiis_resultado = []

##Criando laço para unir dados em um dicionário
for i in range(len(tabela_papel)):
    temp_data = {
                'Papel': tabela_papel[i].text,
                'Segmento': tabela_segm[i].text,
                'Cotação': tabela_cot[i].text,
                'FFO Yield': tabela_ffoyld[i].text,
                'Dividend Yield': tabela_divyld[i].text,
                'P/VP': tabela_pvp[i].text,
                'Valor Mercado':(tabela_valormrc[i].text),
                'Liquidez': tabela_lqdz[i].text,
                'Vacância': tabela_vacanc[i].text}
    tabelafiis_resultado.append(temp_data)


#Criando um DataFrame utilizando o Pandas para organizar em uma 'Tabela'
df_data = pd.DataFrame(tabelafiis_resultado)

# Convertendo as colunas relevantes para float
colunas_float = ['Cotação', 'FFO Yield', 'Dividend Yield', 'P/VP', 'Valor Mercado', 'Liquidez', 'Vacância']

for coluna in colunas_float:
    df_data[coluna] = df_data[coluna].apply(lambda x: float(x.replace('.', '').replace(',', '.').replace('%', '')))

##Filtrando a planilha
df_data_filter = df_data.query("Cotação >= 4 and `P/VP` >= 0.04 and `P/VP` <= 1.2 and `Valor Mercado` >= 500000 and Vacância <= 30 and Liquidez >= 1000000")

#Exportando os dados do DataFrame em excel
data_atual = date.today()

nome_arquivo_SF = f'Fiis_Resultado_SemFiltro_{data_atual}.xlsx'
nome_arquivo_F = f'Fiis_Resultado_Filtrado_{data_atual}.xlsx'

df_data.to_excel(nome_arquivo_SF, index=False)
df_data_filter.to_excel(nome_arquivo_F, index=False)

#Encerrando o navegador
driver.close()
