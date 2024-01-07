#selenium para o site
#openpyxl para ler as tabelas
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Verificar se o arquivo existe


if os.path.exists('produtos.xlsx'):
    # Se existir, apagar o arquivo
    os.remove('produtos.xlsx')
    print("Arquivo 'produtos.xlsx' existente removido com sucesso!")
    
driver = webdriver.Chrome()
driver.get('https://www.amazon.com.br/gp/bestsellers/computers/16364756011?ref_=Oct_d_obs_S&pd_rd_w=0rJIH&content-id=amzn1.sym.ddac7217-197b-4ce9-8f99-a40c3c6e5f1b&pf_rd_p=ddac7217-197b-4ce9-8f99-a40c3c6e5f1b&pf_rd_r=GHXMEA6E079QX0R5PVN5&pd_rd_wg=lsQPP&pd_rd_r=a1b5bb92-3a1c-44f8-b92f-3afe683a12a5')
#input()
driver.close()
#se desejar o site nao feche basta colocar
#input()
#nome do produto
titulos = driver.find_elements(By.XPATH,"//div[@class ='_cDEzb_p13n-sc-css-line-clamp-3_g3dy1']")

#precos 
precos = driver.find_elements(By.XPATH,'//span[@class = "_cDEzb_p13n-sc-price_3mJ9Z"]')

#caso exista produto que nao tenha preço utiliza-se o zip 

#criar a planilhas
workbook = openpyxl.Workbook()

#para criar uma pagina (sheet) produtos na planilha
workbook.create_sheet('produtos')
#selecionar a pagina que deseja
sheet_produtos = workbook ['produtos']

#inserir a coluna exata nela
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'
#salvar na planilha

for titulo, preco in zip (titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])
#.text pq deseja apenas a propriedade texto daquele elemento
 
workbook.save('produtos.xlsx')
