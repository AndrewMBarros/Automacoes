import openpyxl
import pyautogui
import pyperclip
from time import sleep
import time 

# entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook ['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    Nome_do_produto = linha [0].value
    pyperclip.copy(Nome_do_produto)
    pyautogui.click(482,576,duration = 3)
    pyautogui.hotkey('ctrl','v')
    
    
    Descricao = linha [1].value
    pyperclip.copy(Descricao)
    pyautogui.click(504,724,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    #descer a pagina
    pyautogui.click(1435,837,duration = 1)

    
    Categoria = linha [2].value
    pyperclip.copy(Categoria)
    pyautogui.click(480,197,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    
    
    Codigo_do_produto = linha [3].value
    pyperclip.copy(Codigo_do_produto)
    pyautogui.click(491,349,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    
    Peso = linha [4].value
    pyperclip.copy(Peso)
    pyautogui.click(502,501,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    
    Dimensoes = linha [5].value
    pyperclip.copy(Dimensoes)
    pyautogui.click(496,647,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    
    #proxima
    pyautogui.click(458,732,duration = 3)
    


    Preco  = linha [6].value
    pyperclip.copy(Preco)
    pyautogui.click(470,618,duration = 3)
    pyautogui.hotkey('ctrl','v') 
    
    
    
    Quantidade = linha [7].value
    pyperclip.copy(Quantidade)
    pyautogui.click(502,766,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    #descer a pagina
    pyautogui.click(1433,838,duration = 1)
     
    
    
    validade = linha [8].value
    pyperclip.copy(validade)
    pyautogui.click(480,245,duration = 1)
    pyautogui.hotkey('ctrl','v')
    
    
    

    Cor = linha [9].value
    pyperclip.copy(Cor)
    pyautogui.click(458,385,duration = 1)
    pyautogui.hotkey('ctrl','v')  
    
    
    Tamanho = linha [10].value
    pyperclip.copy(Tamanho)
    pyautogui.click(462,535,duration = 1)
    pyautogui.hotkey('ctrl','v')   
    
    
    
    Material = linha [11].value
    pyperclip.copy(Material)
    pyautogui.click(472,691,duration = 1)
    pyautogui.hotkey('ctrl','v')   
    
    
    
    ### proximo 
    pyautogui.click(540,770,duration = 3)

    
    Fabricante = linha [12].value
    pyperclip.copy(Fabricante)
    pyautogui.click(480,618,duration = 3)
    pyautogui.hotkey('ctrl','v')    
    
    #descer a pagina
    pyautogui.click(1431,836,duration = 1)
    
    
    origem = linha [13].value
    pyperclip.copy(origem)
    pyautogui.click(472,166,duration = 1)
    pyautogui.hotkey('ctrl','v')   
    
    
    Observacoes = linha [14].value
    pyperclip.copy(Observacoes)
    pyautogui.click(481,316,duration = 1)
    pyautogui.hotkey('ctrl','v')   
    

    Codigo_de_barras = linha [15].value
    pyperclip.copy(Codigo_de_barras)
    pyautogui.click(486,455,duration = 1)
    pyautogui.hotkey('ctrl','v')   
    

    Localizacao_no_armazem = linha [16].value
    pyperclip.copy(Localizacao_no_armazem)
    pyautogui.click(497,612,duration = 1)
    pyautogui.hotkey('ctrl','v')   
    
    
    ### proximo 
    pyautogui.click(547,688,duration = 1)

    ### reiniciar
    time.sleep(6)
    pyautogui.hotkey('ctrl','r',duration = 6)
