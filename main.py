#Site para teste 
#https://cadastro-produtos-devaprender.netlify.app/etapa2.html

import openpyxl #Abrir a planilha no excel
import pyperclip #Uma biblioteca para copiar e colar com isso mantendo a escrtia correta 
import pyautogui
from time import sleep
# Carregar a planilha
planilha_produtos = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = planilha_produtos['Produtos']

# Iterar sobre as linhas da planilha a partir da segunda linha
for linha in sheet_produtos.iter_rows(min_row=2):
    nome = linha[0].value 
    pyperclip.copy(nome)
    pyautogui.click(1160,298,duration=1)
    pyautogui.hotkey('ctrl','v')    

    descricao = linha[1].value 
    pyperclip.copy(descricao)
    pyautogui.click(1236,425,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1210,565,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_do_produto = linha[3].value
    pyperclip.copy(codigo_do_produto)
    pyautogui.click(1202,673,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1170,795,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1213,895,duration=1)
    pyautogui.hotkey('ctrl','v')


    pyautogui.click(1188,964,duration=1)
    sleep(1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1179,345,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1179,447,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1191,561,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1190,659,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    tamanho = linha[10].value
    pyautogui.click(1186,769,duration=1)
    if(tamanho == 'Pequeno'): 
        pyautogui.click(1246,814,duration=1)
    elif(tamanho == 'Mediano'):
        pyautogui.click(1237,829,duration=1)
    elif(tamanho == 'Grande'):
        pyautogui.click(1301,868,duration=1)
 
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1230,881,duration=1)
    pyautogui.hotkey('ctrl','v')


    pyautogui.click(1189,954,duration=1)
    sleep(1)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1199,362,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(1241,474,duration=1)
    pyautogui.hotkey('ctrl','v')
     
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1191,587,duration=1)
    pyautogui.hotkey('ctrl','v')
     
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1169,746,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    localizacao_no_armazem = linha[16].value
    pyperclip.copy(localizacao_no_armazem)
    pyautogui.click(1185,855,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1174,928,duration=1)
    pyautogui.click(1678,228,duration=1)
    pyautogui.click(1422,647,duration=1)



    

   









