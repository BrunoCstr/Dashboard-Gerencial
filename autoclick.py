import pyautogui
import time
import ctypes

time.sleep(8)

def click_hidden(x, y):
    off_screen_position = (pyautogui.size().width - 10, pyautogui.size().height - 10)
    
    pyautogui.moveTo(x, y)
    pyautogui.click()
    pyautogui.moveTo(*off_screen_position)

def pagination():
    print("Executando cliques...")
    
    print("Comissão e Faturamento")
    click_hidden(103, 187)  # Comissão e Faturamento
    time.sleep(10)
    
    print("Negócios Totais")
    click_hidden(91, 221)  # Negócios Totais
    time.sleep(10)
    
    print("Rankings Unidades")
    click_hidden(88, 270)  # Rankings Unidades
    time.sleep(10)
    
    print("Base Clientes Qualitativo")
    click_hidden(1019, 214)  # Base Clientes Qualitativo
    time.sleep(10)
    
    print("Índice de Renovação")
    click_hidden(1008, 228)  # Índice de Renovação
    time.sleep(10)
    
    print("Novos Negócios Qualitativo")
    click_hidden(1032, 243)  # Novos Negócios Qualitativo
    time.sleep(10)
    
    print("Novos Negócios Quantitativo")
    click_hidden(1018 , 260)  # Novos Negócios Quantitativo
    time.sleep(10)
    
    print("Média de Comissão")
    click_hidden(1021 , 284)  # Média de Comissão
    time.sleep(10)
    
    print("Crescimento de Comissão")
    click_hidden(1028 , 302)  # Crescimento de Comissão
    time.sleep(10)
    
    print("Taxas e Médias")
    click_hidden(101 , 297)  # Taxas e Médias
    time.sleep(10)
    
    print("Centro de Op.")
    click_hidden(102 , 342)  # Centro de Op.
    time.sleep(10)
    
    print("Fornecedores e Produtos")
    click_hidden(109 , 381)  # Fornecedores e Produtos
    time.sleep(10)
    
    print("Expansão")
    click_hidden(97 , 428)  # Expansão
    time.sleep(10)
    
while True:
    pagination()