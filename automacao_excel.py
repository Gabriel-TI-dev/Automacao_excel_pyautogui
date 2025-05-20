#Produtos a serem digitados no excel de forma automatizada
import pyautogui
import time
import pandas as pd

#mensagem alertando o inicio do código
pyautogui.alert("O código vai começar a rodar, não mexa no pc enquanto o código não terminar de rodar.")

#pausa a cada codigo executado
pyautogui.PAUSE = 8


#Passo 1-> Abrir o Excel
pyautogui.press("win")
pyautogui.write("excel")
pyautogui.press("enter")

#tempo para abrir o excel
time.sleep(10)


#abrir a planilha para importar os dados
pyautogui.click(x=80, y=399)
pyautogui.click(x=589, y=265)


#Passo 2-> Importar a base de produtos
planilha = pd.read_csv("produtos.csv")
print(planilha)


#Passo 3-> Digitalizar os produtos na planilha

pyautogui.click(x=70, y=258) #clicar no 1º campo da planilha


for coluna in planilha.index:


    #codigo
    codigo = planilha.loc[coluna, "codigo"]
    pyautogui.write(str(codigo))
    pyautogui.press("tab")

    #marca
    marca = planilha.loc[coluna, "marca"]
    pyautogui.write(str(marca))
    pyautogui.press("tab")

    #tipo
    tipo = planilha.loc[coluna, "tipo"]
    pyautogui.write(str(tipo))
    pyautogui.press("tab")


    #categoria
    categoria = planilha.loc[coluna, "categoria"]
    pyautogui.write(str(categoria))
    pyautogui.press("tab")


    #preco_unitario
    preco_unitario = planilha.loc[coluna, "preco_unitario"]
    pyautogui.write(str(preco_unitario))
    pyautogui.press("tab")


    #custo
    custo = planilha.loc[coluna, "custo"]
    pyautogui.write(str(custo))
    pyautogui.press("tab")


    #obs
    obs = str(planilha.loc[coluna, "obs"])
    if obs != "nan":
        pyautogui.write(obs)

    pyautogui.press("enter")

#mensagem de aviso de termino do codigo
pyautogui.alert("O código foi finalizado, pode voltar a mexer no pc normalmente.")


#Passo 4-> repetir o passo 3 ate o fim da lista

#comentario:
#automação para preencher lista de produtos no excel, se tiverem sugestões de como melhorar agradeço, mas por enquanto é isso,
#e é praticando que se aprende.