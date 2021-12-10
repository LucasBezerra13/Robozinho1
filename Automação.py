import pyautogui
import pyperclip
import time
import pandas as pd

#1

pyautogui.PAUSE = 2.7
pyautogui.alert("O código vai começãr! Afastese do mouse e do teclado.")
time.sleep(5)

#entrar no navegador
pyautogui.press("win")
pyautogui.write("Chrome")
pyautogui.press("enter")

#dentro do navegador chegar até a página de email
time.sleep(6)
#print(pyautogui.position())
pyautogui.moveTo(x=847, y=46)
pyautogui.click()
pyautogui.write("https://www.google.com/intl/pt-BR/drive/")
pyautogui.press("enter")

#entrar no Google drive
time.sleep(5)
button1= pyautogui.locateOnScreen('b.png')
#print(button1)
pyautogui.moveTo(button1)
pyautogui.click()

#fazer login na página de email
time.sleep(3)
button2= pyautogui.locateOnScreen('l.png')
#print(button2)
pyautogui.moveTo(button2)
pyautogui.click(button2)
time.sleep(3)
pyautogui.write("XXXXXXXXX")
pyautogui.press("enter")

#localizar pasta Base de Dados
time.sleep(7)
button3=pyautogui.locateOnScreen('a.png')
#print(button3)
pyautogui.moveTo(button3)
pyautogui.doubleClick(button3)

#Download da base de dados
time.sleep(3)
#print(pyautogui.position())
pyautogui.moveTo(x=290, y=303)
pyautogui.rightClick()

time.sleep(3)
download=pyautogui.locateOnScreen('do.png')
#print(download)
pyautogui.moveTo(download)
#pyautogui.click()

"""
#2

#Análise de dados 

#Chamando a Tabela no pycharmy
tabela= pd.read_excel('Vendas.xlsx')
#print(tabela)

#Panorarama geral sobre a sua análise de dados
faturamento_total= tabela["Valor Final"].sum()
#print(faturamento_total)


#Faturamento por loja
faturamento_por_loja=tabela[["ID Loja","Valor Final"]].groupby("ID Loja").sum()
#print(faturamento_por_loja)

#Faturamento por produto
faturamento_por_produto=tabela[["ID Loja","Produto","Valor Final"]].groupby(["ID Loja","Produto"]).sum()
#print(faturamento_por_produto)

#Produto mais vendido
produto_mais_vendido=tabela[["ID Loja","Produto","Valor Final"]].groupby(["ID Loja","Produto"]).sum().agg(max)
print(produto_mais_vendido)

#Produto menos vendido
produto_menos_vendido=tabela[["ID Loja","Produto","Valor Final"]].groupby(["ID Loja","Produto"]).sum().agg(min)
print(produto_menos_vendido)
"""

#3

#Localizar Gmail
time.sleep(5)
quad=pyautogui.locateOnScreen('quad.png')
#print(quad)
pyautogui.moveTo(quad)
pyautogui.click()
##
time.sleep(5)
gm=pyautogui.locateOnScreen('gm.png')
#print(gm)
pyautogui.moveTo(gm)
pyautogui.click()


#Escrever o email.
time.sleep(6)
#print(pyautogui.position())
pyautogui.moveTo(x=39, y=173)
pyautogui.click()
time.sleep(5)
pyautogui.moveTo(x=881, y=306)
pyautogui.click()
pyautogui.write("presidente@ggmail.com ")
pyautogui.write("vicepresidente@ggmail.com ")
pyautogui.write("gestor@ggmail.com ")
time.sleep(5)
#print(pyautogui.position())
pyautogui.moveTo(x=872, y=374)
pyautogui.click()
pyautogui.write("Resultado do mes de Dezembro")
time.sleep(4)
print(pyautogui.position())
pyautogui.moveTo(x=872, y=374)
pyautogui.click()


"""
pyautogui.write(format("Bom dia ,"
                "segue resultado do mês de dezembro, onde o produto mais  vendido foi " {produto_mais_vendido}", e o produto menos vendido foi " {produto_menos_vendido})
"""

pyautogui.write("Bom dia,segue resultado do mes de dezembro, onde o produto mais  vendido foi {produto_mais_vendido}, e o produto menos vendido foi {produto_menos_vendido}")
time.sleep(4)
#print(pyautogui.position())
pyautogui.moveTo(x=830, y=696)
#pyautogui.click()
pyautogui.alert("Fim do código!")