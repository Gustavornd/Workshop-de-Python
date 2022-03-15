
import pyautogui;
import pyperclip;
import pandas as pd;
import time;

pyautogui.PAUSE = 1


pyautogui.press("win")
pyautogui.write("Opera")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")


time.sleep(3)
pyautogui.click(x=424, y=302, clicks=2)
time.sleep(2)
pyautogui.click(x=424, y=302)
pyautogui.click(x=1708, y=189)
pyautogui.click(x=1524, y=589)
time.sleep(5)
# pyautogui.click(x=1899,y=6)


tabela = pd.read_excel(r"C:\Users\gusta\Downloads\Vendas - Dez.xlsx") 
    #pode colocar o parametro sheets para indicar o numero da aba ou o nome da aba
    #sheets = 1 ou sheets = "Aba 1"
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)


# pyautogui.press("win")
# pyautogui.write("Opera")
# pyautogui.press("enter")
pyautogui.hotkey("ctrl","t")
pyperclip.copy("gmail.com")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(5)
pyautogui.click(x=145,y=198)
time.sleep(2)
pyautogui.write("dinilvallen@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Um email de teste pra tu que é muito feia.")
pyautogui.hotkey("ctrl","v")

pyautogui.press("tab")
texto = f"""
Prezada, isso é um email teste.

O faturamento de ontem foi de R$: {faturamento:,.2f}
A quantidade de produtos vendidos foi de: {quantidade:,}

Abraços,
Eu mesma.
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

pyautogui.hotkey("ctrl","enter")

'''
time.sleep(5)
print(pyautogui.position())
time.sleep(5)
print(pyautogui.position())
time.sleep(5)
print(pyautogui.position())
time.sleep(5)
print(pyautogui.position())
time.sleep(5)
print(pyautogui.position())'''

# Point(x=145, y=198)
# Point(x=1385, y=556)
# Point(x=1306, y=1035)
# Point(x=1908, y=12)
# Point(x=1908, y=12)