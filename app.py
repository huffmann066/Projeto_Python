import pandas as pd
import pyperclip
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

vendas_instaladas = 'Mensagens.xlsx';

df = pd.read_excel(vendas_instaladas, sheet_name='Sheet1');

numeros = [];

for index, row in df.iterrows():
    mensagem_enviada = row['MENSAGEM ENVIADA'];
    tel_contato = row['TEL CONTATO'];
    
    if mensagem_enviada == 'SIM':
        pass;
    elif pd.isna(mensagem_enviada) or mensagem_enviada == '':
        print(f'Telefone de Contato: {tel_contato}');
        numeros.append(tel_contato);

mensagem = "Mensagem de saudação";
driver = webdriver.Firefox();
driver.get("https://web.whatsapp.com/");
input("Pressione Enter depois de fazer o login no WhatsApp Web...");


for numero in numeros:
    url = f"https://web.whatsapp.com/send?phone={numero}";
    driver.get(url);
    
    time.sleep(20);
    
    pyautogui.moveTo(x= 653, y= 700);
    pyautogui.leftClick();
    pyperclip.copy(mensagem);
    pyautogui.hotkey('ctrl', 'v');
    pyautogui.press('enter');

    
    time.sleep(10);

pyautogui.alert('Pronto');

driver.quit();