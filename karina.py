from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import pandas as pd
import openpyxl

browser = webdriver.Chrome()

browser.maximize_window()
browser.get(f"http://www.abccmm.org.br/animais")
time.sleep(15)

tabela = pd.read_excel("karina_insuportavel.xlsx")

coluna_lista = tabela['ANIMAL'].tolist()
time.sleep(1)

tabela = tabela.drop(tabela.index)
for item in coluna_lista:
    try:
        pesquisa = browser.find_element(By.ID, "ContentPlaceHolder1_TxtNomeAnimal")
        pesquisa.click()
        pesquisa.clear()
        pesquisa.send_keys(item)
        btnPesquisa = browser.find_element(By.ID, "ContentPlaceHolder1_BtnPesquisar")
        btnPesquisa.click()
        time.sleep(2)
        btnVer = browser.find_element(By.ID, "ContentPlaceHolder1_GridAnimais_LkbEditar_0")
        btnVer.click()
        time.sleep(2)

        pai = browser.find_element(By.ID, "ContentPlaceHolder1_lblPai_NomePai_Genealogia").text
        mae = browser.find_element(By.ID, "ContentPlaceHolder1_lblMae_NomeMae_Genealogia").text
        avoPaterno = browser.find_element(By.ID, "ContentPlaceHolder1_lblPai_NomePaiPai_Genealogia").text
        avoPaterna = browser.find_element(By.ID, "ContentPlaceHolder1_lblPai_NomePaiMae_Genealogia").text
        avoMaterno = browser.find_element(By.ID, "ContentPlaceHolder1_lblMae_NomePaiMae_Genealogia").text
        avoMaterna = browser.find_element(By.ID, "ContentPlaceHolder1_lblMae_NomeMaeMae_Genealogia").text
        tipo = browser.find_element(By.ID, "ContentPlaceHolder1_LblSexo").text
        dataNasc = browser.find_element(By.ID, "ContentPlaceHolder1_LblDataNascimento").text
        tabela = tabela._append({'ANIMAL': item, 'PAI': pai, 'MAE': mae,
                                'AVO PATERNO': avoPaterno, 'AVO PATERNA': avoPaterna,
                                'AVO MATERNO': avoMaterno, 'AVO MATERNA': avoMaterna,
                                'TIPO': tipo, 'DATA DE NASCIMENTO': dataNasc}, ignore_index=True)
    except:
        print(item)

tabela.to_excel("resultado_karinaChatona.xlsx", index=False)
time.sleep(1)


