from playwright.sync_api import sync_playwright
import time
import openpyxl
import pandas as pd


files = input("digite o nome: ")


df = pd.read_excel(files)
book = openpyxl.Workbook()
book.create_sheet('Planilha')


for index,row in df.iterrows():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://gss.aegea.com.br/ords/riodejaneiro04/aegea/r/gss101/ate3010?session=12526425361445")
        #clicar para pesquisar
        page.locator('xpath=//*[@id="P384_NUM_LIGACAO_lov_btn"]/span').click

        # preencher input da pesquisa
        # page.fill('xpath= //*[@id="PopupLov_384_P384_NUM_LIGACAO_dlg"]/div[1]/input', f'{row["Matrícula"]}')

        #Botão de pesquisa (enter)
        #page.locator("xpath=//*[@id="PopupLov_384_P384_NUM_LIGACAO_dlg"]/div[1]/button/span").click

        page.keyboard.press("Enter")

        #Resultado da pesquisa
        page.locator('xpath=//*[@id="PopupLov_384_P384_NUM_LIGACAO_dlg"]/div[2]/div[2]/div[4]/table/tbody/tr').click

        #lançamento
        page.locator('xpath= //*[@id="CardConsulta"]/ul/li[9]/div').click
        
        full_price = page.locator('//*[@id="P454_VL_TOTAL"]').text_content()

        # phone = page.locator('xpath=//*[@id="rso"]/div[2]/div/div/div[1]/div/div/span/a/div/div/div/cite').text_content()
        lista = []
        lista.append(full_price)

        # print(lista)
        for i in lista:

            if len(i) < 0:
                page_excel= book['Planilha']
                page_excel.append([row['Matrícula'], i])
                book.save("planilha_sem_debito.xlsx")

                print(f" SEM DÉBITO: {len(i)} = {i}")

            elif len(i) >= 1:            
                page_excel = book['Planilha']
                page_excel.append([row['Matrícula'], i])
                book.save("planilha_com_debito.xlsx")

                print(f" EM DÉBITO: {len(i)} = {i}")
            else:
                print("ERRO NO SISTEMA, REINICIE E TENTE NOVAMENTE")
                    
        
        time.sleep(3)
        print(page.title())

         