from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime
import os
import pytz

timezone = pytz.timezone('America/Sao_Paulo')

# Diretório de download para GitHub Actions
download_dir = "/tmp"
os.makedirs(download_dir, exist_ok=True)

def login(page):
    page.goto("https://spx.shopee.com.br/")
    page.wait_for_selector('xpath=//*[@placeholder="Ops ID"]', timeout=15000)
    page.fill('xpath=//*[@placeholder="Ops ID"]', 'Ops322349')
    page.fill('xpath=//*[@placeholder="Senha"]', '@Shopee123')
    page.click('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button')

    page.wait_for_timeout(15000)
    try:
        page.click('css=.ssc-dialog-close', timeout=5000)
    except:
        print("Nenhum pop-up foi encontrado.")
        page.keyboard.press("Escape")

def get_data(page):
    data = []
    try:
        # Primeiro link
        page.goto("https://spx.shopee.com.br/#/dashboard/facility-soc/historical-data")
        page.wait_for_timeout(10000)
        first_value = page.inner_text('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[25]')
        data.append(first_value)

        # Segundo link
        page.goto("https://spx.shopee.com.br/#/dashboard/toProductivity?page_type=Outbound")
        page.wait_for_timeout(10000)
        page.wait_for_selector('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div', timeout=30000)
        second_value = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div')
        data.append(second_value)

        # Terceiro dado
        page.click('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div/div/div/div[3]')
        page.wait_for_timeout(10000)
        page.wait_for_selector('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div', timeout=30000)
        third_value = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div')
        data.append(third_value)

        # Quarto dado
        page.click('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div/div/div/div[4]')
        page.wait_for_timeout(10000)
        page.wait_for_selector('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div', timeout=30000)
        page.wait_for_timeout(20000)
        fourth_value = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div')
        data.append(fourth_value)

    except Exception as e:
        print(f"Erro ao coletar dados: {e}")
        raise
    return data

def update_google_sheets(data):
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1_ySESJhetl_zVvRB1azj-7FjaldHFOQK_DKDQYhNiyc/edit').worksheet("Reporte HxH")

    current_time = datetime.datetime.now(timezone)
    hour = current_time.hour
    
    # 1. Mapeamento da linha anterior
    if 7 <= hour <= 23:
        row_anterior = hour - 5
    elif hour == 0:
        row_anterior = 19
    elif 1 <= hour <= 6:
        row_anterior = hour + 19
    else:
        print(f"Hora fora do intervalo: {hour}:{current_time.minute}")
        return

    # 2. Mapeamento da linha atual
    row_atual = row_anterior + 1
    if row_atual > 25:
        row_atual = 2 # Volta para 06:00

    # 3. Tratamento de Conflito com a Limpeza das 06:00
    if hour == 6:
        # Às 06:xx, o dia virou e a planilha foi limpa.
        # Atualizamos APENAS a linha atual (06h) para não reescrever a linha 25 (05h) apagada.
        rows_to_update = [row_atual]
        print("Nova diária detectada (06:xx). Ignorando linha 25 para evitar sujeira de dados.")
    else:
        rows_to_update = [row_anterior, row_atual]
    
    # 4. Execução da atualização
    for row in rows_to_update:
        if 2 <= row <= 25:
            cell_range = f'B{row}:E{row}'
            sheet.update(values=[data], range_name=cell_range)
            print(f"Linha {row} ({cell_range}) atualizada com sucesso.")

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            login(page)
            data = get_data(page)
            update_google_sheets(data)
            print("Processo concluído.")
        except Exception as e:
            print(f"Erro: {e}")
        finally:
            browser.close()

if __name__ == "__main__":
    main()
