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
    print("🔑 Iniciando login...")
    page.goto("https://spx.shopee.com.br/")
    page.wait_for_selector('xpath=//*[@placeholder="Ops ID"]', timeout=15000)
    page.fill('xpath=//*[@placeholder="Ops ID"]', 'Ops322349')
    page.fill('xpath=//*[@placeholder="Senha"]', '@Shopee123')
    page.click('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button')

    page.wait_for_timeout(15000)
    try:
        page.click('css=.ssc-dialog-close', timeout=5000)
    except:
        page.keyboard.press("Escape")

def get_dual_data(page):
    data_atual = []
    data_anterior = []
    
    print("\n📊 Coletando dados via XPaths Diretos (1ª e 2ª Coluna)...")
    try:
        # 1. INBOUND (Historical Data)
        page.goto("https://spx.shopee.com.br/#/dashboard/facility-soc/historical-data")
        page.wait_for_timeout(10000)
        # Nota: Mantive td[25] para a Atual (seu original) e td[26] para Anterior. 
        # Se o Inbound crescer para o outro lado, basta inverter para td[24].
        inbound_atual = page.inner_text('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[25]')
        inbound_ant = page.inner_text('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[26]')
        data_atual.append(inbound_atual)
        data_anterior.append(inbound_ant)
        print("   [OK] Inbound")

        # 2. PACKING (Productivity)
        page.goto("https://spx.shopee.com.br/#/dashboard/toProductivity?page_type=Outbound")
        page.wait_for_timeout(10000)
        # th[4] = 1ª coluna de hora (Atual) | th[5] = 2ª coluna de hora (Anterior)
        packing_atual = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div')
        packing_ant = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[5]/div/div')
        data_atual.append(packing_atual)
        data_anterior.append(packing_ant)
        print("   [OK] Packing")

        # 3. ASSIGNMENT
        page.click('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div/div/div/div[3]')
        page.wait_for_timeout(10000)
        assign_atual = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div')
        assign_ant = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[5]/div/div')
        data_atual.append(assign_atual)
        data_anterior.append(assign_ant)
        print("   [OK] Assignment")

        # 4. 3PL HANDOVER
        page.click('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div/div/div/div[4]')
        page.wait_for_timeout(10000)
        threepl_atual = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[4]/div/div')
        threepl_ant = page.inner_text('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/table/thead/tr[2]/th[5]/div/div')
        data_atual.append(threepl_atual)
        data_anterior.append(threepl_ant)
        print("   [OK] 3PL")

    except Exception as e:
        print(f"Erro ao coletar dados: {e}")
        raise
        
    return data_anterior, data_atual

def update_google_sheets(data_anterior, data_atual):
    print("\n📝 Conectando ao Google Sheets...")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1_ySESJhetl_zVvRB1azj-7FjaldHFOQK_DKDQYhNiyc/edit').worksheet("Reporte HxH")

    current_time = datetime.datetime.now(timezone)
    hour = current_time.hour
    
    # Mapeamento estrito
    if 7 <= hour <= 23:
        row_anterior = hour - 5
    elif hour == 0:
        row_anterior = 19
    elif 1 <= hour <= 6:
        row_anterior = hour + 19
    else:
        print(f"Hora fora do intervalo: {hour}:{current_time.minute}")
        return

    # Linha atual na planilha
    row_atual = row_anterior + 1
    if row_atual > 25:
        row_atual = 2

    # Trava e Gravação da Hora Anterior
    if hour != 6:
        if 2 <= row_anterior <= 25:
            cell_range_ant = f'B{row_anterior}:E{row_anterior}'
            sheet.update(values=[data_anterior], range_name=cell_range_ant)
            print(f"✅ HORA ANTERIOR inserida na linha {row_anterior} ({cell_range_ant}).")
    else:
        print("⚠️ Virada de diária (06:xx). Ignorando HORA ANTERIOR (linha 25) para não sujar a planilha limpa.")

    # Gravação da Hora Atual
    if 2 <= row_atual <= 25:
        cell_range_atu = f'B{row_atual}:E{row_atual}'
        sheet.update(values=[data_atual], range_name=cell_range_atu)
        print(f"✅ HORA ATUAL inserida na linha {row_atual} ({cell_range_atu}).")

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(accept_downloads=True, viewport={'width': 1920, 'height': 1080})
        page = context.new_page()

        try:
            login(page)
            data_anterior, data_atual = get_dual_data(page)
            update_google_sheets(data_anterior, data_atual)
            print("\n🚀 Processo concluído.")
        except Exception as e:
            print(f"❌ Erro: {e}")
        finally:
            browser.close()

if __name__ == "__main__":
    main()
