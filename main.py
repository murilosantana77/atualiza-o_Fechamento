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

def extract_value_by_header(page, col_text, row_text):
    """
    Função inteligente que varre a tabela HTML procurando o cruzamento exato 
    entre a coluna (ex: 14:00) e a linha (ex: Total).
    """
    page.wait_for_timeout(3000) 
    try:
        # Rola a tabela para a direita para forçar o carregamento das colunas ocultas
        page.evaluate("""() => {
            const tableContainer = document.querySelector('.table-container') || document.querySelector('.ant-table-body') || document.querySelector('table').parentElement;
            if (tableContainer) tableContainer.scrollLeft = tableContainer.scrollWidth;
        }""")
        page.wait_for_timeout(1000)
    except:
        pass

    value = page.evaluate('''([colText, rowText]) => {
        let trs = Array.from(document.querySelectorAll('tr'));
        const cleanText = (text) => text.replace(/\\s+/g, ' ').trim();
        const cleanCol = cleanText(colText);
        const cleanRow = cleanText(rowText);

        // Acha o cabeçalho
        let headerRow = trs.find(tr => cleanText(tr.innerText).includes(cleanCol) && tr.querySelectorAll('th, td').length > 3);
        if (!headerRow) return "0";

        let visualColIndex = -1;
        let currentVisIndex = 0;
        let headers = Array.from(headerRow.querySelectorAll('th, td'));
        
        for (let cell of headers) {
            let colSpan = parseInt(cell.getAttribute('colspan')) || 1;
            if (cleanText(cell.innerText).includes(cleanCol)) {
                visualColIndex = currentVisIndex;
                break;
            }
            currentVisIndex += colSpan;
        }

        if (visualColIndex === -1) return "0";

        // Acha a linha de dados
        let targetRow = trs.find(tr => cleanText(tr.innerText).includes(cleanRow));
        if (!targetRow) return "0";

        let cells = Array.from(targetRow.querySelectorAll('th, td'));
        let targetVisIndex = 0;
        
        for (let cell of cells) {
            let colSpan = parseInt(cell.getAttribute('colspan')) || 1;
            if (visualColIndex >= targetVisIndex && visualColIndex < targetVisIndex + colSpan) {
                let val = cell.innerText.replace(/[^0-9]/g, ''); 
                return val ? val : "0";
            }
            targetVisIndex += colSpan;
        }
        return "0";
    }''', [col_text, row_text])
    
    return value if value else "0"

def get_data_for_target(page, hour_int):
    """Navega pelas abas e coleta os 4 processos para uma hora específica."""
    hour_str = f"{hour_int:02d}:00"
    next_h = (hour_int + 1) % 24
    inbound_col = f"{hour_int:02d}:00-{next_h:02d}:00"
    
    print(f"\n--- 📊 Coletando dados da hora alvo: {hour_str} ---")
    data = []
    
    try:
        # 1. Inbound (Historical Data)
        print("-> Lendo Inbound...")
        page.goto("https://spx.shopee.com.br/#/dashboard/facility-soc/historical-data")
        page.wait_for_timeout(8000)
        val1 = extract_value_by_header(page, inbound_col, "SOC Received")
        if val1 == "0": # Tenta o formato simples se o formato com traço falhar
            val1 = extract_value_by_header(page, hour_str, "SOC Received")
        data.append(val1)
        print(f"   [OK] Inbound: {val1}")

        # 2. Packing (Productivity)
        print("-> Lendo Packing...")
        page.goto("https://spx.shopee.com.br/#/dashboard/toProductivity?page_type=Outbound")
        page.wait_for_timeout(6000)
        val2 = extract_value_by_header(page, hour_str, "Total")
        data.append(val2)
        print(f"   [OK] Packing: {val2}")

        # 3. Assignment
        print("-> Lendo Assignment...")
        page.click('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div/div/div/div[3]')
        page.wait_for_timeout(4000)
        val3 = extract_value_by_header(page, hour_str, "Total")
        data.append(val3)
        print(f"   [OK] Assignment: {val3}")

        # 4. 3PL Handover
        print("-> Lendo 3PL Handover...")
        page.click('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div/div/div/div[4]')
        page.wait_for_timeout(4000)
        val4 = extract_value_by_header(page, hour_str, "Total")
        data.append(val4)
        print(f"   [OK] 3PL: {val4}")
        
    except Exception as e:
        print(f"❌ Erro ao coletar hora {hour_str}: {e}")
        
    return data

def update_google_sheets(data_anterior, data_atual, hour_atual):
    print("\n📝 Conectando ao Google Sheets...")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1_ySESJhetl_zVvRB1azj-7FjaldHFOQK_DKDQYhNiyc/edit').worksheet("Reporte HxH")

    # Mapeamento estrito baseado na hora em que o script roda (hour_atual)
    if 7 <= hour_atual <= 23:
        row_anterior = hour_atual - 5
    elif hour_atual == 0:
        row_anterior = 19
    elif 1 <= hour_atual <= 6:
        row_anterior = hour_atual + 19
    else:
        print(f"Hora fora do intervalo: {hour_atual}")
        return

    # Mapeamento da linha atual (se anterior for 25, atual volta pro começo na linha 2)
    row_atual_sheet = row_anterior + 1
    if row_atual_sheet > 25:
        row_atual_sheet = 2

    # 1. Atualiza Hora Anterior (TRAVA DE SEGURANÇA: ignora às 06:xx para não sujar a limpeza)
    if hour_atual != 6 and data_anterior:
        cell_range_ant = f'B{row_anterior}:E{row_anterior}'
        sheet.update(values=[data_anterior], range_name=cell_range_ant)
        print(f"✅ Dados da HORA ANTERIOR gravados na linha {row_anterior} ({cell_range_ant}).")
    elif hour_atual == 6:
        print("⚠️ Virada do dia (06:xx) detectada. Ignorando a gravação da linha 25 para preservar a planilha.")

    # 2. Atualiza Hora Atual
    if data_atual:
        cell_range_atu = f'B{row_atual_sheet}:E{row_atual_sheet}'
        sheet.update(values=[data_atual], range_name=cell_range_atu)
        print(f"✅ Dados da HORA ATUAL gravados na linha {row_atual_sheet} ({cell_range_atu}).")


def main():
    # Descobre o horário exato agora
    current_time = datetime.datetime.now(timezone)
    hour_atual = current_time.hour
    hour_anterior = (hour_atual - 1) % 24

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(accept_downloads=True, viewport={'width': 1920, 'height': 1080})
        page = context.new_page()

        try:
            login(page)
            
            # Coleta os dois lotes de dados de forma totalmente independente no SPX
            data_anterior = get_data_for_target(page, hour_anterior)
            data_atual = get_data_for_target(page, hour_atual)
            
            # Envia ambos para a planilha
            update_google_sheets(data_anterior, data_atual, hour_atual)
            
            print("\n🚀 Processo concluído com sucesso.")
        except Exception as e:
            print(f"\n❌ Erro crítico: {e}")
        finally:
            browser.close()

if __name__ == "__main__":
    main()
