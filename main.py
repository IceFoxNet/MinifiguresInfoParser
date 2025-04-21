import os

while True:
    try:
        from playwright.sync_api import sync_playwright, expect
        from configparser import ConfigParser
        import gspread, pathlib, requests
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        os.system(f'python -m pip install {package}')
    else:
        break

dir = pathlib.Path(__file__).parent.resolve()

config = ConfigParser()
config.read(os.path.join(dir, 'config.ini'))
sheet_url = config.get('parser', 'url')

def main(start: int, end: int, setup: dict):
    
    print('Подготавливаем всё для работы')
    if start < 3: start = 3
    cache = {}

    # ==> ПОДКЛЮЧЕНИЕ ГУГЛ-АККАУНТА
    creds = setup.get('GoogleCredentials')
    google_client = gspread.authorize(creds)
    spreadsheet = google_client.open_by_url(sheet_url)
    sheet = spreadsheet.worksheet('Минифигурки')
    
    # ==> ПОЛУЧЕНИЕ ДАННЫХ С ТАБЛИЦЫ
    articles = sheet.range(f'C{start}:C{end}')
    qty_res = []
    price_res = []
    name_res = []
    series_res = []

    # ==> ПОЛУЧЕНИЕ КУРСА ДОЛЛАРА
    rub = requests.get('https://www.cbr-xml-daily.ru/daily_json.js').json()['Valute']['USD']['Value']

    print('Начинаем работу с браузером')

    # ==> РАБОТА С БРАУЗЕРОМ
    for idx in range(len(articles)):
        try: 
            if not articles[idx].value: continue
        except: 
            break
        art = articles[idx].value
        if art in cache:
            qty_res.append([cache[art][0]])
            price_res.append([cache[art][1]])
            name_res.append([cache[art][2]])
            series_res.append([cache[art][3]])
            continue
        with sync_playwright() as p:
            driver = p.chromium.launch(proxy={
                'server': 'http://166.0.211.142:7576',
                'username': 'user258866',
                'password': 'pe9qf7'
            })
            page = driver.new_page()
            try:
                print(f'Обработка артикула {articles[idx].value}')
                page.goto(f'https://www.bricklink.com/v2/catalog/catalogitem.page?M={articles[idx].value}#T=P')
                page.wait_for_selector('table.pcipgMainTable')
                table = page.query_selector('#_idPGContents > table > tbody > tr:nth-child(3) > td:nth-child(4)')
                rows = table.query_selector_all('tr')
                qty_val = int(rows[1].query_selector_all('td')[-1].text_content())
                prc_val = round(float(rows[4].query_selector_all('td')[-1].text_content()[4:]) * rub)
                name_val = page.query_selector('#item-name-title').text_content()
                catalog_line = page.query_selector('#content > div > table > tbody > tr > td:nth-child(1)').query_selector_all('a')
                series_val = catalog_line[2].inner_text()
                if series_val == 'Super Heroes':
                    series_val = page.query_selector('#id_divBlock_Main > table:nth-child(1) > tbody > tr:nth-child(3) > td > div:nth-child(1) > table > tbody > tr > td > a').inner_text().split()[0]
                elif series_val == 'Town':
                    series_val = 'City'
            except Exception as e:
                print(f'Ошибка при работе с артикулом {articles[idx].value} (https://www.bricklink.com/v2/catalog/catalogitem.page?M={articles[idx].value}#T=P')
                qty_res.append([None])
                price_res.append([None])
                name_res.append([None])
                series_res.append([None])
            else:
                print(series_val, name_val, prc_val, qty_val)
                qty_res.append([qty_val])
                price_res.append([prc_val])
                name_res.append([name_val])
                series_res.append([series_val])
                cache[art] = (qty_val, prc_val, name_val, series_val)
            finally: continue
    print(f'Загружаем информацию на таблицу')
    sheet.update(qty_res, f'V{start}:V{len(qty_res)+start}')
    sheet.update(price_res, f'T{start}:T{len(price_res)+start}')
    sheet.update(name_res, f'B{start}:B{len(name_res)+start}')
    sheet.update(series_res, f'A{start}:A{len(series_res)+start}')
    print(f'Программа завершила выполнение')

if __name__ == '__main__':
    from Setup.setup import setup
    main(6, 7, setup)