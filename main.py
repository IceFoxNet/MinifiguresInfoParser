import os

while True:
    try:
        from playwright.async_api import async_playwright, Browser
        import gspread, pathlib, requests, asyncio
        from async_lru import alru_cache
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        os.system(f'python -m pip install {package}')
    else:
        break

dir = pathlib.Path(__file__).parent.resolve()
rub = requests.get('https://www.cbr-xml-daily.ru/daily_json.js').json()['Valute']['USD']['Value']

async def main(start: int, end: int, setup: dict):
    
    print('Подготавливаем всё для работы')
    if start < 3: start = 3

    # ==> ПОДКЛЮЧЕНИЕ ГУГЛ-АККАУНТА
    spreadsheet: gspread.spreadsheet.Spreadsheet = setup.get('MinifiguresSheet')
    sheet = spreadsheet.worksheet('Минифигурки')
    
    # ==> ПОЛУЧЕНИЕ ДАННЫХ С ТАБЛИЦЫ
    articles = sheet.range(f'C{start}:C{end}')
    qty_res = []
    price_res = []
    name_res = []
    series_res = []

    print('Начинаем работу с браузером')

    # ==> РАБОТА С БРАУЗЕРОМ

    semaphore = asyncio.Semaphore(4)
    async with async_playwright() as p: 
        browser = await p.chromium.launch(proxy={
            'server': 'http://166.0.211.142:7576',
            'username': 'user258866',
            'password': 'pe9qf7'
        })
        tasks = [parse_item(semaphore, browser, articles[idx].value) for idx in range(len(articles))]
        results = await asyncio.gather(*tasks)
    for res in results:
        qty_res.append([res[0]])
        price_res.append([res[1]])
        name_res.append([res[2]])
        series_res.append([res[3]])
    print(f'Загружаем информацию на таблицу')
    sheet.update(qty_res, f'V{start}:V{len(qty_res)+start}')
    sheet.update(price_res, f'T{start}:T{len(price_res)+start}')
    sheet.update(name_res, f'B{start}:B{len(name_res)+start}')
    sheet.update(series_res, f'A{start}:A{len(series_res)+start}')
    print(f'Программа завершила выполнение')

@alru_cache(None)
async def parse_item(semaphore: asyncio.Semaphore, driver: Browser, art: str):
    global rub
    async with semaphore:
        print(f'Обработка артикула {art}')
        page = await driver.new_page()
        try:
            await page.goto(f'https://www.bricklink.com/v2/catalog/catalogitem.page?M={art}#T=P')
            await page.wait_for_selector('table.pcipgMainTable')
            table = await page.query_selector('#_idPGContents > table > tbody > tr:nth-child(3) > td:nth-child(4)')
            rows = await table.query_selector_all('tr')
            # Get Qty Value
            rows_1_all_td = await rows[1].query_selector_all('td')
            qty_val = int(await rows_1_all_td[-1].text_content())
            # Get Price Value
            rows_4_all_td = await rows[4].query_selector_all('td')
            rows_4_text_content = await rows_4_all_td[-1].text_content()
            prc_val = round(float(rows_4_text_content[4:]) * rub)
            # Get Name Value
            item_name_title_div = await page.query_selector('#item-name-title')
            name_val = await item_name_title_div.text_content()
            # Get Series Value
            catalog_line = await (await page.query_selector('#content > div > table > tbody > tr > td:nth-child(1)')).query_selector_all('a')
            series_val = await catalog_line[2].inner_text()
            if series_val == 'Super Heroes':
                co_autor = await page.query_selector('#id_divBlock_Main > table:nth-child(1) > tbody > tr:nth-child(3) > td > div:nth-child(1) > table > tbody > tr > td > a')
                co_autor_text = await co_autor.inner_text()
                series_val = co_autor_text.split()[0]
            elif series_val == 'Town':
                series_val = 'City'
            elif series_val == 'Space':
                try:
                    if await catalog_line[3].inner_text() == 'Galaxy Squad':
                        series_val = 'Galaxy Squad'
                except:
                    series_val = 'Space'
        except Exception as e:
            await page.close()
            return((None, None, None, None))
        else:
            await page.close()
            return((qty_val, prc_val, name_val, series_val))

if __name__ == '__main__':
    from Setup.setup import setup
    asyncio.run(main(3, 10, setup))