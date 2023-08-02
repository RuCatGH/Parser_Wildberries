import asyncio

import jmespath
import aiohttp
import pandas as pd
from fake_useragent import UserAgent
from urllib.parse import quote, unquote

ua = UserAgent()

def read_xlsx():
    df = pd.read_excel("data 5.xlsx")
    request_name = df.iloc[:, 0].to_list()
    return request_name, df

def get_basket_url(e):
    if 0 <= e <= 143:
        return "//basket-01.wb.ru/"
    elif 144 <= e <= 287:
        return "//basket-02.wb.ru/"
    elif 288 <= e <= 431:
        return "//basket-03.wb.ru/"
    elif 432 <= e <= 719:
        return "//basket-04.wb.ru/"
    elif 720 <= e <= 1007:
        return "//basket-05.wb.ru/"
    elif 1008 <= e <= 1061:
        return "//basket-06.wb.ru/"
    elif 1062 <= e <= 1115:
        return "//basket-07.wb.ru/"
    elif 1116 <= e <= 1169:
        return "//basket-08.wb.ru/"
    elif 1170 <= e <= 1313:
        return "//basket-09.wb.ru/"
    elif 1314 <= e <= 1601:
        return "//basket-10.wb.ru/"
    elif 1602 <= e <= 1655:
        return "//basket-11.wb.ru/"
    elif 1656 <= e <= 1919:
        return "//basket-12.wb.ru/"
    else:
        return "//basket-13.wb.ru/"

async def get_data_from_id(country, json):
    try:
        for json_basket_item in json['options']:
            if json_basket_item['value'] == country:
                return 1
    except:
        pass
    return 0

async def get_data(query):
    try:
        if isinstance(query,str):
            query = quote(query)
            link = f'https://www.wildberries.ru/catalog/0/search.aspx?search={query}'
            async with aiohttp.ClientSession() as session:
                async with session.get(
    f'https://search.wb.ru/exactmatch/ru/common/v4/search?TestGroup=test_1&TestID=155&appType=1&curr=rub&dest=-1252424&query={query}&regions=80,38,83,4,64,33,68,70,30,40,86,75,69,1,66,110,48,22,31,71,114&resultset=catalog&sort=popular&spp=31&suppressSpellcheck=false&uclusters=9',headers={"User-Agent": ua.random}) as response:
                    json_search=await response.json(content_type=None)
                    ids = jmespath.search('data.products[:30].id', json_search)
                    try:
                        brandId = jmespath.search('data.products[0].brandId', json_search)
                        kindId = jmespath.search('data.products[0].kindId', json_search)
                        subjectId = jmespath.search('data.products[0].subjectId', json_search)
  
                        params = {
                                'subject': subjectId,
                                'kind': kindId,
                                'brand': brandId,
                            }
                        async with session.get(f'https://www.wildberries.ru/webapi/product/{ids[0]}/data', params=params, headers={"User-Agent": ua.random}) as r:
                            json = await r.json(content_type=None)
                            category = '/'.join(jmespath.search('value.data.sitePath[*].name', json))
                    except:
                        category = ''

                    first_product_price = str(jmespath.search('data.products[0].salePriceU', json_search))[:-2]


                    number_of_china = 0
                    for id in ids:
                        part = str(id)[:-3]
                        vol = int(part[:-2])
                        async with session.get(f'https:{get_basket_url(vol)}vol{vol}/part{part}/{id}/info/ru/card.json', headers={"User-Agent": ua.random}) as r:
                            json_basket = await r.json(content_type=None)

                            number_of_china += await get_data_from_id('Китай', json_basket)
                    data_china = f'{number_of_china}/30'
        else:
            data_china = ''
            category = ''
            link = f'https://www.wildberries.ru/catalog/0/search.aspx?search={query}'
            first_product_price=''

    except Exception as ex:
        data_china = ''
        category = ''
        link = f'https://www.wildberries.ru/catalog/0/search.aspx?search={query}'
        first_product_price = ''
    return data_china, link, first_product_price, category


async def main():
    queries, df = read_xlsx()
    data = []
    tasks = []  # Список для хранения задач get_data()
    for query in queries:
        task = asyncio.create_task(get_data(query))  # Создание задачи get_data()
        tasks.append(task)
        if len(tasks) == 1000:
            results = await asyncio.gather(*tasks)
            for result in results:
                data.append(result)
            tasks = []
    else:
        results = await asyncio.gather(*tasks)
        for result in results:
            data.append(result)

        
    # Обработка данных и сохранение результата в файл
    count_china = [item[0] for item in data]
    link = [item[1] for item in data]
    first_price = [item[2] for item in data]
    category = [item[3] for item in data]

    df['Ссылка на запрос'] = link
    df['Количество китайских товаров в выдаче'] = count_china
    df['Цена топ-1 товара по запросу'] = first_price
    df['Категория'] = category
    df.to_excel("result.xlsx", index=False)

    
if __name__ == "__main__":
    asyncio.run(main())
        
        
    
    
