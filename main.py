import asyncio
import json

import aiohttp
from openpyxl import Workbook
from fake_useragent import UserAgent

ua = UserAgent()

headers = {
    'user-agent':f'{ua.random}'
}

wb = Workbook()
ws = wb.active
ws.append(['Название каталога','Ссылка на каталог','Количество','Из 10','Ссылки на товар'])


async def get_data():
    async with aiohttp.ClientSession() as session:
        try:
            async with session.get('https://www.wildberries.ru/webapi/menu/main-menu-ru-ru.json', headers=headers) as response:
                catalogs_data = await response.json()
                for catalog_data in catalogs_data:
                    for child in catalog_data['childs']: # Забираем query, chard и вызываем get_products для получения данных
                        child_name = child['name']
                        try:
                            for sub_catalog in child['childs']: # Если есть ещё один подкаталог
                                sub_catalog_name = '/'+sub_catalog['name']
                                try:
                                    for sub_sub_catalog in sub_catalog['childs']:
                                        name_sub_sub_catalog = sub_sub_catalog['name']
                                        query = sub_sub_catalog['query']
                                        shard = sub_sub_catalog['shard']
                                        pages_data = await get_products(session,query, shard) # Get products data
                                        if pages_data:
                                            for page_data in pages_data:
                                                catalog_name = catalog_data['name']+'/'+child_name+sub_catalog_name+'/'+name_sub_sub_catalog+'/'+page_data[3]
                                                catalog_url = 'https://www.wildberries.ru'+child['url']+f'?page=1&xsubject={page_data[2]}'
                                                ws.append([catalog_name,catalog_url,page_data[0],page_data[4]]+page_data[1])
                                            wb.save('Data.xlsx')     
                                except:
                                    try:
                                        query = sub_catalog['query']
                                        shard = sub_catalog['shard']
                                        pages_data = await get_products(session,query, shard) # Get products data
                                        if pages_data:
                                            for page_data in pages_data:
                                                catalog_name = catalog_data['name']+'/'+child_name+sub_catalog_name+'/'+page_data[3]
                                                catalog_url = 'https://www.wildberries.ru'+child['url']+f'?page=1&xsubject={page_data[2]}'
                                                ws.append([catalog_name,catalog_url,page_data[0],page_data[4]]+page_data[1])
                                            wb.save('Data.xlsx')
                                            # print('Done')
                                    except:
                                        continue

                        except:
                            try:
                                sub_catalog_name=''
                                query = child['query']
                                shard = child['shard']
                                pages_data = await get_products(session, query,shard)
                                if pages_data:
                                    for page_data in pages_data:
                                        catalog_url = 'https://www.wildberries.ru'+child['url']+f'?page=1&xsubject={page_data[2]}'
                                        catalog_name = catalog_data['name']+'/'+child_name+sub_catalog_name+'/'+page_data[3]
                                        ws.append([catalog_name,catalog_url,page_data[0],page_data[4]]+page_data[1])
                                    wb.save('Data.xlsx') 
                                    # print('Done')
                            except:
                                continue
        except Exception as ex:
            print('Ошибка',ex)


async def get_products(session, query, shard, retry=5):
    url = f'https://catalog.wb.ru/catalog/{shard}/v4/filters?appType=1&couponsGeo=12,3,18,15,21&curr=rub&dest=-1029256,-102269,-2162196,-1275551&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&reg=0&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,22,66,31,40,1,48,71&spp=0&{query}'
    try:
        async with session.get(url, headers=headers) as response:
            json_data =  json.loads(await response.read())
            for filter in json_data['data']['filters']:
                if filter['name'] == 'Категория':
                    data = []
                    for item in filter['items']:
                        tasks = []
                        result_sum = []
                        necessary_urls = []
                        first_10_products = []
                        filter_id = item['id'] # ID фильтра
                        filter_name = item['name'] # Название фильтра 
                        link = f'https://catalog.wb.ru/catalog/{shard}/catalog?appType=1&couponsGeo=12,3,18,15,21&curr=rub&dest=-1029256,-102269,-2162196,-1275551&emp=0&lang=ru&locale=ru&page=1&pricemarginCoeff=1.0&reg=0&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,22,66,31,40,1,48,71&sort=popular&spp=0&{query}&xsubject={filter_id}'
                        try:
                            async with session.get(link,headers=headers) as r: # Проходимся по всем карточкам и возращаем нужные данные
                                json_catalog_data = json.loads(await r.read())
                                for product in json_catalog_data['data']['products']:
                                    id = product['id']
                                    task = asyncio.create_task(get_product_data(session,id))
                                    tasks.append(task)
                                products_data = await asyncio.gather(*tasks) 
                                for i,product_data in enumerate(products_data):
                                    if i < 10:
                                        first_10_products.append(product_data[0])
                                    
                                    result_sum.append(product_data[0]) # 0 или 1
                                    if product_data[0] == 1:
                                        necessary_urls.append(product_data[1]) # Добавление ссылки на товар
                                data.append([f'{sum(result_sum)}/{len(result_sum)}', necessary_urls,filter_id,filter_name,f'{sum(first_10_products)}/{len(first_10_products)}'])
                        except Exception as ex:
                            print(ex)
                    return data
    except Exception as ex:
        if retry:
            print(f'[INFO] retry={retry} => {url}')
            print(ex)
            await get_products(session, query, shard, retry=(retry-1))
        else:
            raise
    return False

async def get_product_data(session,id): # Здесь мы возращаем 1 вместе со ссылкой, если страна Турция, если нет то 0
    async with session.get(f'https://wbx-content-v2.wbstatic.net/ru/{id}.json',headers=headers) as response:
        data = await response.json()
        link = f'https://www.wildberries.ru/catalog/{id}/detail.aspx?targetUrl=GP'
        try:
            for country in data['options']:
                if country['name'] == 'Страна производства':
                    if country['value'] == 'Турция':
                        return [1,link]  # Ссылка на товар
                    else:
                        return [0]
        except:
            pass
    return [0]


def main():
    asyncio.get_event_loop().run_until_complete(get_data())
   
if __name__ == '__main__':
    main()