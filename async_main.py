import json

import aiohttp
import asyncio
import aiofiles
import xlsxwriter
from typing import Dict
from typing import List

from bs4 import BeautifulSoup

categories = {
    'sportivnye-tovary-dlja-muzhchin.html': ['Мужсккие', 74],
    'sportivnye-tovary-dlja-zhenshhin.html': ['Женские', 68],
    'sportivnye-tovary-dlja-detej.html': ['Дети', 25]
}
url = 'https://ru.puma.com/'
headers = {
    'X-Requested-With': 'XMLHttpRequest',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15'}


async def main():
    params_list = []
    async with aiohttp.ClientSession(trust_env=True, timeout=aiohttp.ClientTimeout()) as session:
        urls = []
        for sex_key, sex in categories.items():
            sex_display, page_size = sex
            print(f'Loading sex: {sex_display}')
            urls += [url + sex_key + f'?p={page}' for page in range(1, page_size + 1)]
        params = await asyncio.gather(*[download_product_list(link, session) for link in urls])
        params_list += params
        data = await asyncio.gather(*[parse_product_list(**params) for params in params_list])

    async with aiofiles.open('products.json', 'w') as f:
        await f.write(json.dumps(data, ensure_ascii=False))
    loop = asyncio.get_running_loop()
    await loop.run_in_executor(None, save_as_sheet, data[0])


async def download_product_list(link: str, session):
    print(f'Loading: {link}')
    async with session.get(link, headers=headers) as resp:
        json_data = await resp.json()
    return {'html': json_data['content'], 'session': session}


async def parse_product_list(html: str, session) -> List[Dict[str, str]]:
    data = []
    bs = BeautifulSoup(html, 'html.parser')
    products = bs.find_all('div', attrs={'class': 'product-item'})

    for product in products:
        name = product.find(attrs={'class': 'product-item__name'}).get_text()
        link = product.find('a')['href']

        item_data = await parse_product_item(link, session)
        category = name.split(' ')[0]
        data.append({
            'name': name,
            'category': category,
            **item_data
        })

    return data


async def parse_product_item(link: str, session: aiohttp.ClientSession) -> Dict[str, str]:
    print(link)
    async with session.get(link) as resp:
        content = await resp.content.read()
    bs4 = BeautifulSoup(content, 'html.parser')
    sex = bs4.find_all('li', attrs={'class': 'breadcrumbs__item'})[1].get_text()
    vendor = bs4.find('span', attrs={'class': 'product-article__value'}).get_text()
    description = bs4.find('div', attrs={'class': 'product-attribute-description'}).find('p').get_text()
    seo_title = bs4.find('meta', attrs={'name': 'title'})['content']
    seo_description = bs4.find('meta', attrs={'name': 'description'})['content']
    seo_keywords = bs4.find('meta', attrs={'name': 'keywords'})['content']

    colors = []
    for color in bs4.find_all('a', attrs={'class': 'color-item__link'}):
        colors.append(color['title'])

    return {
        'sex': sex.strip(),
        'vendor': vendor,
        'description': description,
        'colors': ','.join(colors),
        'seo_title': seo_title,
        'seo_description': seo_description,
        'seo_keywords': seo_keywords
    }


def save_as_sheet(data: List[Dict[str, str]]) -> None:
    sheet_headers = ['Название', 'Категория', 'Пол', 'Артикул', 'Описание', 'Цвета',
                     'SEO заголовок', 'SEO описание', 'SEO ключевые слова ']
    workbook = xlsxwriter.Workbook('puma_products.xlsx', {'in_memory': True})
    worksheet = workbook.add_worksheet()

    for i, header in enumerate(sheet_headers):
        worksheet.write(0, i, header)

    for i, row in enumerate(data, 1):
        for j, value in enumerate(row.values(), 0):
            worksheet.write(i, j, value)

    workbook.close()


if __name__ == '__main__':
    asyncio.run(main())
