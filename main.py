import json
import traceback

import requests
import xlsxwriter
from typing import Dict
from typing import List

from bs4 import BeautifulSoup

sex_categories = {
    'sportivnye-tovary-dlja-muzhchin.html': 74,
    # 'sportivnye-tovary-dlja-zhenshhin.html': 68,
    # 'sportivnye-tovary-dlja-detej.html': 25
}
url = 'https://ru.puma.com/'
headers = {
    'X-Requested-With': 'XMLHttpRequest',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15'}
timeout = 30


def main():
    urls = []
    for sex_key, page_size in sex_categories.items():
        urls += [url + sex_key + f'?p={page}' for page in range(1, 2)]
    params = []
    for page_url in urls:
        try:
            params.append(download_product_list(page_url))
        except Exception:
            traceback.print_exc()

    success_products = []
    error_products = []
    for param in params:
        try:
            success_parses, error_parses = parse_product_list(param)
            success_products += success_parses
            error_products += error_parses
        except Exception:
            traceback.print_exc()
    with open('products.json', 'w') as f:
        f.write(json.dumps(success_products, ensure_ascii=False))
    save_as_sheet(success_products, error_products)


def download_product_list(link: str):
    print(f'Loading: {link}')
    res = requests.get(link, headers=headers, timeout=timeout)
    json_data = res.json()
    return json_data['content']


def parse_product_list(html: str) -> [List[Dict[str, str]], List[Dict[str, str]]]:
    data = []
    errors = []
    bs = BeautifulSoup(html, 'html.parser')
    products = bs.find_all('div', attrs={'class': 'product-item'})

    for i, product in enumerate(products):
        name = product.find(attrs={'class': 'product-item__name'}).get_text()
        category = name.split(' ')[0]
        link = product.find('a')['href']
        try:
            if i > 10:
                raise ValueError
            item_data = parse_product_item(link)
            data.append({
                'name': name,
                'category': category,
                **item_data
            })
        except Exception:
            traceback.print_exc()
            errors.append({
                'name': name,
                'category': category,
                'link': link
            })

    return data, errors


def parse_product_item(link: str) -> Dict[str, str]:
    print(f'Loading product by link: {link}')
    res = requests.get(link, headers=headers, timeout=timeout)
    content = res.content.decode()
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


def save_as_sheet(success_products: List[Dict[str, str]], error_products: List[Dict[str, str]]) -> None:
    sheet_headers = ['Название', 'Категория', 'Пол', 'Артикул', 'Описание', 'Цвета',
                     'SEO заголовок', 'SEO описание', 'SEO ключевые слова ']
    workbook = xlsxwriter.Workbook('puma_products.xlsx', {'in_memory': True})
    worksheet = workbook.add_worksheet()

    for i, header in enumerate(sheet_headers):
        worksheet.write(0, i, header)

    for i, row in enumerate(success_products, 1):
        for j, value in enumerate(row.values(), 0):
            worksheet.write(i, j, value)

    worksheet.write(i + 2, 0, 'Ошибки')
    for j, header in enumerate(['Название', 'Категория', 'Ссылка на товар']):
        worksheet.write(i + 3, j, header)

    for i, row in enumerate(error_products, i + 3):
        for j, value in enumerate(row.values(), 0):
            worksheet.write(i, j, value)

    workbook.close()


if __name__ == '__main__':
    main()
