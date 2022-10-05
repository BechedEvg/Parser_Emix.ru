import requests
import urllib3
import ssl
import xlsxwriter
import pandas


class CustomHttpAdapter(requests.adapters.HTTPAdapter):
    # "Transport adapter" that allows us to use custom ssl_context.

    def __init__(self, ssl_context=None, **kwargs):
        self.ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False):
        self.poolmanager = urllib3.poolmanager.PoolManager(
            num_pools=connections, maxsize=maxsize,
            block=block, ssl_context=self.ssl_context)


def get_legacy_session():
    ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
    ctx.options |= 0x4  # OP_LEGACY_SERVER_CONNECT
    session = requests.session()
    session.mount('https://', CustomHttpAdapter(ctx))
    return session


def checking_the_volume_of_liters(description):
    try:
        if (description.split()[-1][-1]).lower() == "л":
            if (description.split()[-1][:-1]).isdigit():
                return description.split()[-1][:-1]
        elif (description.split()[-2]).isdigit():
            return description.split()[-2]
        else:
            return ''
    except:
        return ''


# Reading a list of elements from a source file.
def read_exel(doc):
    workbook = pandas.read_excel(doc)
    list_product = []
    for elements in workbook.values:
        list_product.append(list(elements)[:5])
    return list_product


def write_exel(write_list):
    workbook = xlsxwriter.Workbook("Result.xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column('A:I', 25)
    worksheet.write('A1', 'ID', bold)
    worksheet.write('B1', 'Марка', bold)
    worksheet.write('C1', 'Модель', bold)
    worksheet.write('D1', 'Наименование запчасти', bold)
    worksheet.write('E1', 'Артикул', bold)
    worksheet.write('F1', 'Рейтинг', bold)
    worksheet.write('G1', 'Цена', bold)
    worksheet.write('H1', 'Описание  товара', bold)
    worksheet.write('I1', 'Объем, л. (Для жидкостей)', bold)
    counter = 2
    for elements in write_list:
        worksheet.write(f'A{counter}', f'{elements[0]}')
        worksheet.write(f'B{counter}', f'{elements[1]}')
        worksheet.write(f'C{counter}', f'{elements[2]}')
        worksheet.write(f'D{counter}', f'{elements[3]}')
        worksheet.write(f'E{counter}', f'{elements[4]}')
        worksheet.write(f'F{counter}', f'{elements[5]}')
        worksheet.write(f'G{counter}', f'{elements[6]}')
        worksheet.write(f'H{counter}', f'{elements[7]}')
        worksheet.write(f'I{counter}', f'{elements[8]}')
        counter += 1
    workbook.close()


def get_html(url):
    html = get_legacy_session().get(url)
    return html


# Get a list of products from a dictionary.
def get_emex_list_products(html_dict, elements_list):
    item_list = []
    list_dict_elements = html_dict["originals"][0]["offers"]
    description = html_dict["originals"][0]["name"]
    volume = checking_the_volume_of_liters(description)

    # We check the presence of values in the dictionary of goods.
    # If it is empty, then the item is out of stock.

    if len(list_dict_elements) != 0:
        for dict_values in list_dict_elements:
            rating = dict_values["rating2"]["rating"]
            price = dict_values["displayPrice"]["value"]
            item_list.append(elements_list + [rating, price, description, volume])
    else:
        price = "Товар закончился"
        rating = "Товар закончился"
        item_list.append(elements_list + [rating, price, description, volume])
    return item_list


# Get a ready-made list of goods for recording.
def get_write_list_products(list_product_elements):
    write_list = []
    for elements_list in list_product_elements:
        vendor_cod = elements_list[4]
        html = get_html(f"https://emex.ru/api/search/search?detailNum={vendor_cod}&locationId=29241&showAll=true&longitude=37.5739&latitude=55.8095")
        html_dict = html.json().get('searchResult')

        # Check the product dictionary for the list key for all the products you are looking for.
        # If not, then we get a list of manufacturers
        # and add them to the query to get a list of all products by manufacturer.

        if "originals" not in html_dict:
            manufacturer_list = html_dict["makes"]["list"]
            for make in manufacturer_list:
                manufacturer = make["make"]
                html = get_html(f"https://emex.ru/api/search/search?detailNum={vendor_cod}&make={manufacturer}&locationId=29241&showAll=true&longitude=37.5739&latitude=55.8095")
                html_dict = html.json().get('searchResult')
                write_list += get_emex_list_products(html_dict, elements_list)
        write_list += get_emex_list_products(html_dict, elements_list)
    return write_list


def main():
    list_read_product = read_exel("Vendor.xlsx")
    list_write = get_write_list_products(list_read_product)
    write_exel(list_write)


if __name__ == '__main__':
    main()
