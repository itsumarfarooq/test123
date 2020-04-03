import requests
from parsel import Selector
import datetime
from xlsx_file_handling import read_xlsx_list_of_dict, write_list_of_dict_xlsx

date_part=datetime.datetime.now().strftime('%d %b %Y %H_%M_%S')

input_file_name='Copy of Top Knobs Price List 2020 Final.xlsx'
output_file_name=f'output_{date_part}.xlsx'
notfound_file_name=f'notfound_{date_part}.xlsx'
noimage_file_name=f'noimage_{date_part}.xlsx'
output_data=[]
notfound_data=[]
noimage_data=[]
def search_sku(product_sku):
    search_url=f'https://www.topknobs.com/catalogsearch/result/index/?limit=36&q={product_sku}'
    r=requests.get(search_url)
    sel=Selector(r.text)
    product_url=sel.xpath(f'//span[starts-with(@id,"sku-") and text()="{product_sku}"]/parent::a/@href').extract_first()
    return product_url

def extract_data(product_url):
    r=requests.get(product_url)
    sel=Selector(r.text)
    product_info={}
    product_info['product_url']=product_url
    product_info['product_name']=sel.xpath('//h1[@class="product-name"]/text()').extract_first()
    product_info['product_image_url']=sel.xpath('//img[@class="gallery-image"]/@src').extract_first()
    for li_elem in sel.xpath('//ul[@class="attributes"]/li'):
        key_value=li_elem.xpath('./span[@class="name"]/text()').extract_first()
        product_info['Extracted'+key_value]=li_elem.xpath('./span[@class="value"]/text()').extract_first().strip()
    return product_info


input_data=read_xlsx_list_of_dict(input_file_name)
for n,product_info in enumerate(input_data):
    product_url=search_sku(product_info['Part Number'])
    if product_url:
        extracted_product_info=extract_data(product_url)
        output_row = {**product_info, **extracted_product_info}
        notfound_data.append(output_row)
        if not extracted_product_info['product_image_url']:
            print(f'{n}. No Image Found For :', product_info['Part Number'])
            noimage_data.append(output_row)
        print(f'{n}. Scraped :', product_info['Part Number'])
    else:
        output_row=product_info
        print(f'{n}. Not Found :', product_info['Part Number'])
    output_data.append(output_row)

write_list_of_dict_xlsx(output_file_name, output_data)
write_list_of_dict_xlsx(notfound_file_name, notfound_data)
write_list_of_dict_xlsx(noimage_file_name, noimage_data)
