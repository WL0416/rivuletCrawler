import scrapy
from xlwt import *
import os
from xlutils.copy import copy
from xlrd import open_workbook
import urllib

class NewlineSpider(scrapy.Spider):

    name = "newline"

    def start_requests(self):

        urls = [
            'https://www.newlineofficefurniture.com.au/',
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):

        # page = response.url.split('/')[-2]
        
        # filename = 'pages/%s.html' % page
        
        # with open(filename, 'wb') as f:
            
        #     f.write(response.body)
        
        # self.log('Saved file %s' % filename)
        
        crawled_book = Workbook()
        product_sheet = crawled_book.add_sheet('Products', cell_overwrite_ok=True)
        category_sheet = crawled_book.add_sheet('Category', cell_overwrite_ok=True)

        categories = []
        links = []

        for category in response.css('div.dropdown-menu'):
            categories.append(category.css('a::text').extract())
            links.append(category.css('a::attr(href)').extract())

        # check if the element is empty in category list, if it is empty, remove it.
        # remove category list and link list at the same time
        for index in range(len(categories)):
            if len(categories[index]) is 0:
                categories.pop(index)
                links.pop(index)
        
        category_id = 0
        category_data = []
        previous_len = None
        parents = []
        ready_crawl_links = []
        # operate the link string, check the 
        for link_group_index in range(len(links)):
            for link_index in range(len(links[link_group_index])):
                current_link = links[link_group_index][link_index]
                link_splited = current_link.split('/')
                # if link char list's length is 4, it indicates the top category
                # category is tuple (parent id, category name, category id)
                current_len = len(link_splited)
                if current_len is 4:
                    parents = []
                    category = ('', categories[link_group_index][link_index].strip(), str(category_id))
                    previous_len = current_len
                else:
                    if current_len > previous_len:
                        parents.append(str(category_id))
                    elif current_len < previous_len:
                        parents.pop()
                        ready_crawl_links.pop()
                    category_id += 1
                    parent_str = ''
                    for index in range(len(parents)):
                        if index > 0:
                            parent_str += ','
                        parent_str += parents[index]
                    
                    category = (parent_str, 
                                categories[link_group_index][link_index].strip(), 
                                str(category_id))

                    ready_crawl_links.append(current_link)

                    previous_len = current_len

                category_data.append(category)

        for row_index in range(len(category_data)):
            for column_index in range(len(category_data[row_index])):
                value = category_data[row_index][column_index]
                category_sheet.write(row_index, column_index, value)

        crawled_book.save('newline_crawled.xls')

        # if the category link list is not empty.
        while len(ready_crawl_links) > 0:
            next_page = ready_crawl_links.pop()
            next_page = response.urljoin(next_page)
            yield scrapy.Request(next_page, callback=self.parse_category)

    def parse_category(self, response):
        
        content = response.css('div#content h1::text').extract()

        if len(content) > 0:
            
            for product in response.css('div.product-thumb'):

                products = product.css('h4 a::attr(href)').extract()

                while len(products) > 0:
                    next_page = products.pop()
                    next_page = response.urljoin(next_page)
                    yield scrapy.Request(next_page, callback=self.parse_products)     

    def parse_products(self, response):  

        rexcel = open_workbook("newline_crawled.xls")
        row = rexcel.sheets()[0].nrows
        excel = copy(rexcel)
        table = excel.get_sheet(0)
        item_category = response.css('ul.breadcrumb li span::text').extract()[-2]
        
        for product_info in response.css('div.col-sm-9'):
                
            item_sku = product_info.css('ul.list-unstyled.description li span#uo_sku_model::text').extract_first()
            item_mode = product_info.css('ul.list-unstyled.description li span#uo_sku_model::text').extract_first()
            item_otherid = ''
            item_name = product_info.css('h1::text').extract_first().strip()
            # print(item_name)
            item_description = product_info.css('span#tab-description p::text').extract()
            item_description = '\n'.join(item_description)
            # print(item_description)
            item_photo_name = product_info.css('img::attr(src)').extract_first().split('/')[-1].strip()
            # print(item_photo_name)
            item_photo_link = product_info.css('img::attr(src)').extract()

            if not os.path.exists('images/newline/' + item_name):
                os.makedirs('images/newline/' + item_name)

            image_count = 0
            for link in item_photo_link:

                urllib.urlretrieve(link, 'images/newline/' + item_name + '/' + item_name + '_' + str(image_count) + '.png')

                image_count += 1

            item_photo_link = '\n'.join(item_photo_link)
            # print(item_photo_link)
            item_price = product_info.css('span#uo_price::text').extract_first()
            # print(item_price)
            
            item = (item_sku, 
                    item_mode, 
                    item_otherid, 
                    item_name, 
                    item_description, 
                    item_photo_name,
                    item_photo_link,
                    item_category,
                    item_price)

            for index in range(len(item)):
                table.write(row, index, item[index])

        excel.save("newline_crawled.xls")