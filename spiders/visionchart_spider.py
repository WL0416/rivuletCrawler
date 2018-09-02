import scrapy
from xlwt import *
import os
from xlutils.copy import copy
from xlrd import open_workbook
from scrapy.selector import Selector as selector

class NewlineSpider(scrapy.Spider):

    name = "visionchart"

    def start_requests(self):

        urls = [
            'http://www.visionchart.com.au/',
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):

        page = response.url.split('/')[-2]
        
        filename = 'pages/%s.html' % page
        
        with open(filename, 'wb') as f:
            
            f.write(response.body)
        
        self.log('Saved file %s' % filename)
        
        crawled_book = Workbook()
        product_sheet = crawled_book.add_sheet('Products', cell_overwrite_ok=True)
        category_sheet = crawled_book.add_sheet('Category', cell_overwrite_ok=True)

        categories = []
        links = []

        for category in response.css('div#u290672'):

            categories = category.css('div.MenuItemContainer.clearfix.colelem').xpath('string(.)').extract()

            links.append(category.css('a::attr(href)').extract())

        new_categories = []
        
        for index in range(len(categories)):
            
            categories[index] = categories[index].replace('\r\n','').split('                        ')
            
            child_in_new_categories = []
            
            for e in range(len(categories[index])):
                
                categories[index][e] = categories[index][e].strip()

                if categories[index][e] != '':
                
                    child_in_new_categories.append(categories[index][e])

            new_categories.append(child_in_new_categories)

        count = 0
        categories = []
        for index in range(len(new_categories)):

            if count == 0:

                categories.append(new_categories[index])

                count = len(new_categories[index])

            count -= 1

        # print(categories)
        # print(links)
        
        category_id = 0
        category_data = []
        links = links[0]
        ready_crawl_links = []
        # write categories into the excel and filter exsiting links
        for category_group in categories:

            if len(category_group) > 1:
                parent_id = category_id

            for index in range(len(category_group)):
                
                if index == 0:
                    
                    category = ('', category_group[index], str(category_id))
                
                else:
                    
                    category = (str(parent_id), category_group[index], str(category_id))
                    
                    ready_crawl_links.append(links[category_id])
                
                category_id += 1
                
                category_data.append(category)
        
        for row_index in range(len(category_data)):
            for column_index in range(len(category_data[row_index])):
                value = category_data[row_index][column_index]
                category_sheet.write(row_index, column_index, value)

        crawled_book.save('visionchart_crawled.xls')

        # if the category link list is not empty.
        while len(ready_crawl_links) > 0:
            next_page = ready_crawl_links.pop()
            next_page = response.urljoin(next_page)
            yield scrapy.Request(next_page, callback=self.parse_category)

    def parse_category(self, response):
        
        product_links = []

        for product in response.css('div.clearfix.colelem.shared_content'):

            product_links = product.css('span.actAsInlineDiv.normal_text a::attr(href)').extract()

            print(product_links)
                
        # while len(product_links) > 0:
        #     next_page = product_links.pop()
        #     next_page = response.urljoin(next_page)
        #     yield scrapy.Request(next_page, callback=self.parse_products)     

    # def parse_products(self, response):  

    #     rexcel = open_workbook('visionchart_crawled.xls')
    #     row = rexcel.sheets()[0].nrows
    #     excel = copy(rexcel)
    #     table = excel.get_sheet(0)
    #     item_category = response.css('ul.breadcrumb li span::text').extract()[-2]
        
    #     for product_info in response.css('div.col-sm-9'):
                
    #         item_sku = product_info.css('ul.list-unstyled.description li span#uo_sku_model::text').extract_first()
    #         item_mode = product_info.css('ul.list-unstyled.description li span#uo_sku_model::text').extract_first()
    #         item_otherid = ''
    #         item_name = product_info.css('h1::text').extract_first().strip()
    #         # print(item_name)
    #         item_description = product_info.css('span#tab-description p::text').extract()
    #         item_description = '\n'.join(item_description)
    #         # print(item_description)
    #         item_photo_name = product_info.css('img::attr(src)').extract_first().split('/')[-1].strip()
    #         # print(item_photo_name)
    #         item_photo_link = product_info.css('img::attr(src)').extract()
    #         item_photo_link = '\n'.join(item_photo_link)
    #         # print(item_photo_link)
    #         item_price = product_info.css('span#uo_price::text').extract_first()
    #         # print(item_price)
            
    #         item = (item_sku, 
    #                 item_mode, 
    #                 item_otherid, 
    #                 item_name, 
    #                 item_description, 
    #                 item_photo_name,
    #                 item_photo_link,
    #                 item_category,
    #                 item_price)

    #         for index in range(len(item)):
    #             table.write(row, index, item[index])

    #     excel.save('visionchart_crawled.xls')