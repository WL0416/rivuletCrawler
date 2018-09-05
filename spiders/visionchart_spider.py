import scrapy
from xlwt import *
import os
from xlutils.copy import copy
from xlrd import open_workbook
from scrapy.selector import Selector as selector
import urllib
import time

class NewlineSpider(scrapy.Spider):

    name = "visionchart"

    def start_requests(self):

        urls = [
            'http://www.visionchart.com.au/',
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):

        page = response.url.split('/')[-1]
        
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
        # print(ready_crawl_links)
        while len(ready_crawl_links) > 0:
            next_page = ready_crawl_links.pop()
            next_page = response.urljoin(next_page)
            print('This is next cate page: ' + next_page)
            yield scrapy.Request(next_page, callback=self.parse_category)

    def parse_category(self, response):

        products = response.css('span.actAsInlineDiv.normal_text a::attr(href)').extract()

        while len(products) > 0:
            next_page = products.pop()
            next_page = response.urljoin(next_page)
            print('This is next prod page: ' + next_page)
            yield scrapy.Request(next_page, callback=self.parse_products)

    def parse_products(self, response):  

        rexcel = open_workbook('visionchart_crawled.xls')
        row = rexcel.sheets()[0].nrows
        excel = copy(rexcel)
        table = excel.get_sheet(0)

        item_name = response.css('h3.H3::text').extract_first()
        item_category = response.css('a.nonblock::text').extract()[-1].strip()

        item_photo_link = response.css('img.block::attr(data-src)').extract()

        if not os.path.exists('images/visionchart/' + item_name):
            os.makedirs('images/visionchart/' + item_name)

        image_count = 0
        for link in item_photo_link:
            
            image_link = 'http://www.visionchart.com.au' + link

            urllib.urlretrieve(image_link, 'images/visionchart/' + item_name + '/' + item_name + '_' + str(image_count) + '.png')

            time.sleep(1)

            image_count += 1

        links_in_excel = '\n'.join(item_photo_link)

        item_description = None
        for item in response.css('div.clearfix.grpelem'):
            if len(item.css('h3.H3').extract()) > 0:
                item_description = item.xpath('string(.)').extract()
        
        description = ''
        item_sku = []
        item_size = []
        seperate_des = []
        rest_des = []
        description_finish = False
        column_num = 0
        devident = 1
        if item_description is not None:
            for des in item_description:
                des = des.replace('\r\n','').split('       ')
                for d in des:
                    if len(d) > 0:
                        d = d.strip()
                        if d == 'CODE':
                            
                            description_finish = True
                            continue

                        if not description_finish and d is not item_name:
                            
                            description += d
                            continue
                        
                        if d == 'DESCRIPTION':
                            devident += 1
                            continue

                        d = d.replace(u'\xa0', u' ')

                        if d == 'SIZE  (mm)':
                            devident +=1
                            continue

                        rest_des.append(d + ',')

        # stop = len(rest_des) / devident
        # list_count = 0
        # for info_index in range(len(rest_des)):

        #     if list_count == 0:

        #         item_sku.append(rest_des[info_index])
            
        #     if info_index > 0 and (info_index + 1) % devident == 0:

        #         list_count += 1

        #     if list_count == 1:
                
        #         seperate_des.append(rest_des[info_index])
        #         continue

        #     if list_count == 2:
                
        #           item_size.append(rest_des[info_index])
        #     # print(item_photo_link)
        #     item_price = product_info.css('span#uo_price::text').extract_first()
        #     # print(item_price)
            
        item = (#item_sku,
                # item_mode, 
                # item_otherid, 
                item_name, 
                description, 
                # item_photo_name,
                links_in_excel,
                item_category,
                rest_des,
                # item_price
                )

        item_name = ('Name', 'Description', 'Images', 'Category', 'SKU & SIZE')

        for index in range(len(item)):
            
            if row == 0:
                table.write(row, index, item_name[index])
                table.write(row + 1, index, item[index])
            else:
                table.write(row, index, item[index])

        excel.save('visionchart_crawled.xls')