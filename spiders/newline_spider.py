import scrapy
from xlwt import *

class NewlineSpider(scrapy.Spider):

    name = "newline"

    def start_requests(self):

        urls = [
            'https://www.newlineofficefurniture.com.au/',
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
        # operate the link string, check the 
        for link_group_index in range(len(links)):
            for link_index in range(len(links[link_group_index])):
                link_splited = links[link_group_index][link_index].split('/')
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
                    category_id += 1
                    parent_str = ''
                    for index in range(len(parents)):
                        if index > 0:
                            parent_str += ','
                        parent_str += parents[index]
                    
                    category = (parent_str, categories[link_group_index][link_index].strip(), str(category_id))

                    previous_len = current_len

                category_data.append(category)

        for row_index in range(len(category_data)):
            for column_index in range(len(category_data[row_index])):
                value = category_data[row_index][column_index]
                category_sheet.write(row_index, column_index, value)

        crawled_book.save('newline_crawled.xls')