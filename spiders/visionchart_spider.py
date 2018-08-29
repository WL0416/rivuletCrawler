import scrapy

class NewlineSpider(scrapy.Spider):

    name = "visionchat"

    def start_requests(self):

        urls = [
            'http://www.visionchart.com.au/',
        ]

        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        page = response.url.split('/')[-2]
        filename = 'pages/visionchat-%s.html' % page
        with open(filename, 'wb') as f:
            f.write(response.body)
        self.log('Saved file %s' % filename)