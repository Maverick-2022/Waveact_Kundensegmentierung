import scrapy
from scrapy.linkextractors import LinkExtractor
from scrapy.spiders import CrawlSpider, Rule
from openpyxl import Workbook


class CryptorankSpider(CrawlSpider):
    name = 'crawler_crypto'
    allowed_domains = ['cryptorank.io']
    start_urls = ['https://cryptorank.io/funding-rounds']

    rules = (
        Rule(LinkExtractor(allow=r'/funding-rounds'), callback='parse_item', follow=True),
    )

    def __init__(self, *args, **kwargs):
        super(CryptorankSpider, self).__init__(*args, **kwargs)
        self.results = []

    def parse_item(self, response):
        project = response.css('sc-e17fa6bf-2 dImrEl sticky::text').get()
        date = response.css('sc-e17fa6bf-2 dImrEl::text').get()
        raise_amount = response.css('sc-e17fa6bf-2 dImrEl::text').get()
        stage = response.css('sc-e17fa6bf-2 dImrEl::text').get()
        funds_investors = response.css('sc-e17fa6bf-2 dImrEl::text').getall()
        category = response.css('.sc-e17fa6bf-2 dImrEl::text').get()

        self.results.append({
            'Project': project,
            'Date': date,
            'Raise': raise_amount,
            'Stage': stage,
            'Funds and Investors': funds_investors,
            'Category': category
        })

    def closed(self, reason):
        wb = Workbook()
        ws = wb.active
        ws.append(['Project', 'Date', 'Raise', 'Stage', 'Funds and Investors', 'Category'])

        for result in self.results:
            ws.append([
                result['Project'],
                result['Date'],
                result['Raise'],
                result['Stage'],
                ", ".join(result['Funds and Investors']),
                result['Category']
            ])

        wb.save('D:\cryptorank_data.xlsx')
        self.log('Saved data to cryptorank_data.xlsx')
