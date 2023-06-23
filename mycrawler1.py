import scrapy
import pandas as pd
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings


class CrawlMasterSpider(scrapy.Spider):
    name = "mycrawler1"
    allowed_domains = ["cryptorank.io"]
    start_urls = ["https://cryptorank.io/funding-rounds?sort=fullName&rows=500&page=2"]

    def parse(self, response):
        table = response.css("table")  # Wählen Sie die Tabelle aus

        Project = table.css('.sc-e17fa6bf-4 > tr:nth-child(1) > td:nth-child(1) > div:nth-child(1) > a:nth-child(1) > div:nth-child(1) > div:nth-child(2) > span:nth-child(1)').get()
        Date = table.css('th.sc-e17fa6bf-2:nth-child(2) > abbr:nth-child(1)').get()
        Raise = table.css('.sc-e17fa6bf-4 > tr:nth-child(1) > td:nth-child(3) > span:nth-child(1)').get()
        Stage = table.css('.sc-e17fa6bf-4 > tr:nth-child(1) > td:nth-child(4) > span:nth-child(1)').get()
        Funds_and_Investors = table.css('th.sc-e17fa6bf-2:nth-child(5) > abbr:nth-child(1)').get()
        Category = table.css('th.sc-e17fa6bf-2:nth-child(6) > abbr:nth-child(1)').get()

        data = {
            'Project': Project,
            'Date': Date,
            'Raise': Raise,
            'Stage': Stage,
            'Funds_and_Investors': Funds_and_Investors,
            'Category': Category
        }

        self.results.append(data)  # Hinzufügen der Daten zur results-Liste

    def __init__(self, *args, **kwargs):
        self.results = []
        super().__init__(*args, **kwargs)

    def closed(self, reason):
        df = pd.DataFrame(self.results)
        print("DataFrame:", df)
        excel_datei = "C:/Users/tirya/Downloads/tek1/tek1.xlsx"

        if excel_datei:
            try:
                df.to_excel(excel_datei, index=False)
                print("Daten wurden erfolgreich in die Excel-Datei geschrieben.")
            except Exception as e:
                print("Fehler beim Schreiben der Daten in die Excel-Datei:", e)
        else:
            print("Keine Excel-Datei angegeben.")
        return df


if __name__ == "__main__":
    process = CrawlerProcess(get_project_settings())
    process.crawl(CrawlMasterSpider)
    process.start()
