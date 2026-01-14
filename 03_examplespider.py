import scrapy
from scrapy.spiders import CrawlSpider, Rule
from scrapy.linkextractors import LinkExtractor
import re
from WEBSITEcrawler.items import WEBSITEcrawlerItem

class WebsiteSpider(CrawlSpider):
    name = 'WEBSITE_SPIDER_NAME'
    allowed_domains = ['WEBSITE_DOMAIN']
    start_urls = ['WEBSITE_URL']

    rules = [Rule(LinkExtractor(allow=()), callback="parse", follow=True)]

    custom_settings={'FEED_URI': 'WEBSITE_SPIDER_NAME_Pipeline.csv', 'FEED_FORMAT': 'csv'}

    def parse(self, response):
        webpage = WebsiteSpidercrawlerItem()
        webpage['URL']=response.url
        webpage['Source']='WEBSITE_NAME'
        webpage['Title']=response.css('title::text').extract()
        webpage['Bodytext']=response.css('p::text').extract()
        yield webpage
