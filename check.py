from lxml import html
import requests
import urllib.request

mvcr_url = 'http://www.mvcr.cz/mvcren/'

page = requests.get(mvcr_url + 'article/status-of-your-application.aspx')
tree = html.fromstring(page.content)

file_link = tree.xpath('//div[@id="content"]/div/ul/li/a/@href')

mvcr_file = mvcr_url + file_link[0]

urllib.request.urlretrieve (mvcr_file, "file.xls")
