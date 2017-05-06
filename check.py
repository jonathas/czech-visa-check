import json
import smtplib
import urllib.request
import requests
import xlrd
from lxml import html

class CzechVisaCheck:
    """Checks if an application number is in the current approved ones"""

    mvcr_url = 'http://www.mvcr.cz/mvcren/'
    mvcr_filename = 'applications.xls'

    def __init__(self):
        config = self.load_config_file()
        self.download_sheet_file(self.extract_sheet_url())
        result = self.is_application_there(config)
        self.sendmail(config, result)

    def load_config_file(self):
        """Open the json file with some parameters"""
        with open('config.json') as config_file:
            return json.load(config_file)

    def extract_sheet_url(self):
        """Extracts the current application sheet url from the page"""
        page = requests.get(self.mvcr_url + 'article/status-of-your-application.aspx')
        tree = html.fromstring(page.content)
        file_link = tree.xpath('//div[@id="content"]/div/ul/li/a/@href')
        return self.mvcr_url + file_link[0]

    def download_sheet_file(self, mvcr_file_url):
        """Downloads the sheet from the url"""
        urllib.request.urlretrieve(mvcr_file_url, self.mvcr_filename)

    def is_application_there(self, config):
        """Compares the application in the config with the list of retrieved applications"""
        application = config['application']
        book = xlrd.open_workbook(self.mvcr_filename)

        for sheet in book.sheets():
            print("Checking sheet named {0}".format(sheet.name))
            for row in range(sheet.nrows):
                if application == sheet.cell_value(rowx=row, colx=1).strip():
                    print(config['success_msg'].format(application))
                    return True

        print(config['fail_msg'].format(application))
        return False

    def sendmail(self, config, result):
        """Sends an email with the result"""
        smtp = config['smtp']
        header = 'From: %s\n' % smtp['from']
        header += 'To: %s\n' % smtp['to']
        header += 'Subject: %s\n\n' % smtp['subject']

        if result:
            body = config['success_msg'].format(config['application'])
        else:
            body = config['fail_msg'].format(config['application'])

        message = header + body

        server = smtplib.SMTP(smtp['host'] + ':' + smtp['port'])
        server.starttls()
        server.login(smtp['user'], smtp['password'])
        problems = server.sendmail(smtp['from'], smtp['to'], message)
        server.quit()
        if not problems:
            print("E-Mail sent successfully.")


czech_this_out = CzechVisaCheck()
