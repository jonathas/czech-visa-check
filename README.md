# Czech visa check

## Note: This script was useful when the Ministry of Interior used to have their updates in a xls file. Since they started working with their portal (https://frs.gov.cz/en) this script is not needed anymore.

The Ministry of Interior in the Czech Republic updates [this](http://www.mvcr.cz/mvcren/article/status-of-your-application.aspx) page every week with the visa application numbers that have been approved.

The problem is that it's not a system and there's no way to search for your application number on that page. They update an Excel file every week, so you have to download it and CTRL+F the file yourself, looking for your application number.

Because of that, I decided to create this script to do this task for me and optimize my time. I can configure it to run for example once a week on cron and it will do the job for me and then send me an e-mail letting me know if my application number is there or not yet.

## Configuring

In order to configure it, open the config.json file and enter your application number and then your SMTP config for it to send you an e-mail with the result.

## Running

```bash
pip install -r requirements.txt
python check.py
```
