import smtplib
import requests
import pandas as pd
import logging
import logging.config
import sys
from bs4 import BeautifulSoup as bs
from config import email, password
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

logging.basicConfig(
    filename='main.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def mailInit():
    content = MIMEMultipart()  #建立MIMEMultipart物件
    content["subject"] = "建築商情資料整理"  #郵件標題
    content["from"] = "schindler.jtpkb@outlook.com"  #寄件者
    content["to"] = "wutiger555@gmail.com" #收件者
    
    try:
        xlsx = MIMEApplication(open("df.xlsx", 'rb').read())
    except Exception as e:
        logging.error(e)
        sys.exit(e)
    xlsx.add_header('Content-Disposition', 'attachment', filename= "建照.xlsx")
    content.attach(xlsx)
    html = """
    <html lang="zh-TW">
        <head>
            <meta charset="UTF-8">
        </head>
        <body>
            <p>附件為近3個月之全台縣市<a href="https://cpabm.cpami.gov.tw/e_help.jsp">建築商情資料</a>（建造執照）</p>
            <p>&nbsp;</p>
            <p><img src="https://gm1.ggpht.com/cgt9rFCch69zq2bmR8uzcX4pROiT4BEeDzLsHKSSNPwhS3J8lX5x8mo-WSONctoR5tZ_A4qpQnjKqJJ8Kslg_xbPlQpTFNOXSblkt5DPOmGoxF4bCu-l5ExSHSiKKt2nf6oe_7EhMksNGgfEkSc94qYlGVuARc7OyvnMePNbMgi3vTTceKbqyIx7W9rJBb3AyryPrBOvq25pE_0TdqMeVlerJ_8HqDFoPYWqL-b-8drdnoAXMOwuYQfdj_AS_jTyZzDdqqk1-0b-5MRy9f3E_vkKz65coxJAGQ9rFWaG1SN0DPbNEL4SkUsB28lcj__PnyaGJTsqzyFoOkXS6rarJOb85ylnR7ZDfFQ6rbGq3HIwWgDNbN2i6qkn7a9t0maifUz7BRGAVvsHTP25DJZtxWAX4-kcP32QlxX1PtPr8hsxGz-oGqebTe2vIO0YoNl-XJe0Gq_r2HCxWrdg1sCBHi-H8Nrp5TzrN8nKrJQ87gMAoGnvZC9TpDdrNksXGK8rUR13cUxFbsR7nmNUObVUiemv4YlQ9YRTZIGDhBt8yNsFX6BlXjEp8_x6h9xYgH1CHsg-w4XG-YJCKbCMC6owfPUgU_yxUMesKj9lklZ9XW5HqT2JtjtxqAGLutFwkhVPS6B0OyEg-SkjIlY1FQvT5FE3SYkp_mL2N4jM7evEtOIVo_X65h7YUiINhlhP8Exw2eRXj2Bza-FcQe5NLNAclZPqrkIed0m9VrqyFgxoHB4c0G7WZLuPU1yXj9O0j45TSnIG=s0-l75-ft-l75-ft" alt="" /></p>
            <p>Jardine Schindler Lifts Limited | Jardine Schindler Lifts Ltd (Taiwan)<br />9/F, 35 Kwang Fu South Road, Taipei 105, Taiwan<br /><br /><a href="http://www.schindler.com/tw" target="_blank" rel="noopener" data-saferedirecturl="https://www.google.com/url?q=http://www.schindler.com/tw&amp;source=gmail&amp;ust=1617928856258000&amp;usg=AFQjCNFYRvNdS4hbdPhjDjh5Nb7IW8B3Jw">www.schindler.com/tw</a><br /><br /><span style="color: #ff0000;">We Elevate&nbsp;</span></p>
        </body>
    </html>
    """
    logging.info('email info: \nsubject:{}\nfrom:{}to:\n'.format(content['subject'],content['from'],content['to']))
    body = MIMEText(html, 'html')
    content.attach(body)
    return content
def sendMail(content):
    logging.info('Sendding Mail Process Start.')
    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
        try:
            smtp.ehlo()  # 驗證SMTP伺服器
            smtp.starttls()  # 建立加密傳輸
            smtp.login(email, password)  # 登入寄件者
            smtp.send_message(content)  # 寄送郵件
            logging.info('Mail Sending Completed.')
        except Exception as e:
            logging.error(e)

def request():
    url = "https://cpabm.cpami.gov.tw/e_help.jsp"
    r = requests.get(url)
    soup = bs(r.text, 'html.parser')
    trs = soup.select("tr[class*=list]")
    links = []

    for tr in trs:
        tds = tr.select("td")
        for i,td in enumerate(tds):
            if td.find('a') and i==1:
                links.append(td.find('a').get('href'))
    table = pd.read_html(links[0], header=0)
    df = table[0]
    logging.info('request processing start.')
    for i,link in enumerate(links,start=1):
        logging.info('requesting: {}/{}'.format(i,len(links)))
        if(i==1):
            continue
        t = pd.read_html(link, header=0)
        df = df.append(t[0], ignore_index=True)
    logging.info('request processing finished.')
    df=df.set_index('建造號碼')
    df1 = df
    df1['總樓層數'] = df1['地下層數']+df1['地上層數']
    df1 = df[(df['總樓層數']>2)&(df['變更次數']<4)]
    df1 = df1.sort_values(['總樓層數'], ascending=(False))
    try:
        df1.to_excel("df.xlsx", index=False)
    except Exception as e:
        logging.error(e)
        sys.exit(e)
    try:
        writer = pd.ExcelWriter('df.xlsx', engine = 'xlsxwriter')
        df1.to_excel(writer, sheet_name='Sheet1')
        # Get the xlsxwriter objects from the dataframe writer object.
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        # Apply the autofilter based on the dimensions of the dataframe.
        worksheet.autofilter(0, 1, df.shape[0], df.shape[1])
        writer.save()
    except Exception as e:
        logging.error(e)

try:
    request()
except Exception as e:
    logging.error(e)
    sys.exit(e)
content = mailInit()
sendMail(content)