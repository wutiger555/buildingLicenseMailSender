# 建築商情資料爬去/定期寄信通知
Medium:https://medium.com/%E4%B8%80%E5%80%8B%E8%B3%87%E7%AE%A1%E5%AD%B8%E7%94%9F%E7%9A%84%E7%A7%81%E5%AF%86%E8%99%95/%E5%BB%BA%E7%AF%89%E5%95%86%E6%83%85%E8%B3%87%E6%96%99-%E5%BB%BA%E7%AF%89%E5%9F%B7%E7%85%A7-python%E7%88%AC%E8%9F%B2%E4%B8%A6%E9%80%8F%E9%81%8Esmtp%E4%BB%A5outlook%E5%AF%84%E4%BF%A1%E9%80%9A%E7%9F%A5-a414549c9ff4

Python Modules：
> beautifulsoup4==4.9.3
> certifi==2020.12.5
> chardet==4.0.0
> et-xmlfile==1.0.1
> idna==2.10
> lxml==4.6.3
> numpy==1.20.2
> openpyxl==3.0.7
> pandas==1.2.3
> python-dateutil==2.8.1
> pytz==2021.1
> requests==2.25.1
> six==1.15.0
> soupsieve==2.2.1
> urllib3==1.26.4
> XlsxWriter==1.3.8
## source:
> https://cpabm.cpami.gov.tw/e_help.jsp

![](https://i.imgur.com/XsgU6fv.png)

![](https://i.imgur.com/GE0u1CL.png)

## result:
> (filter:總樓層數降冪排列, 變更次數<4, 總樓層數>2)

![](https://i.imgur.com/IaO5BFf.png)

## mail sender result:
> using: smtplib.SMTP(host="smtp.office365.com", port="587")

![](https://i.imgur.com/J2gA1Y5.png)

