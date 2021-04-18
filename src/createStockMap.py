#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import pandas as pd
import lxml.html
import sys
import time
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import strftime
from utility import createBook,send_message,parse_dom_tree

TARGET_DROP_RATE = 0.5

if __name__ == '__main__':
    source_path = 'data_collection_210418.xlsx'
    data_frame = pd.read_excel(source_path,usecols=['コード','銘柄名','権利確定月','優待内容'],sheet_name='Sheet1')

    dt_now = datetime.now()
    tmp = dt_now + relativedelta(months=1)
    dt_month = str(dt_now.month)
    dt_next_month = str(tmp.month)

    idx=0
    output_list = []
    for index,row in data_frame.iterrows():
        stock_code = row[0] # コード
        stock_name = str(row[1]) # 銘柄名
        stock_vesting_date = str(row[2]) # 権利確定月
        if stock_vesting_date == "随時":
            continue
        else:
            month_list = re.findall("\d+月",stock_vesting_date)
            vesting_month_list = [s.replace("月","") for s in month_list]
            if dt_month in vesting_month_list or dt_next_month in vesting_month_list:
                print(stock_code)
            else:
                continue

        stock_benefits = str(row[3]) # 優待内容
        chart_url = "https://minkabu.jp/stock/{}/chart".format(str(stock_code))
        
        try:
            chart_html = requests.get(chart_url)
            chart_html.raise_for_status()
            chart_dom_tree = lxml.html.fromstring(chart_html.content)

            # 現在値
            tmp = chart_dom_tree.xpath('//*[@id="layout"]/div[2]/div[3]/div[2]/div/div[1]/div/div/div[1]/div[2]/div/div[2]/div/text()')
            current_price = float(re.sub(r'\D','',tmp[0]))

            # 前日終値
            stock_previous_closing_place = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[1]/td[1]','円','')            
            price_drop_rate = (1 - current_price / stock_previous_closing_place) * 100
            if price_drop_rate >= TARGET_DROP_RATE:
                output_list.append([stock_code,stock_name,current_price,stock_previous_closing_place,price_drop_rate])

        except requests.exceptions.RequestException as e:
            print(e)
    
    message = ""
    for idx, data_list in enumerate(output_list):
        message += "【銘柄名:{}】(コード:{})\n".format(data_list[1],data_list[0])
        message += "現在値:{:.0f}\n".format(data_list[2])
        message += "前日終値:{:.0f}\n".format(data_list[3])
        message += "急落率:{:.1f}%\n\n".format(data_list[4])
    send_message(message)
