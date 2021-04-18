#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import pandas as pd
import lxml.html
import time
from time import strftime
from utility import createBook

if __name__ == '__main__':
    # Obtain the code and stock information listed on the First Section of the Tokyo Stock Exchange from https://www.jpx.co.jp/. 
    source_path = 'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls'
    data_frame = pd.read_excel(source_path,usecols=['コード','銘柄名','市場・商品区分','17業種区分'],sheet_name='Sheet1')
    stock_data_frame = data_frame[data_frame['市場・商品区分']=='市場第一部（内国株）']

    idx=0
    output_list = []
    for index,row in stock_data_frame.iterrows():
        stock_code = row[0]
        stock_name = str(row[1])
        stock_industry_type = str(row[3])
        print(stock_code)
        url = "https://kabuoji3.com/stockholder/{}/".format(stock_code)
        headers = {
            "User-Agent": "Chrome"
        }
        try:
            html = requests.get(url, headers=headers)
            html.raise_for_status()
            dom_tree = lxml.html.fromstring(html.content)

            # 権利確定月
            vesting_date = dom_tree.xpath('//*[@id="base_box"]/div/div[3]/dl[2]/dd/text()')
            if not len(vesting_date):
                continue
            stock_vesting_date = vesting_date[0]

            # 優待内容
            benefits = dom_tree.xpath('//*[@id="base_box"]/div/div[3]/dl[1]/dd/text()')
            stock_benefits = benefits[0]

            print(stock_vesting_date)
            stock_list = [stock_code,\
                            stock_name,\
                            stock_vesting_date,\
                            stock_benefits,\
                            stock_industry_type]
            output_list.append(stock_list)
        except requests.exceptions.RequestException as e:
            print(e)
    columns = ['コード',\
                '銘柄名',\
                '権利確定月',\
                '優待内容',\
                '業種'
                ]

    localtime = strftime("%y%m%d",time.localtime())
    output_file_name = './data_collection_' + localtime + '.xlsx'
    output_data_frame = pd.DataFrame(data=output_list,columns=columns)
    createBook(output_data_frame,output_file_name,'B2')