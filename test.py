import os
from xlfuns import xlApp
from gethistory import D3DataFilter
import requests
from lxml import etree
from datetime import datetime
import time
import random

def getdata():
    book = r"f:\ssq\3d 删减版\3d.xlsm"
    # TODO 打开的是含有宏的 excel 文件，如果设置了自动运行，那么两这会冲突
    # 这里应该要有一个判断的函数或者，最好可以是一个函数判断 excel 进程
    # 状态，因为已经出过问题了，导致两这都在写入数据，毁了原始数据不说，
    # 还出错了，要么就是用个锁
    # 暂时没有找到号方法，用最后修改时间来控制吧，如果昨天21:15 之后没有更新，
    # 那么就 sleep 15s
    xl = xlApp(book = book,visible=False)
    sht_3d = xl.active_wb.sheets['3d']

    max_row_num = 2 + sht_3d.range('A3').expand(mode='down').count
    row_start = 3
    src_addr = f'A{row_start}:G{max_row_num}'
    data = sht_3d.range(src_addr).value

    d3 = D3DataFilter(data=data)

    # 需要 workbook name ,从path 中提取
    xl.close_book(os.path.basename(book))
    book_to_write = r'C:\Users\z_hao\Desktop\3d\对比.xlsx'
    xl.open(book_to_write)
    # 要写入的 sheet 也改名为3d
    sht_to_write = xl.active_wb.sheets['3d']

    # 要写入数据的位置
    dic_pos = {
        "sjh_zt":("d1","b5"), # 昨天试机号跟随，b5 是数据位置，d1 是标题
        "sjh_jt":("t1","r5"), # 今天试机号跟随
        "kjh_jt_gs_1":("ag1","ah5"), # 开奖跟随1
        "kjh_jt_gs_2":("ag1","ay5"), # 开奖跟随2
        "kjh_zt_2":("bq1","bo5"), # 昨天开将号跟随2
        "kjh_zj":("d25","a28"), # 最近开奖走势
        "kjh_zt":("t25","R28"), # 昨天开奖号跟随
        "lstq_jt":("aj25","ah28"), # 历史同期
        "lstq_zt":("ba25","ay28"), # 历史同期上期
    }

    str_today = datetime.now().strftime('%Y-%m-%d')
    sjh_update_time = datetime.strptime(str_today + " 18:01:00",'%Y-%m-%d %H:%M:%S')
    time_now = datetime.now()

    if time_now > sjh_update_time:
        forecast_issue,forecast_sjh = get_forecast_issue_and_sjh()
    else:
        forecast_issue = ""
        forecast_sjh = ""

    row_count_1=20  # label pos 是第一行的数据能使用的总行数
    column_count_1 = 1
    row_count_2 = d3.current_year - d3.start_year +1

    # 开奖号跟随
    kjh_jt_gs1,kjh_jt_gs2 = d3.get_kjh_gensui()
    label_pos,data_pos = dic_pos['kjh_jt_gs_1']
    sht_to_write.range(label_pos).value = d3.get_last_nth_kjh(1)

    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_1,
                                             column_count_1)).clear_contents()
    sht_to_write.range(data_pos).options(transpose=True).value = kjh_jt_gs1
    label_pos,data_pos = dic_pos['kjh_jt_gs_2']
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_1,
                                             column_count_1)).clear_contents()
    sht_to_write.range(data_pos).options(transpose=True).value = kjh_jt_gs2


    # 开奖号前一期跟随
    kjh_zt = d3.get_kjh_previous_gensui()
    label_pos,data_pos = dic_pos['kjh_zt']
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_2,
                                             column_count_1)).clear_contents()
    sht_to_write.range(label_pos).value = d3.get_last_nth_kjh(2)
    sht_to_write.range(data_pos).options(transpose=True).value = kjh_zt

    # TODO开奖号昨天的第二期数据
    kjh_zt_2 = d3.get_kjh_previous_gensui2()
    label_pos,data_pos = dic_pos['kjh_zt_2']
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_1,
                                             column_count_1)).clear_contents()
    sht_to_write.range(label_pos).value ="上期开奖跟随 "+ \
        str(d3.get_last_nth_kjh(2))
    sht_to_write.range(data_pos).options(transpose=True).value = kjh_zt_2


    # TODO开奖数据最近
    kjh_zj = d3.get_last_n_sjh_kjh()
    label_pos,data_pos = dic_pos['kjh_zj']
    # 最近开奖数据是两列的
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_2, 2)).clear_contents()
    last_issue = d3.get_last_nth_issue(1)
    sht_to_write.range(label_pos).value = f'最近开奖号{last_issue}'
    # sht_to_write.range(data_pos).options(transpose=True).value = kjh_zj
    sht_to_write.range(data_pos).value = kjh_zj


    # 最新试机号的历史开奖数据，要等到最新试机号出来才能使用
    # TODO 函数没有问题，差获取试机号的函数和时间控制，或者获取失败时候的逻辑

    # 今天最新试机号跟随
    # TODO 处理时间问题

    if forecast_issue != "" and forecast_sjh != "":
        sjh_jt = d3.get_sjh_gensui(forecast_sjh)
        label_pos,data_pos = dic_pos['sjh_jt']
        sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_1,
                                                column_count_1)).clear_contents()
        sht_to_write.range(label_pos).value = forecast_sjh
        sht_to_write.range(data_pos).options(transpose=True).value = sjh_jt

    # 昨天的试机号跟随
    sjh_zt = d3.get_sjh_previous_gensui()
    label_pos,data_pos = dic_pos['sjh_zt']
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_1,
                                             column_count_1)).clear_contents()
    sht_to_write.range(label_pos).value = d3.get_last_nth_sjh(1)
    sht_to_write.range(data_pos).options(transpose=True).value = sjh_zt


    # 历史上的今天出号
    lstq_jt = d3.get_lishichuhao(forecast_issue)
    label_pos,data_pos = dic_pos['lstq_jt']
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_2,
                                             column_count_1)).clear_contents()
    sht_to_write.range(label_pos).value = forecast_issue[-3:]
    sht_to_write.range(data_pos).options(transpose=True).value = lstq_jt

    # 历史上的昨天出号
    lstq_zt = d3.get_previous_lishichuhao()
    label_pos,data_pos = dic_pos['lstq_zt']
    sht_to_write.range(d3.excel_cell_expand(data_pos,row_count_2,
                                             column_count_1)).clear_contents()
    sht_to_write.range(label_pos).value = "历史同期上" + \
                str(d3.get_last_nth_issue(1))[-3:]
    sht_to_write.range(data_pos).options(transpose=True).value = lstq_zt


    # 关闭excel
    xl.close_book(os.path.basename(book_to_write))
    xl.appclose()


def my_get_response_from_url(url):
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) \
        AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.3"}
    response = requests.get(url,headers=headers)
    # 停留 一段时间
    time.sleep(0.5+random.random())
    return response



def get_forecast_issue_and_sjh():
    # 获取要开奖的期号和试机号
    # 试机号只有18:00 之后才正确
    # //a[@class="fb"]
    url = "https://www.17500.cn/3d.html"

    response = my_get_response_from_url(url)
    if response.status_code == 200:
        parse_html = etree.HTML(response.text)
        label_a = parse_html.xpath('//a[@class="fb"]')[0]
        issue = label_a.xpath('./text()')[0][:7]
        sjh = label_a.xpath('./font/text()')[0]
        return (issue,sjh)

getdata()
