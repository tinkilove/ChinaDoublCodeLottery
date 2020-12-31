# 爬虫获取双色球的全部开奖数据
# 开奖信息保存到本地文件内，txt文件和excel文件同时保存
# 使用class，
# 格式：
import urllib.request
import platform
from bs4 import BeautifulSoup
import os
import sys
import inspect
import operator
import time
import re
import shutil
import platform
import smtplib
import threading
import xlwt
import xlrd
from xlutils.copy import copy
import pandas as pd
from datetime import date, datetime, timedelta
import calendar
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.mime.text import MIMEText
from email.header import Header
#import openpyxl
# 1-5
# 6-10
# 11-16
# 17-22
# 23-28
# 29-33
# 开奖期号/开奖日期/星期/红球/蓝球/和值/奇数数量/偶数数量/各个段内红球个数/与上一期重复的球

FILE_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_DIR = os.path.dirname(FILE_DIR)  # 找到父级目录的父级目录
TEMP_DIR = os.path.dirname(PYTHON_DIR)  # 找到父级目录的父级目录
if platform.system().lower() == 'windows':
    TEMP_DIR = TEMP_DIR + "\\tempfile\\"
else:
    TEMP_DIR = TEMP_DIR + "/tempfile/"
sys.path.append(TEMP_DIR)  # 添加环境变量

if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

CONST_MAX_NR = 0xFFFF
CONST_SPLIT_CHAR = '/'

# encoding: utf-8
g_lunar_month_day = [
    0x00752, 0x00ea5, 0x0ab2a, 0x0064b, 0x00a9b, 0x09aa6, 0x0056a, 0x00b59, 0x04baa, 0x00752,  # 1901 ~ 1910
    0x0cda5, 0x00b25, 0x00a4b, 0x0ba4b, 0x002ad, 0x0056b, 0x045b5, 0x00da9, 0x0fe92, 0x00e92,  # 1911 ~ 1920
    0x00d25, 0x0ad2d, 0x00a56, 0x002b6, 0x09ad5, 0x006d4, 0x00ea9, 0x04f4a, 0x00e92, 0x0c6a6,  # 1921 ~ 1930
    0x0052b, 0x00a57, 0x0b956, 0x00b5a, 0x006d4, 0x07761, 0x00749, 0x0fb13, 0x00a93, 0x0052b,  # 1931 ~ 1940
    0x0d51b, 0x00aad, 0x0056a, 0x09da5, 0x00ba4, 0x00b49, 0x04d4b, 0x00a95, 0x0eaad, 0x00536,  # 1941 ~ 1950
    0x00aad, 0x0baca, 0x005b2, 0x00da5, 0x07ea2, 0x00d4a, 0x10595, 0x00a97, 0x00556, 0x0c575,  # 1951 ~ 1960
    0x00ad5, 0x006d2, 0x08755, 0x00ea5, 0x0064a, 0x0664f, 0x00a9b, 0x0eada, 0x0056a, 0x00b69,  # 1961 ~ 1970
    0x0abb2, 0x00b52, 0x00b25, 0x08b2b, 0x00a4b, 0x10aab, 0x002ad, 0x0056d, 0x0d5a9, 0x00da9,  # 1971 ~ 1980
    0x00d92, 0x08e95, 0x00d25, 0x14e4d, 0x00a56, 0x002b6, 0x0c2f5, 0x006d5, 0x00ea9, 0x0af52,  # 1981 ~ 1990
    0x00e92, 0x00d26, 0x0652e, 0x00a57, 0x10ad6, 0x0035a, 0x006d5, 0x0ab69, 0x00749, 0x00693,  # 1991 ~ 2000
    0x08a9b, 0x0052b, 0x00a5b, 0x04aae, 0x0056a, 0x0edd5, 0x00ba4, 0x00b49, 0x0ad53, 0x00a95,  # 2001 ~ 2010
    0x0052d, 0x0855d, 0x00ab5, 0x12baa, 0x005d2, 0x00da5, 0x0de8a, 0x00d4a, 0x00c95, 0x08a9e,  # 2011 ~ 2020
    0x00556, 0x00ab5, 0x04ada, 0x006d2, 0x0c765, 0x00725, 0x0064b, 0x0a657, 0x00cab, 0x0055a,  # 2021 ~ 2030
    0x0656e, 0x00b69, 0x16f52, 0x00b52, 0x00b25, 0x0dd0b, 0x00a4b, 0x004ab, 0x0a2bb, 0x005ad,  # 2031 ~ 2040
    0x00b6a, 0x04daa, 0x00d92, 0x0eea5, 0x00d25, 0x00a55, 0x0ba4d, 0x004b6, 0x005b5, 0x076d2,  # 2041 ~ 2050
    0x00ec9, 0x10f92, 0x00e92, 0x00d26, 0x0d516, 0x00a57, 0x00556, 0x09365, 0x00755, 0x00749,  # 2051 ~ 2060
    0x0674b, 0x00693, 0x0eaab, 0x0052b, 0x00a5b, 0x0aaba, 0x0056a, 0x00b65, 0x08baa, 0x00b4a,  # 2061 ~ 2070
    0x10d95, 0x00a95, 0x0052d, 0x0c56d, 0x00ab5, 0x005aa, 0x085d5, 0x00da5, 0x00d4a, 0x06e4d,  # 2071 ~ 2080
    0x00c96, 0x0ecce, 0x00556, 0x00ab5, 0x0bad2, 0x006d2, 0x00ea5, 0x0872a, 0x0068b, 0x10697,  # 2081 ~ 2090
    0x004ab, 0x0055b, 0x0d556, 0x00b6a, 0x00752, 0x08b95, 0x00b45, 0x00a8b, 0x04a4f, ]
 
# 农历数据 每个元素的存储格式如下：
#    12~7         6~5    4~0
#  离元旦多少天  春节月  春节日
#####################################################################################
g_lunar_year_day = [
    0x18d3, 0x1348, 0x0e3d, 0x1750, 0x1144, 0x0c39, 0x15cd, 0x1042, 0x0ab6, 0x144a,  # 1901 ~ 1910
    0x0ebe, 0x1852, 0x1246, 0x0cba, 0x164e, 0x10c3, 0x0b37, 0x14cb, 0x0fc1, 0x1954,  # 1911 ~ 1920
    0x1348, 0x0dbc, 0x1750, 0x11c5, 0x0bb8, 0x15cd, 0x1042, 0x0b37, 0x144a, 0x0ebe,  # 1921 ~ 1930
    0x17d1, 0x1246, 0x0cba, 0x164e, 0x1144, 0x0bb8, 0x14cb, 0x0f3f, 0x18d3, 0x1348,  # 1931 ~ 1940
    0x0d3b, 0x16cf, 0x11c5, 0x0c39, 0x15cd, 0x1042, 0x0ab6, 0x144a, 0x0e3d, 0x17d1,  # 1941 ~ 1950
    0x1246, 0x0d3b, 0x164e, 0x10c3, 0x0bb8, 0x154c, 0x0f3f, 0x1852, 0x1348, 0x0dbc,  # 1951 ~ 1960
    0x16cf, 0x11c5, 0x0c39, 0x15cd, 0x1042, 0x0a35, 0x13c9, 0x0ebe, 0x17d1, 0x1246,  # 1961 ~ 1970
    0x0d3b, 0x16cf, 0x10c3, 0x0b37, 0x14cb, 0x0f3f, 0x1852, 0x12c7, 0x0dbc, 0x1750,  # 1971 ~ 1980
    0x11c5, 0x0c39, 0x15cd, 0x1042, 0x1954, 0x13c9, 0x0e3d, 0x17d1, 0x1246, 0x0d3b,  # 1981 ~ 1990
    0x16cf, 0x1144, 0x0b37, 0x144a, 0x0f3f, 0x18d3, 0x12c7, 0x0dbc, 0x1750, 0x11c5,  # 1991 ~ 2000
    0x0bb8, 0x154c, 0x0fc1, 0x0ab6, 0x13c9, 0x0e3d, 0x1852, 0x12c7, 0x0cba, 0x164e,  # 2001 ~ 2010
    0x10c3, 0x0b37, 0x144a, 0x0f3f, 0x18d3, 0x1348, 0x0dbc, 0x1750, 0x11c5, 0x0c39,  # 2011 ~ 2020
    0x154c, 0x0fc1, 0x0ab6, 0x144a, 0x0e3d, 0x17d1, 0x1246, 0x0cba, 0x15cd, 0x10c3,  # 2021 ~ 2030
    0x0b37, 0x14cb, 0x0f3f, 0x18d3, 0x1348, 0x0dbc, 0x16cf, 0x1144, 0x0bb8, 0x154c,  # 2031 ~ 2040
    0x0fc1, 0x0ab6, 0x144a, 0x0ebe, 0x17d1, 0x1246, 0x0cba, 0x164e, 0x1042, 0x0b37,  # 2041 ~ 2050
    0x14cb, 0x0fc1, 0x18d3, 0x1348, 0x0dbc, 0x16cf, 0x1144, 0x0a38, 0x154c, 0x1042,  # 2051 ~ 2060
    0x0a35, 0x13c9, 0x0e3d, 0x17d1, 0x11c5, 0x0cba, 0x164e, 0x10c3, 0x0b37, 0x14cb,  # 2061 ~ 2070
    0x0f3f, 0x18d3, 0x12c7, 0x0d3b, 0x16cf, 0x11c5, 0x0bb8, 0x154c, 0x1042, 0x0ab6,  # 2071 ~ 2080
    0x13c9, 0x0e3d, 0x17d1, 0x1246, 0x0cba, 0x164e, 0x10c3, 0x0bb8, 0x144a, 0x0ebe,  # 2081 ~ 2090
    0x1852, 0x12c7, 0x0d3b, 0x16cf, 0x11c5, 0x0c39, 0x154c, 0x0fc1, 0x0a35, 0x13c9,  # 2091 ~ 2100
]
 
# ==================================================================================
 

# 开始年份
START_YEAR = 1901
 
month_DAY_BIT = 12
month_NUM_BIT = 13
 
# 　todo：正月初一 == 春节   腊月二十九/三十 == 除夕
yuefeng = ["正月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "冬月", "腊月"]
riqi = ["初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", "初九", "初十",
        "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "廿十",
        "廿一", "廿二", "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十"]
 
xingqi = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
 
tiangan = ["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"]
dizhi = ["子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"]
shengxiao = ["鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊", "猴", "鸡", "狗", "猪"]
 
def change_year(num):
    dx = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    tmp_str = ""
    # 将年份 转换为字符串，然后进行遍历字符串 ，将字符串中的数字转换为中文数字
    for i in str(num):
        tmp_str += dx[int(i)]
    return tmp_str
 
# 获取星期
def week_str(tm):
    return xingqi[tm.weekday()]
 
# 获取天数
def lunar_day(day):
    return riqi[(day - 1) % 30]
 
 
def lunar_day1(month, day):
    if day == 1:
        return lunar_month(month)
    else:
        return riqi[day - 1]
 
# 判断是否是闰月
def lunar_month(month):
    leap = (month >> 4) & 0xf
    m = month & 0xf
    month = yuefeng[(m - 1) % 12]
    if leap == m:
        month = "闰" + month
    return month
 
#求什么年份，中国农历的年份和 什么生肖年
def lunar_year(year):
    return tiangan[(year - 4) % 10] + dizhi[(year - 4) % 12] + '[' + shengxiao[(year - 4) % 12] + ']'
 
 
# 返回：
# a b c
# 闰几月，该闰月多少天 传入月份多少天
def lunar_month_days(lunar_year, lunar_month):
    if (lunar_year < START_YEAR):
        return 30
 
    leap_month, leap_day, month_day = 0, 0, 0  # 闰几月，该月多少天 传入月份多少天
 
    tmp = g_lunar_month_day[lunar_year - START_YEAR]
 
    if tmp & (1 << (lunar_month - 1)):
        month_day = 30
    else:
        month_day = 29
 
        # 闰月
    leap_month = (tmp >> month_NUM_BIT) & 0xf
    if leap_month:
        if (tmp & (1 << month_DAY_BIT)):
            leap_day = 30
        else:
            leap_day = 29
 
    return (leap_month, leap_day, month_day)
 
 
# 算农历日期
# 返回的月份中，高4bit为闰月月份，低4bit为其它正常月份
def get_ludar_date(tm):
    year, month, day = tm.year, 1, 1
    code_data = g_lunar_year_day[year - START_YEAR]
    days_tmp = (code_data >> 7) & 0x3f
    chunjie_d = (code_data >> 0) & 0x1f
    chunjie_m = (code_data >> 5) & 0x3
    span_days = (tm - datetime(year, chunjie_m, chunjie_d)).days
    # print("span_day: ", days_tmp, span_days, chunjie_m, chunjie_d)
 
    # 日期在该年农历之后
    if (span_days >= 0):
        (leap_month, foo, tmp) = lunar_month_days(year, month)
        while span_days >= tmp:
            span_days -= tmp
            if (month == leap_month):
                (leap_month, tmp, foo) = lunar_month_days(year, month)  # 注：tmp变为闰月日数
                if (span_days < tmp):  # 指定日期在闰月中
                    month = (leap_month << 4) | month
                    break
                span_days -= tmp
            month += 1  # 此处累加得到当前是第几个月
            (leap_month, foo, tmp) = lunar_month_days(year, month)
        day += span_days
        return year, month, day
        # 倒算日历
    else:
        month = 12
        year -= 1
        (leap_month, foo, tmp) = lunar_month_days(year, month)
        while abs(span_days) >= tmp:
            span_days += tmp
            month -= 1
            if (month == leap_month):
                (leap_month, tmp, foo) = lunar_month_days(year, month)
                if (abs(span_days) < tmp):  # 指定日期在闰月中
                    month = (leap_month << 4) | month
                    break
                span_days += tmp
            (leap_month, foo, tmp) = lunar_month_days(year, month)
        day += (tmp + span_days)  # 从月份总数中倒扣 得到天数
        return year, month, day
 
# 打印 某个时间的农历
def _show_month(tm):
    (year, month, day) = get_ludar_date(tm)
    #print("%d年%d月%d日" % (tm.year, tm.month, tm.day), week_str(tm), end='')
    #print("\t农历 %s年 %s年%s%s " % (lunar_year(year), change_year(year), lunar_month(month), lunar_day(day)))  # 根据数组索引确定
    return (year, month, day)

# 判断输入的数据是否符合规则
def show_month(year, month, day):
    if year > 2100 or year < 1901:
        return
    if month > 13 or month < 1:
        return
 
    tmp = datetime(year, month, day)
    (year, month, day) = _show_month(tmp)
    return str(year)+"-"+str(month)+"-"+str(day)
 
 
# 显示现在的日期
def this_month():
    show_month(datetime.now().year, datetime.now().month, datetime.now().day)


class FetchDoubleBallFromNet():
    def __init__(self, _iBallLimit=154, _iMaxDayLimit=365):
        self.m_strUrlPart = 'http://kaijiang.zhcw.com/zhcw/inc/ssq/ssq_wqhg.jsp?pageNum='
        self.m_strBeginUrl = 'http://kaijiang.zhcw.com/zhcw/html/ssq/list_1.html'
        self.m_iBallTotalPage = 0  # 网站号码总数量
        self.m_iBallTotalCount = 0  # 实际总开奖数量
        self.m_iEveryPageCount = 20  # 每页的记录数
        self.m_iBallLimit = _iBallLimit  # 获取双色球的号码个数上限
        self.m_iFetchedBallNr = 0  # 已经获取的双色球数量
        self.m_iMaxDayLimit = -(_iMaxDayLimit*1)  # 获取记录为向前N年内的开奖记录，超过的不再需要
        self.m_iPagePerThread = 10
        self.m_strResPath = TEMP_DIR + "doubleball.xls"
        self.m_bDebug = True
        self.lock = threading.Lock()
    # ==============================================================================

    def __cPrint(self, _strContext):
        if self.m_bDebug:
            print(_strContext)
    # ==============================================================================

    def initSysType(self):
        self.m_strSysType = platform.system()
        self.__cPrint(("Current OS is:", self.m_strSysType))
    # ==============================================================================

    def __urlOpen(self, _strUrl):
        try:
            req = urllib.request.Request(_strUrl)
            req.add_header(
                'User-Agent', 'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6')
            html = urllib.request.urlopen(req).read()
            time.sleep(0.2)
            return html
        except:
            self.__cPrint(('error:'+_strUrl))
    # ===============================================================================
    # 获取url总页数

    def __getTotalPageNum(self, _strUrl):
        if len(_strUrl) == 0:
            return 0
        num = 0
        page = self.__urlOpen(_strUrl)
        soup = BeautifulSoup(page, "lxml")
        strong = soup.find('td', colspan='7')
        if strong:
            result = strong.get_text().split(' ')
            list_num = re.findall("[0-9]{1}", result[1])
            for i in range(len(list_num)):
                num = num*10 + int(list_num[i])
            self.__cPrint(str("__getPageNum = " + str(num)))
            return num
        else:
            return 0
    # ===============================================================================
    # 获取开奖号码总数

    def __getBallTotalCount(self, _strUrl):
        if len(_strUrl) == 0:
            return 0
        num = 0
        page = self.__urlOpen(_strUrl)
        soup = BeautifulSoup(page, "lxml")
        strong = soup.find('td', colspan='7')
        if strong:
            result = strong.get_text().split(' ')
            list_num = re.findall("[0-9]{1}", result[3])
            for i in range(len(list_num)):
                num = num*10 + int(list_num[i])
            self.__cPrint(str("__getBallTotalCount = " + str(num)))
            return num
        else:
            return 0
    # ===============================================================================
    # 1-5
    # 6-10
    # 11-16
    # 17-22
    # 23-28
    # 29-33
    # 获取红球的分布区间

    def __get_red_pos(self, strCode):
        lstcode = strCode.split(',')
        numbers = list(map(int, lstcode))
        redpos = {'5': 0, '10': 0, '16': 0, '22': 0, '28': 0, '33': 0}
        for code in numbers:
            for di in redpos.keys():
                if code <= int(str(di)):
                    redpos[di] = redpos[di] + 1
                    break
        return redpos

    # 获取红球的总和
    def __get_red_sum(self, strCode):
        redsum = 0
        lstcode = strCode.split(',')
        numbers = list(map(int, lstcode))
        for code in numbers:
            redsum += (code)
        return redsum

    # 获取红球的奇偶个数
    def __get_red_sd(self, strCode):
        sigrednr = 0
        doubrednr = 0
        lstcode = strCode.split(',')
        numbers = list(map(int, lstcode))
        for code in numbers:
            if (code) % 2 != 0:
                sigrednr += 1
        doubrednr = 6 - sigrednr
        return sigrednr, doubrednr
    # 设置表格样式

    def __set_style(self, name, height, bold=False):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = name
        font.bold = bold
        font.color_index = 4
        font.height = height
        style.font = font
        return style
    # 写Excel

    def __WriteSheetRow(self, sheet, rowValueList, rowIndex, isBold):
        i = 0
        # xlwt.easyxf('font: bold 1')
        style = self.__set_style('Times New Roman', 220, isBold)
        # style = xlwt.easyxf('font: bold 0, color red;')#红色字体
        # style2 = xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: bold on;') # 设置Excel单元格的背景色为黄色，字体为粗体
        for svalue in rowValueList:
            if isBold:
                sheet.write(rowIndex, i, svalue, style)
            else:
                if ('-' not in svalue and ',' not in svalue):
                    sheet.write(rowIndex, i, int(svalue))
                else:
                    sheet.write(rowIndex, i, svalue)
            i = i + 1

    def __write_excel(self, lstreport):

        self.lock.acquire()
        lstitem = list()
        _strFilePath = self.m_strResPath

        data = xlrd.open_workbook(_strFilePath, formatting_info=True)
        excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
        excel_table = excel.get_sheet(0)  # 获得要操作的页
        table = data.sheets()[0]
        nrows = table.nrows  # 获得行数
        ncols = table.ncols  # 获得列数
        rowIndex = nrows

        for item in lstreport:
            self.__cPrint(item)
            lstitem = item.strip('\n').split(CONST_SPLIT_CHAR)
            valueList = []
            for i in range(0, len(lstitem)):
                valueList.append(str(lstitem[i]))
            self.__WriteSheetRow(excel_table, valueList, rowIndex, False)
            rowIndex = rowIndex + 1
        excel.save(_strFilePath)

        self.lock.release()
    #获取农历日期
    def __getnongli_date(self, strDateTime):
        dtItemDate = datetime.strptime(
                    strDateTime, '%Y-%m-%d')
        return show_month(dtItemDate.year, dtItemDate.month, dtItemDate.day)

    # 从网上获取开奖号码
    def __fetch_ball_code(self, _threadName, _dtLimitDay, _endpage):
        lstContent = list()
        lstreport = list()
        bOverRun = False
        _startpage = 1
        if _endpage <= self.m_iPagePerThread:
            _startpage = 1
        else:
            _startpage = _endpage - self.m_iPagePerThread
        print("获取页面范围:" + str(_startpage), ":", str(_endpage))

        #_strFilePath = TEMP_DIR + self.m_strTempFile + str(_startpage)
        # with open(_strFilePath, 'w+') as fp:
        for iPage in range(_startpage, _endpage, 1):  # 1-6=>1,2,3,4,5
            if bOverRun == True:
                break
            lstContent = self.__getBallContentByPage(iPage)
            for each in lstContent:
                strDateTime = str(each.strip('\n').split(CONST_SPLIT_CHAR)[0])

                strRedCode = str(each.strip('\n').split(CONST_SPLIT_CHAR)[2])

                dtItemDate = datetime.strptime(
                    strDateTime, '%Y-%m-%d')
                #strDateTime = dtItemDate.strftime('%Y-%m-%d')
                strweekday = dtItemDate.weekday() + 1

                redsum = self.__get_red_sum(strRedCode)

                sigrednr, doubrednr = self.__get_red_sd(strRedCode)

                redpos = self.__get_red_pos(strRedCode)

                if dtItemDate <= _dtLimitDay or self.m_iFetchedBallNr >= self.m_iBallLimit:
                    bOverRun = True
                    break

                res = str(each) + CONST_SPLIT_CHAR + \
                    str(strweekday) + CONST_SPLIT_CHAR + \
                    str(redsum) + CONST_SPLIT_CHAR + \
                    str(sigrednr) + CONST_SPLIT_CHAR + \
                    str(doubrednr)

                for item in redpos.values():
                    res += CONST_SPLIT_CHAR + str(item)
                
                res += CONST_SPLIT_CHAR + str(self.__getnongli_date(strDateTime))
                res += "\n"

                lstreport.append(res)
                # fp.write(each)
                # fp.write(res)
                self.m_iFetchedBallNr += 1
                time.sleep(0.1)
            # fp.flush()
            time.sleep(1)
        self.__write_excel(lstreport)

    # 开奖期号/开奖日期/红球/蓝球/星期/和值/奇数数量/偶数数量/各个段内红球个数/与上一期重复的球
    # ===============================================================================
    # 1-5
    # 6-10
    # 11-16
    # 17-22
    # 23-28
    # 29-33
    def __createNew(self, _dtLimitDay):
        # 垃圾资源清理
        _strFilePath = self.m_strResPath

        if os.path.exists(_strFilePath):
            os.remove(_strFilePath)

        outputfile = xlwt.Workbook()
        sheet = outputfile.add_sheet('doublecode', cell_overwrite_ok=True)
        rowIndex = 0
        headList = ['开奖日期', '开奖期号', '红球', '蓝球', '星期几',
                    '红球和值',
                            '红奇数',
                            '红偶数',
                            '1-5个数',
                            '6-10个数',
                            '11-16个数',
                            '17-22个数',
                            '23-28个数',
                            '29-33个数',
                            '农历日期'
                    ]
        self.__WriteSheetRow(sheet, headList, rowIndex, True)
        outputfile.save(_strFilePath)

        theads = []
        for iPage in range(1, self.m_iBallTotalPage + 1):  # 按照5页进行迭代
            if iPage % self.m_iPagePerThread == 0:
                # _thread.start_new_thread(
                #    self.__fetch_ball_code, ("fetch_ball_code", _dtLimitDay, iPage + 1))
                tp = threading.Thread(target=self.__fetch_ball_code, args=(
                    "fetch_ball_code", _dtLimitDay, iPage + 1))
                theads.append(tp)
            # 最后一批剩余不足5页情况
            if (self.m_iBallTotalPage - iPage) < self.m_iPagePerThread:
                break

        if iPage < self.m_iBallTotalPage:  # 全部不足5页情况
            # _thread.start_new_thread(
            #    self.__fetch_ball_code, ("fetch_ball_code", _dtLimitDay, self.m_iBallTotalPage + 1))
            tp = threading.Thread(target=self.__fetch_ball_code, args=(
                "fetch_ball_code", _dtLimitDay, self.m_iBallTotalPage + 1))
            theads.append(tp)

        for t in theads:
            t.start()

        for t in theads:
            t.join()
    # ===============================================================================

    # 获取每页双色球的信息 2018-07-08:2018078:03,10,14,17,18,30,12

    def __getBallContent(self):
        # 获取当前的日期，时间，月
        dtNow = datetime.now()
        dtTimeSpan = timedelta(days=self.m_iMaxDayLimit)
        dtLimitDay = dtNow + dtTimeSpan  # 得到新的日期,2年前的今天，txt里面保留这些日期的内容

        dtLimitDay = datetime.strptime("1970-01-01", '%Y-%m-%d')

        self.m_iBallTotalPage = self.__getTotalPageNum(self.m_strBeginUrl)
        self.m_iBallTotalCount = self.__getBallTotalCount(self.m_strBeginUrl)
        self.__createNew(dtLimitDay)
    # ==============================================================================
    # 获取指定页码的双色球的信息，并进行计算和分析
    # 1-5
    # 6-10
    # 11-16
    # 17-22
    # 23-28
    # 29-33
    # 开奖期号/开奖日期/红球/蓝球/星期/和值/奇数数量/偶数数量/各个段内红球个数/与上一期重复的球

    def __getBallContentByPage(self, _iPageNo):
        if _iPageNo == 0:
            return
        href = self.m_strUrlPart + str(_iPageNo)  # + '.html'  # 调用新url链接
        # for listnum in len(list_num):
        page = BeautifulSoup(self.__urlOpen(href), "lxml")
        time.sleep(0.2)
        em_list = page.find_all('em')  # 匹配em内容
        # 匹配 <td align=center>这样的内容
        div_list = page.find_all('td', {'align': 'center'})
        # 匹配 <td align=center>这样的内容
        num_list = page.find_all('td', {'align': 'center'})
        # 初始化
        strCodeNoList = list()  # 开奖期号
        dtDatetimeList = list()  # 开奖日期
        strRedBallCodeList = list()  # 开奖号码
        strBlueBallCodeList = list()  # 开奖号码
        strDataList = list()
        # 开奖号码
        strCode = ''
        n = 0
        for div in em_list:
            text = div.get_text()
            text = text.encode('utf-8')
            n = n + 1
            if n == 7:
                text = text.decode()
                strCode += text
                strRedBallCodeList.append(
                    str(strCode)[0: len(str(strCode)) - 3])
                strBlueBallCodeList.append(str(strCode)[-2:])
                strCode = ''
                n = 0
            else:
                text = text.decode() + ","
                strCode += text
        # 开奖日期
        for div2 in div_list:  # <td align="center">2018-06-24</td>
            text = div2.get_text().strip('')
            # print text
            list_num = re.findall(r'\d{4}-\d{2}-\d{2}', text)
            list_num = str(list_num[::1])
            list_num = list_num[2:12]
            if len(list_num) == 0:
                continue
            elif len(list_num) > 1:
                dtDatetimeList.append(str(list_num))
        # 开奖期号
        for div in num_list:  # <td align="center">2018072</td>
            text = div.get_text().strip('')
            list_num1 = re.findall(r'\d{7}', text)
            list_num1 = str(list_num1[::1])
            list_num1 = list_num1[2:9]
            if len(list_num1) == 0:
                continue
            elif len(list_num1) > 1:
                strCodeNoList.append(str(list_num1))
        # i = 0
        for i in range(len(dtDatetimeList)):
            strDataList.append(str(dtDatetimeList[i]) + CONST_SPLIT_CHAR +
                               str(strCodeNoList[i]) + CONST_SPLIT_CHAR +
                               str(strRedBallCodeList[i]) + CONST_SPLIT_CHAR +
                               str(strBlueBallCodeList[i])
                               )
        # i = i + 1
        return strDataList
     # 开奖期号/开奖日期/红球/蓝球/星期/和值/奇数数量/偶数数量/各个段内红球个数/与上一期重复的球
    # ==============================================================================
    # 对外接口，触发调用，获取开奖号码
    # _iCreateType：0-新建，1-扩展
    # _iLimitEnable：0-全部开奖号码，1-默认上限期数的开奖号码

    def GetBallDataFromNet(self):
        self.__getBallContent()
    # ===============================================================================

    def ExcelSort(self):
        _strFilePath = self.m_strResPath
        dfexcel = pd.read_excel(_strFilePath)
        dfexcel.sort_values(by='开奖日期', ascending=False, inplace=True)
        dfexcel.to_excel(_strFilePath, sheet_name='双色球',
                         encoding='utf-8', index=False)
        print(dfexcel)

    def GetExportFile(self):
        return self.m_strResPath
    # ===============================================================================

if __name__ == "__main__":
    ballget = FetchDoubleBallFromNet(CONST_MAX_NR, CONST_MAX_NR)  # 开奖信息获取对象
    ballget.initSysType()
    ballget.GetBallDataFromNet()        # 重新获取新的数据
    ballget.ExcelSort()                 # 重新进行排序
