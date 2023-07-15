import xlwings as xw
import numpy as np
from datetime import datetime
import pandas as pd
app = xw.App()
workbook = app.books.open('/Users/ze/Desktop/23S执行率计算表.xlsx')
worksheet = workbook.sheets['收益测算']
worksheet1 = workbook.sheets['时刻']
rng = worksheet.range('A1').expand('table')  # 收益测算表
rng1 = worksheet1.range('A1').expand('table')  # 时刻正表
flight_list = worksheet1.range('A3').expand('down').value  # 时刻正表里的航班号，保持更新
warning_airports = ['北京', '首都',' 大兴', '上海', '浦东', '虹桥', '广州', '白云', '成都', '双流', '成都双流', '成都天府',
                    '天府', '双流', '深圳', '昆明', '西安', '重庆', '杭州', '南京', '郑州', '厦门', '武汉', '长沙', '青岛',
                    '海口', '乌鲁木齐', '天津', '贵阳', '哈尔滨', '沈阳', '三亚', '大连', '济南', '南宁', '兰州',
                    '福州', '太原', '长春', '南昌', '呼和浩特', '呼和']
def main():
    time_calculate = pd.read_excel('/Users/ze/Desktop/23S执行率计算表.xlsx', sheet_name=1)
    time_calculate[['班期']] = time_calculate[['班期']].astype(str)
    for i in rng.value[1:]:  # 遍历从第二行开始的每一行，到最后一行
        if i[2][0:6] not in flight_list:
            print(i[2][0:6])
            continue
        if len(i[2]) <= 6:  # 航班号类似GJ8888
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            d_list = [datetime.date(x).month for x in d]
            for d_month in range(1, 13):
                if d_list.count(d_month) == 0:  # 如果每个月的数字是0，就忽略
                    continue
                else:
                    d_1 = []
                    for x_1 in d:
                        if datetime.date(x_1).month == d_month:
                            d_1.append(x_1)
                    month_main(d_1, d_month, i[2], i[6])
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1

        elif len(i[2]) == 8:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            d_list = [datetime.date(x).month for x in d]
            for d_month in range(1, 13):
                if d_list.count(d_month) == 0:  # 如果每个月的数字是0，就忽略
                    continue
                else:
                    d_1 = []
                    for x_1 in d:
                        if datetime.date(x_1).month == d_month:
                            d_1.append(x_1)
                    month_main(d_1, d_month, i[2], i[6])
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:5]+i[2][7]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1
        elif len(i[2]) == 9:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            d_list = [datetime.date(x).month for x in d]
            for d_month in range(1, 13):
                if d_list.count(d_month) == 0:  # 如果每个月的数字是0，就忽略
                    continue
                else:
                    d_1 = []
                    for x_1 in d:
                        if datetime.date(x_1).month == d_month:
                            d_1.append(x_1)
                    month_main(d_1, d_month, i[2], i[6])
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:4]+i[2][7:9]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1


        elif len(i[2]) == 10:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            d_list = [datetime.date(x).month for x in d]
            for d_month in range(1, 13):
                if d_list.count(d_month) == 0:  # 如果每个月的数字是0，就忽略
                    continue
                else:
                    d_1 = []
                    for x_1 in d:
                        if datetime.date(x_1).month == d_month:
                            d_1.append(x_1)
                    month_main(d_1, d_month, i[2], i[6])
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:3]+i[2][7:10]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1


        elif len(i[2]) == 11:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            d_list = [datetime.date(x).month for x in d]
            for d_month in range(1, 13):
                if d_list.count(d_month) == 0:  # 如果每个月的数字是0，就忽略
                    continue
                else:
                    d_1 = []
                    for x_1 in d:
                        if datetime.date(x_1).month == d_month:
                            d_1.append(x_1)
                    month_main(d_1, d_month, i[2], i[6])
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:6]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1
            for datetime_everyday in d:
                everyday = str(datetime_everyday.weekday()+1)
                if i[6] == "取消":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 0
                elif i[6] == "取前":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                        time_calculate.航站 == '航站2' ), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站3'), datetime_everyday] = 0
                elif i[6] == "取后":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0.5
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站1'), datetime_everyday] = 0
                elif i[6] == "拉直":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)) & (
                                               time_calculate.航站 == '航站2'), datetime_everyday] = 0
                elif i[6] == "还原":
                    time_calculate.loc[(time_calculate.航班号 == i[2][0:2]+i[2][7:11]) & (
                        time_calculate['班期'].fillna(" ").str.contains(everyday, na=False)), datetime_everyday] = 1
    time_calculate_list = np.array(time_calculate.iloc[1:, 15:]).tolist()
    worksheet1.range('P3').value = time_calculate_list
    workbook.save()
    read_result()

def month_main(d_list, month, flightnumber, statement):
    month = str(month) + "月"
    worksheet_month = workbook.sheets[month]
    flight_list_month = worksheet_month.range('A3').expand('down').value  # 时刻正表里的航班号，保持更新
    time_list_month = workbook.sheets[month].range('A1').expand('table').value[0][4:-2]  # 时刻表里的时间日期，换季了再更新！平时不要动
    if len(flightnumber) <= 6:  # 航班号类似GJ8888
        if flightnumber in flight_list_month:
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          # 右偏移n+1个格子，下偏移n+1
                                                          row_offset=flight_list_month.index(flightnumber) + 1)  # 锁定单元格
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
        else:
            print(flightnumber)
    elif len(flightnumber) == 8:
        if flightnumber[0:6] in flight_list_month:
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(flightnumber[0:6]) + 1)
                if statement == "还原":
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    cell.value = ['X']
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(
                                                              flightnumber[0:5] + flightnumber[7]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
        else:
            print(flightnumber)

    elif len(flightnumber) == 9:
        if flightnumber[0:6] in flight_list_month:
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(flightnumber[0:6]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(
                                                              flightnumber[0:4] + flightnumber[7:9]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
        else:
            print(flightnumber)

    elif len(flightnumber) == 10:
        if flightnumber[0:6] in flight_list_month:
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(flightnumber[0:6]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(
                                                              flightnumber[0:3] + flightnumber[7:10]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
        else:
            print(flightnumber)
    elif len(flightnumber) == 11:
        if flightnumber[0:6] in flight_list_month:
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(flightnumber[0:6]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
            for b in d_list:
                cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(b) + 1,
                                                          row_offset=flight_list_month.index(
                                                              flightnumber[0:2] + flightnumber[7:11]) + 1)
                if statement == "还原":
                    if cell.value is None:
                        continue
                    cell.value = [1]  # 取消具有绝对权限
                else:
                    if cell.value is None:
                        continue
                    cell.value = ['X']
        else:
            print(flightnumber)

def read_result():
    last_eight_row = workbook.sheets['总汇'].range('J1').expand('table').value
    SanDaChang_amont = [last_eight_row[1][1], last_eight_row[1][3], last_eight_row[1][5], last_eight_row[1][7],
                        last_eight_row[1][9],
                        last_eight_row[1][11], last_eight_row[1][13], last_eight_row[1][15], last_eight_row[1][17],
                        last_eight_row[1][19],
                        last_eight_row[1][21], last_eight_row[1][23]]  # 1月到12月的三大场正班数量，随时更新
    zhengban_amont = [last_eight_row[5][1], last_eight_row[5][3], last_eight_row[5][5], last_eight_row[5][7],
                      last_eight_row[5][9],
                      last_eight_row[5][11], last_eight_row[5][13], last_eight_row[5][15], last_eight_row[5][17],
                      last_eight_row[5][19],
                      last_eight_row[5][21], last_eight_row[5][23]]  # 1月到12月的总正班数量，随时更新
    Sandachang_zhixinglv = [last_eight_row[0][1], last_eight_row[0][3], last_eight_row[0][5], last_eight_row[0][7],
                            last_eight_row[0][9],
                            last_eight_row[0][11], last_eight_row[0][13], last_eight_row[0][15],
                            last_eight_row[0][17],
                            last_eight_row[0][19],
                            last_eight_row[0][21], last_eight_row[0][23]]  # 1月到12月的三大场执行率，随时更新
    zhengban_zhixinglv = [last_eight_row[4][1], last_eight_row[4][3], last_eight_row[4][5], last_eight_row[4][7],
                          last_eight_row[4][9],
                          last_eight_row[4][11], last_eight_row[4][13], last_eight_row[4][15],
                          last_eight_row[4][17],
                          last_eight_row[4][19],
                          last_eight_row[4][21], last_eight_row[4][23]]  # 1月到12月的正班执行率，随时更新
    Allban = []  # 每一次报取总共正班数计算空列表
    SanDaChang_name = ["浦东", "广州", "北京"]  # 三大场名字，注意新航季可能需要更新！比如加了大兴或者虹桥
    sandachang = []  # 每一次报取三大场总班数计算空列表
    for i in rng.value[1:]:  # 遍历从第二行开始的每一行，到最后一行
        if i[2][0:6] not in flight_list:
            continue
        if len(i[2]) <= 6:  # 航班号类似GJ8888
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            Allban.extend(d)
            if any(field in i[1] for field in SanDaChang_name):
                sandachang.extend(d)

        elif len(i[2]) == 8:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            Allban.extend(d)
            Allban.extend(d)
            if any(field in i[1] for field in SanDaChang_name):
                sandachang.extend(d)
                sandachang.extend(d)

        elif len(i[2]) == 9:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            Allban.extend(d)
            Allban.extend(d)
            if any(field in i[1] for field in SanDaChang_name):
                sandachang.extend(d)
                sandachang.extend(d)

        elif len(i[2]) == 10:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            Allban.extend(d)
            Allban.extend(d)
            if any(field in i[1] for field in SanDaChang_name):
                sandachang.extend(d)
                sandachang.extend(d)

        elif len(i[2]) == 11:
            if len(i[3]) > 10:  # 日期拆分
                startDay = i[3].split('-')[0]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3].split('-')[1]
                year = startDay.year
                if len(endDay) <= 5:  # 出现日期类似3.12，却没有加上年的后半部分
                    endDay = str(year)+str(".")+endDay
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
                elif len(endDay) >= 6:
                    endDay = datetime.strptime(endDay, "%Y.%m.%d")
            elif len(i[3]) <= 10:
                startDay = i[3]
                startDay = datetime.strptime(startDay, "%Y.%m.%d")
                endDay = i[3]
                endDay = datetime.strptime(endDay, "%Y.%m.%d")
            t = pd.date_range(start=startDay, end=endDay, freq="D")
            d = []
            for x in t:
                for a in str(i[4]):
                    if a == ".":
                        continue
                    else:
                        c = int(a) - 1
                    if x.weekday() == c:
                        d.append(x)
            Allban.extend(d)
            Allban.extend(d)
            if any(field in i[1] for field in SanDaChang_name):
                sandachang.extend(d)
                sandachang.extend(d)
    list1 = []
    YUJING_eighty = []
    rng_column = rng.columns[2].value
    for index, value in enumerate(rng_column):
        if len(value) == 6:
            list1.append(value)
        elif len(value) == 8:
            list1.append(value[0:6])
            list1.append(value[0:5] + value[7])
        elif len(value) == 9:
            list1.append(value[0:6])
            list1.append(value[0:4] + value[7:9])
        elif len(value) == 10:
            list1.append(value[0:6])
            list1.append(value[0:3] + value[7:10])
        elif len(value) == 11:
            list1.append(value[0:6])
            list1.append(value[0:2] + value[7:11])
    list2 = list(set(list1))
    list2.sort(key=list1.index)
    shuzu = [int(x[2:6]) for x in list2]
    huibao_list = []
    k = 0
    for flight in list2:
        panduan = int(flight[2:6])
        if panduan % 2 == 0:
            if panduan - 1 in shuzu:
                continue
        k += 1
        index_list = [index for index, x in enumerate(flight_list) if x == flight]
        if len(index_list) % 3 == 0:
            cell1 = worksheet1.range('E2').offset(row_offset=index_list[0] + 1).value
            cell2 = worksheet1.range('E2').offset(row_offset=index_list[1] + 1).value
            cell3 = worksheet1.range('E2').offset(row_offset=index_list[2] + 1).value
            zhixinglv1 = worksheet1.range('O2').offset(row_offset=index_list[0] + 1).value
            zhixinglv2 = worksheet1.range('O2').offset(row_offset=index_list[1] + 1).value
            zhixinglv3 = worksheet1.range('O2').offset(row_offset=index_list[2]+ 1).value

            if isinstance(zhixinglv1, float):
                zhixinglv1_format_output = format(zhixinglv1, '.2%')
            else:
                zhixinglv1_format_output = zhixinglv1
            if isinstance(zhixinglv2, float):
                zhixinglv2_format_output = format(zhixinglv2, '.2%')
            else:
                zhixinglv2_format_output = zhixinglv2
            if isinstance(zhixinglv3, float):
                zhixinglv3_format_output = format(zhixinglv3, '.2%')
            else:
                zhixinglv3_format_output = zhixinglv3
            hangxian = (str(k) + "、" + "{0}={1}={2}".format(cell1, cell2, cell3))
            huibao_list.append(hangxian)
            zhiixnglv = ("整航季：" +
                         "{0}{1}、{2}{3}、{4}{5}".format(cell1, (zhixinglv1_format_output), cell2,
                                                       (zhixinglv2_format_output),
                                                       cell3, (zhixinglv3_format_output)))
            if isinstance(zhixinglv1, float):
                if 0.30 <= zhixinglv1 < 0.8:
                    YUJING_eighty.append(cell1)
            if isinstance(zhixinglv2, float):
                if 0.30 <= zhixinglv2 < 0.8:
                    YUJING_eighty.append(cell2)
            if isinstance(zhixinglv3, float):
                if 0.30 <= zhixinglv3 < 0.8:
                    YUJING_eighty.append(cell3)
            huibao_list.append(zhiixnglv)
        elif len(index_list) % 3 != 0:
            cell1 = worksheet1.range('E2').offset(row_offset=index_list[0] + 1).value
            cell2 = worksheet1.range('E2').offset(row_offset=index_list[1] + 1).value
            zhixinglv1 = worksheet1.range('O2').offset(row_offset=index_list[0] + 1).value
            zhixinglv2 = worksheet1.range('O2').offset(row_offset=index_list[1] + 1).value
            if isinstance(zhixinglv1, float):
                zhixinglv1_format_output = format(zhixinglv1, '.2%')
            else:
                zhixinglv1_format_output = zhixinglv1
            if isinstance(zhixinglv2, float):
                zhixinglv2_format_output = format(zhixinglv2, '.2%')
            else:
                zhixinglv2_format_output = zhixinglv2
            hangxian = (str(k) + "、" + "{0}={1}".format(cell1, cell2))
            huibao_list.append(hangxian)
            zhiixnglv = ("整航季：" +
                         "{0}{1}、{2}{3}".format(cell1, (zhixinglv1_format_output), cell2,
                                                       (zhixinglv2_format_output)))
            if isinstance(zhixinglv1, float):
                if 0.30 <= zhixinglv1 < 0.8:
                    YUJING_eighty.append(cell1)
            if isinstance(zhixinglv2, float):
                if 0.30 <= zhixinglv2 < 0.8:
                    YUJING_eighty.append(cell2)
            huibao_list.append(zhiixnglv)
    for result_point in (huibao_list):
        print(result_point)

    mon = [datetime.date(x).month for x in sandachang]  # 本次报取三大场涉及的所有班数对应的月份
    mon1 = [datetime.date(x).month for x in Allban]  # 本次报取涉及的所有班数对应的月份
    mona_all = 0
    for i in range(1, 13):
        mona = mon.count(i)  # 统计本次报取三大场涉及的每一个月份的数字
        mona_all = mona_all + mona
        if mona == 0:
            continue
        else:
            print("方案涉及取消%d月份三大场正班%d班，影响执行率%.2f%%，取消后执行率为%.2f%%" %
                  (i, mona, ((mona / SanDaChang_amont[i - 1]) * 100),
                   ((Sandachang_zhixinglv[i - 1]) * 100)))
    if mona_all != 0:
        print("方案涉及取消整航季三大场正班%d班，影响执行率%.2f%%，取消后执行率为%.2f%%" %
              (mona_all, ((mona_all / last_eight_row[1][25]) * 100),
               ((last_eight_row[0][25] - mona_all / last_eight_row[1][25]) * 100)))
    for i in range(1, 13):
        monb = mon1.count(i)  # 统计本次所有正班涉及的每一个月份的数字
        if monb == 0:
            continue
        else:
            print("方案涉及取消%d月份正班%d班，影响执行率%.2f%%，取消后执行率为%.2f%%" %
                  (i, monb, ((monb / zhengban_amont[i - 1]) * 100),
                   ((zhengban_zhixinglv[i - 1]) * 100)))
    list3 = list(set(YUJING_eighty).intersection(set(warning_airports)))  # 判断预警里的机场是否处于全部的预警机场里
    list3.sort(key=warning_airports.index)
    if not list3:
        pass
    else:
        print("以上" + "、".join(list3) + "执行率低于80%，存在丢失时刻风险")  # 较低风险的预警，存在YUJING_eighty和list3中

def yufei_to_main():
    workbook1 = app.books.open('/Users/ze/Desktop/1.xlsx')
    df_list = workbook1.sheets['CDC'].range('A1').expand('table').value
    df_list = df_list[1:]
    all_list = []
    for index, row in enumerate(df_list):
        if row[0] == df_list[index-1][0]:
            if row[4] is None:
                all_list[-1][5] = all_list[-1][5] + row[1]
                all_list[-2][5] = all_list[-2][5] + row[1]
            else:
                all_list[-1][5] = all_list[-1][5] + row[1]
                all_list[-2][5] = all_list[-2][5] + row[1]
                all_list[-3][5] = all_list[-3][5] + row[1]
            continue
        else:
            if row[4] is None:
                air_line = "{0}-{1}".format(row[2], row[3])
                all_list.append([row[0], air_line, None, "航站1", row[2], row[1]])
                all_list.append([row[0], air_line, None, "航站2", row[3], row[1]])
            else:
                air_line = "{0}-{1}-{2}".format(row[2], row[3], row[4])
                all_list.append([row[0], air_line, None, "航站1", row[2], row[1]])
                all_list.append([row[0], air_line, None, "航站2", row[3], row[1]])
                all_list.append([row[0], air_line, None, "航站3", row[4], row[1]])
    worksheet1.range('A3').value = all_list
    workbook.save()

def FOC_main():
    # 抓foc.xlsx 的数据，每日更新。foc的数据要更新，地址需要改
    df = pd.read_excel('/Users/ze/Desktop/foc.xlsx', sheet_name=0, header=1, usecols=[0, 1, 3, 9, 24, 25])
    time_calculate = pd.read_excel('/Users/ze/Desktop/23S执行率计算表.xlsx', sheet_name=1)
    # 1 简单的条件筛选：单一条件筛选
    df = df[df["飞行时间"] >= 10]
    data1 = df.query('性质=="客正" | 性质=="货正"')
    data1[['航班日期']] = data1[['航班日期']].astype(str)
    data1["航班日期"] = data1.航班日期.map(lambda x:'20' +x) #foc 要加上20，因为他的格式是22-12-7
    copy_array = np.array(data1)  # 列转成array
    copy_list = copy_array.tolist()
    ws = workbook.sheets['FOC数据汇总'] #更新到表格里
    maintable = ws.range('A1').expand('table')
    maintable.clear_contents()# 清洗内容
    ws.range((1, 1)).value = copy_list
    maintable = ws.range('A1').expand('table').value #重新选择宣布数据
    #data1["航班日期"] = '20' + data1["航班日期"]
    foc_value = {}
    for row in maintable:
        flight_date = row[0]
        # 计算总时刻和经停时刻
        if flight_date not in foc_value.keys():
            foc_value[flight_date] = {}
        flight_number = row[2]
        a111 = [1, row[3], row[4]]
        if flight_number not in foc_value[flight_date].keys():
            foc_value[flight_date][flight_number] = a111
        else:
            foc_value[flight_date][flight_number][0] = 2
            foc_value[flight_date][flight_number].append(row[3])
            foc_value[flight_date][flight_number].append(row[4])

    # 批量生成foc的字典，这一步是为了
    for date_1 in foc_value.keys():
        d_month = datetime.date(date_1).month
        # 月度数据处理
        part2(d_month, date_1, foc_value[date_1])
        time_calculate[date_1] = time_calculate[date_1].map({1: 0, 0.5: 0, 0: 0})  # 替换每一列中的值为0
        for b in foc_value[date_1].keys():
            if b not in flight_list:
                print(b)
                continue
            if len(time_calculate.loc[(time_calculate.航班号 == b), "航站名称"]) % 3 == 2:
                time_calculate.loc[(time_calculate.航班号 == b) & (time_calculate.航站 == '航站1'), date_1] = 1
                time_calculate.loc[(time_calculate.航班号 == b) & (time_calculate.航站 == '航站2'), date_1] = 1
            elif len(time_calculate.loc[(time_calculate.航班号 == b), "航站名称"]) % 3 == 0:
                cell1 = time_calculate.loc[(time_calculate.航班号 == b), "航站名称"].tolist()[0]
                cell2 = time_calculate.loc[(time_calculate.航班号 == b), "航站名称"].tolist()[1]
                cell3 = time_calculate.loc[(time_calculate.航班号 == b), "航站名称"].tolist()[2]
                if cell3 is not None:
                    if foc_value[date_1][b][0] == 2:
                        time_calculate.loc[(time_calculate.航班号 == b) & (time_calculate.航站 == '航站1'), date_1] = 1
                        time_calculate.loc[(time_calculate.航班号 == b) & (time_calculate.航站 == '航站2'), date_1] = 1
                        time_calculate.loc[(time_calculate.航班号 == b) & (time_calculate.航站 == '航站3'), date_1] = 1
                    else:
                        if foc_value[date_1][b].count(cell1) == 0:  # 如果第一段是没有值的，则取前段
                            time_calculate.loc[
                                (time_calculate.航班号 == b) & (time_calculate.航站 == '航站2'), date_1] = 0.5
                            time_calculate.loc[
                                (time_calculate.航班号 == b) & (time_calculate.航站 == '航站3'), date_1] = 1
                        if foc_value[date_1][b].count(cell2) == 0:  # 如果第二段是没有值的，则拉直
                            time_calculate.loc[
                                (time_calculate.航班号 == b) & (time_calculate.航站 == '航站1'), date_1] = 1
                            time_calculate.loc[
                                (time_calculate.航班号 == b) & (time_calculate.航站 == '航站3'), date_1] = 1
                        if foc_value[date_1][b].count(cell3) == 0:  # 如果第三段是没有值的，则取后
                            time_calculate.loc[
                                (time_calculate.航班号 == b) & (time_calculate.航站 == '航站2'), date_1] = 0.5
                            time_calculate.loc[
                                (time_calculate.航班号 == b) & (time_calculate.航站 == '航站1'), date_1] = 1
    time_calculate_list = np.array(time_calculate.iloc[1:, 15:]).tolist()
    worksheet1.range('P3').value = time_calculate_list
    workbook.save()

def foc_month_main(month, date_1, dict):
    month = str(month) + "月"
    worksheet_month = workbook.sheets[month]
    flight_list_month = worksheet_month.range('A3').expand('down').value  # 月度航班号表
    time_list_month = workbook.sheets[month].range('A1').expand('table').value[0][4:-2] # 月度时间表
    # 列表更替，将x更换到每一个空格里，记住，长度可能有问题，len(flight_list_month)，这个是要更新的
    X_list= ['X' if i is not None else i for i in worksheet_month.range((3, time_list_month.index(date_1)+5),(len(flight_list_month)-6, time_list_month.index(date_1)+5)).value]
    #整个竖着的列全部替换成x
    worksheet_month.range((3, time_list_month.index(date_1)+5)).options(transpose=True).value = X_list
    # 依次豁免
    for b in dict.keys():
        if b not in flight_list_month:
            print(b)
            continue
        cell = worksheet_month.range('D2').offset(column_offset=time_list_month.index(date_1) + 1,  # 右偏移n+1个格子，下偏移n+1
                                                  row_offset=flight_list_month.index(b) + 1)  # 锁
        cell3 = worksheet_month.range('D2').offset(column_offset=-2,
                                              row_offset=flight_list_month.index(b) + 1).value

        if cell3.count('-') == 1:  # 意味着这个是点对点航班，没有第三段
            cell.value = [1]
        else:
            if dict[b][0] == 2:
                cell.value = [1]


def part2(month, date_1, dict):
    month = str(month) + "月"
    maintable = pd.read_excel('/Users/ze/Desktop/23S执行率计算表.xlsx', sheet_name=month)
    maintable[date_1] = maintable[date_1].map({1: "X", "X": "X"})  # 替换每一列中的值为0
    flight_zhengban_list = maintable['航班号'].tolist()[1:-8]
    for b in dict.keys():
        if b not in flight_zhengban_list:
            print(b)
            continue
        if maintable.loc[(maintable.航班号 == b), "航线"].str.count("-").tolist()[0] == 1:
            maintable.loc[(maintable.航班号 == b), date_1] = 1
        elif maintable.loc[(maintable.航班号 == b), "航线"].str.count("-").tolist()[0] == 2:
            if dict[b][0] == 2:
                maintable.loc[(maintable.航班号 == b), date_1] = 1
    maintable_list = np.array(maintable[date_1]).tolist()[1:]
    workbook.sheets[month].range((3, maintable.columns.get_loc(date_1)+1)).options(transpose=True).value = maintable_list


def part1():
    time_calculate = pd.read_excel('/Users/ze/Desktop/23S执行率计算表.xlsx', sheet_name=1)
    maintable = pd.read_excel('/Users/ze/Desktop/23S执行率计算表.xlsx', sheet_name=2)
    # data1 = maintable.iloc[9:12, 9:44]
    citylist = maintable.iloc[9, 9:46].tolist()
    startday = time_calculate.columns.get_loc(datetime(2023, 4, 16))
    endday = time_calculate.columns.get_loc(datetime(2023, 6, 10)) + 1
    nowday = time_calculate.columns.get_loc(datetime(datetime.now().year, datetime.now().month, datetime.now().day)) + 1
    value_list = [[], []]
    for index, city in enumerate(citylist):
        if index >= 1:
            time_calculate.update(time_calculate.loc[(time_calculate.航站名称 == city)].loc[(time_calculate.航线.str.count('-') == 2) & (time_calculate.航站 == '航站2')].iloc[:,15:].apply(lambda x: x*2))
            slot_use_airport = time_calculate.loc[(time_calculate.航站名称 == city)].iloc[:, startday:endday].sum().sum()
            slot_time = time_calculate.loc[(time_calculate.航站名称 == city)].iloc[:, startday:endday].count().sum()+time_calculate.loc[(time_calculate.航站名称 == city)].loc[(time_calculate.航线.str.count('-') == 2) & (time_calculate.航站 == '航站2')].iloc[:, startday:endday].count().sum()
            # schedule_slot_time = (np.array(maintable.iloc[8, 10:46]) * 8).tolist()
            schedule_slot_time = (np.array(maintable.iloc[8, index+9]) * 8).tolist()
            city_percent = slot_use_airport/schedule_slot_time
            nowday_slot_use_airport = time_calculate.loc[(time_calculate.航站名称 == city)].iloc[:,
                               startday:nowday].sum().sum()
            nowday_slot_time = time_calculate.loc[(time_calculate.航站名称 == city)].iloc[:, startday:nowday].count().sum() + \
                        time_calculate.loc[(time_calculate.航站名称 == city)].loc[
                            (time_calculate.航线.str.count('-') == 2) & (time_calculate.航站 == '航站2')].iloc[:,
                        startday:nowday].count().sum()
            nowday_slot_time = time_calculate.loc[(time_calculate.航站2) == city].count().sum()
            nowday_city_percent = nowday_slot_use_airport / schedule_slot_time
            value_list[0].append(city_percent)
            value_list[1].append(nowday_city_percent)
    workbook.sheets['总汇'].range('K12').value = value_list
    workbook.save()

main()
workbook.close()
app.quit()
