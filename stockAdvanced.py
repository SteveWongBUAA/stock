#!/usr/bin/python
#encoding:utf-8
import xlwt
import xlrd
import datetime
import numpy as np
import matplotlib.pyplot as plt
#读文件
name = '159929'
title = xlrd.open_workbook(name + '.xls')
table = title.sheet_by_name('Table')
print 'fund: ', name
#写文件
wb = xlwt.Workbook(encoding='utf-8')
sheet1 = wb.add_sheet('result')
#读取每一列
price = table.col_values(1)
date = table.col_values(0)
limit = table.col_values(2)
flag = []
#第一步：判断当天收盘价是否高于20日线
sheet1.write(20, 0, 'date')
sheet1.write(20, 1, 'price')
sheet1.write(20, 2, 'limit')
sheet1.write(20, 3, 'MA20')
sheet1.write(20, 4, 'price>MA20?')

sheet1.write(20, 5, 'stockNum')
sheet1.write(20, 6, 'stockMoney')
sheet1.write(20, 7, 'restMoney')
sheet1.write(20, 8, 'savedMoney')
sheet1.write(20, 9, 'totalMoney')
sheet1.write(20, 10, 'ZRBM')
sheet1.write(20, 11, 'stockNum1')
sheet1.write(20, 12, 'stockMoney1')
sheet1.write(20, 13, 'restMoney1')
sheet1.write(20, 14, 'totalMoney1')

#初始化flag为-1
for i in range(0, len(price)):
    flag.append(-1)

dateShow = []
priceShow = []
buyShow = []
sellShow = []
MA20Show = []
for i in range(21, len(price)):
    #每20个price组成一个temp数组
    tmplist = price[i-20:i]
    #print tmplist
    sheet1.write(i, 0, datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').strftime('%Y-%m-%d'))
    sheet1.write(i, 1, price[i])
    sheet1.write(i, 2, limit[i])
    dateShow.append(datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').date())
    priceShow.append(price[i])
    buyShow.append(0)
    sellShow.append(0)
    #print tmplist
    MA20 = sum(tmplist)/20.0
    MA20Show.append(MA20)
    if(MA20 > price[i]):
        flag[i] = 0
    else:
        flag[i] = 1
    sheet1.write(i, 3, MA20)
    sheet1.write(i, 4, flag[i])

#基础设定
inputMoney = 10000
print 'inputMoney', inputMoney
monthIn = 1000
fee = 0.001#交易手续费
ZRB = 0.1/12#真融宝收益比较
ZRBM = inputMoney
#开始投资
tmpMonth = date[21][5:7]#初始月份
#初始状态
if flag[21] == 1:#初始状态为持有
    stockNum = int(inputMoney/(1+fee) / price[21] / 100) * 100
    stockMoney = stockNum * price[21]
    restMoney = inputMoney - stockMoney - stockMoney * fee
    savedMoney = 0
    buyShow[0] = price[21]
else:#初始状态为不持有
    stockNum = 0
    stockMoney = 0
    restMoney = inputMoney
    savedMoney = 0
    buyShow[0] = 0
totalMoney = stockMoney + restMoney + savedMoney
cost = inputMoney
sheet1.write(21, 5, stockNum)
sheet1.write(21, 6, stockMoney)
sheet1.write(21, 7, restMoney)
sheet1.write(21, 8, savedMoney)
sheet1.write(21, 9, totalMoney)

#无脑投初始状态
stockNum1 = int(inputMoney/(1+fee) / price[21] / 100) * 100
stockMoney1 = stockNum1 * price[21]
restMoney1 = inputMoney - stockMoney1 - stockMoney1 * fee
totalMoney1 = stockMoney1 + restMoney1

#开始循环投资
for i in range(22, len(flag)-1):
    #先算今天的资产情况
    stockMoney = stockNum * price[i]
    stockMoney1 = stockNum1 * price[i]
    #计算当前月份
    nowMonth = date[i][5:7]
    if nowMonth != tmpMonth:#月份发生改变,定投
        #二十日线定投
        cost += monthIn
        savedMoney += monthIn
        tmpMonth = nowMonth
        #真融宝
        ZRBM = ZRBM * (1+ZRB) + monthIn
        sheet1.write(i, 10, ZRBM)
        #无脑定投
        tmpStockNum1 = int((monthIn + restMoney1)/(1+fee) / price[i] / 100) * 100
        if tmpStockNum1 != 0:#够买一份了，那就买买买
            stockNum1 += tmpStockNum1
            stockMoney1 += (tmpStockNum1 * price[i])
            restMoney1 = restMoney1 + monthIn - tmpStockNum1 * price[i] - tmpStockNum1 * price[i] * fee
            buyShow[i-21] = price[i]
        sheet1.write(i, 11, stockNum1)
        sheet1.write(i, 12, stockMoney1)
        sheet1.write(i, 13, restMoney1)
        totalMoney1 = stockMoney1 + restMoney1
        sheet1.write(i, 14, totalMoney1)
    if flag[i] == 1:#如果持有标志为1则再把剩余的钱买入
        tmpStockNum = int((savedMoney + restMoney)/(1+fee) / price[i] / 100) * 100
        if tmpStockNum != 0:#购买一份了，那就买买买
            stockNum += tmpStockNum
            stockMoney += (tmpStockNum * price[i])
            restMoney = restMoney + savedMoney - tmpStockNum * price[i] - tmpStockNum * price[i] * fee
            savedMoney = 0#存起来的钱清零
            buyShow[i-21] = price[i]
    else:#如果持有标志为0, 卖掉
        restMoney += (stockMoney * (1-fee))
        stockMoney = 0
        stockNum = 0
        sellShow[i-21] = price[i]
    sheet1.write(i, 5, stockNum)
    sheet1.write(i, 6, stockMoney)
    sheet1.write(i, 7, restMoney)
    sheet1.write(i, 8, savedMoney)
    totalMoney = stockMoney + restMoney + savedMoney
    sheet1.write(i, 9, totalMoney)

daysPast = datetime.datetime.strptime(date[-1][0:10], '%Y-%m-%d') - \
                    datetime.datetime.strptime(date[1][0:10], '%Y-%m-%d')
out = "从%s到%s，购买%s的天数为%s天(共%s年)，初始资金%d，每月定投%d，最后总投入为%d。同期定投真融宝(年化收益0.1)得到%d。" \
      % (str(date[1][0:10]), str(date[-1][0:10]), name, str(daysPast), str(daysPast.days/365.0), inputMoney, monthIn, cost, ZRBM)
out0 = "20日线定投基金最终得到%d。" % totalMoney
out1 = "无脑月初定投%d，最终得到%d。" \
      % (monthIn, totalMoney1)
out00 = "交易规则：初始资金10000，每月定投1000，每天检查，如果当天价格在二十日线以上，再判断剩余的钱是否足够买100股以上，足够的话，买买买。当天价格在二十日线以下就全仓卖掉"
sheet1.write(0, 0, out)
sheet1.write(1, 0, out0)
sheet1.write(2, 0, out1)
sheet1.write(3, 0, out00)

print out.decode('utf-8')
print out0.decode('utf-8')
print out1.decode('utf-8')

fig = plt.figure()
plt.title(name)
plt.xlabel('time')
plt.ylabel('price')
plt.grid(True)
plt.ylim(min(priceShow), max(priceShow))
x = np.linspace(0, 1, len(priceShow))
plt.plot_date(dateShow, priceShow, 'b-', label='Daily Closing Price')
plt.plot_date(dateShow, MA20Show, 'k--', label='MA20', linewidth=2)
plt.scatter(dateShow, buyShow, s=80, label='Buy', c='red', marker='*')
#plt.plot_date(dateShow, buyShow)
plt.scatter(dateShow, sellShow, s=100, label='Sell', c='yellow', marker='.')
#plt.plot_date(dateShow, sellShow)
plt.legend(numpoints=1, fontsize=18)
#plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
plt.legend(loc='upper left')
plt.show()
wb.save(name + '_' + str(datetime.datetime.now().strftime('%d_%H')) +'.xls')