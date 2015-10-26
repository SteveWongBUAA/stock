#!/usr/bin/python
#encoding:utf-8
import xlwt
import xlrd
import datetime
import matplotlib.pyplot as plt

def norm(list, gain=1.0):
    #normalize totalMoneyShow
    listMax = max(list)
    listMin = min(list)
    times = float(listMax - listMin)
    listNorm = []
    for i in list:
        listNorm.append(gain * (i - listMin) / times)
    return listNorm

def normWithMaxNMin(list, max, min, gain=1.0):
    #normalize totalMoneyShow
    times = float(max - min)
    listNorm = []
    for i in list:
        listNorm.append(gain * (i - min) / times)
    return listNorm


#############################################################
#基础设定
name = '159902'
inputMoney = 10000
monthIn = 1000
fee = 0.001#交易手续费
ZRB = 1.1**(1.0/12)-1#真融宝收益比较
ZRBM = inputMoney
timesM = 10000.0
timesPrice = 1.0
gainNorm = 0.8
myDPI = 120
length = 40
width = 20
def zoomPrice(data, timesPrice=1.0):
    ret = []
    for i0 in data:
        ret.append(i0 * timesPrice)
    return ret
#############################################################
#读文件
title = xlrd.open_workbook(name + '.xls')
table = title.sheet_by_name('Table')
#写文件
wb = xlwt.Workbook(encoding='utf-8')
sheet1 = wb.add_sheet('result')
#读取每一列
tmpPrice = table.col_values(1)
date = table.col_values(0)
limit = table.col_values(2)
price = [0]
if tmpPrice[1] > 10:#股价太高，说明是指数，降低1000倍
    for i in range(1, len(tmpPrice)):
        price.append(tmpPrice[i]/1000.0)
else:
    for i in range(1, len(tmpPrice)):
        price.append(tmpPrice[i])
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
buyDateShow = []
sellShow = []
sellDateShow = []
MA20Show = []
totalMoneyShow = []
costShow = []
totalMoneyStupidShow = []
totalMoneyZRBShow = []

for i in range(21, len(price)):
    #每20个price组成一个temp数组
    tmplist = price[i-20:i]
    #print tmplist
    sheet1.write(i, 0, datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').strftime('%Y-%m-%d'))
    sheet1.write(i, 1, price[i])
    sheet1.write(i, 2, limit[i])
    dateShow.append(datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').date())
    priceShow.append(price[i])
    # buyShow.append(0)
    # sellShow.append(0)
    #print tmplist
    MA20 = sum(tmplist)/20.0
    MA20Show.append(MA20)
    if(MA20 > price[i]):
        flag[i] = 0
    else:
        flag[i] = 1
    sheet1.write(i, 3, MA20)
    sheet1.write(i, 4, flag[i])




#开始投资
tmpMonth = date[21][5:7]#初始月份
#初始状态
if flag[21] == 1:#初始状态为持有
    stockNum = int(inputMoney/(1+fee) / price[21] / 100) * 100
    stockMoney = stockNum * price[21]
    restMoney = inputMoney - stockMoney - stockMoney * fee
    savedMoney = 0
    buyShow.append(priceShow[0])
    buyDateShow.append(dateShow[0])
else:#初始状态为不持有
    stockNum = 0
    stockMoney = 0
    restMoney = inputMoney
    savedMoney = 0
    sellShow.append(priceShow[0])
    sellDateShow.append(dateShow[0])
totalMoney = stockMoney + restMoney + savedMoney
cost = inputMoney
sheet1.write(21, 5, stockNum)
sheet1.write(21, 6, stockMoney)
sheet1.write(21, 7, restMoney)
sheet1.write(21, 8, savedMoney)
sheet1.write(21, 9, totalMoney)
costShow.append(cost/timesM)
totalMoneyShow.append(totalMoney/timesM)


#无脑投初始状态
stockNum1 = int(inputMoney/(1+fee) / price[21] / 100) * 100
stockMoney1 = stockNum1 * price[21]
restMoney1 = inputMoney - stockMoney1 - stockMoney1 * fee
totalMoney1 = stockMoney1 + restMoney1
totalMoneyStupidShow.append(totalMoney1/timesM)

totalMoneyZRBShow.append(ZRBM/timesM)

#开始循环投资
for i in range(22, len(flag)):
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
            buyShow.append(price[i])
            buyDateShow.append(datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').date())
        else:
            restMoney1 += monthIn
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
            buyShow.append(price[i])
            buyDateShow.append(datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').date())
    else:#如果持有标志为0, 卖掉
        restMoney += (stockMoney * (1-fee))
        stockMoney = 0
        stockNum = 0
        sellShow.append(price[i])
        sellDateShow.append(datetime.datetime.strptime(date[i][0:10], '%Y-%m-%d').date())
    sheet1.write(i, 5, stockNum)
    sheet1.write(i, 6, stockMoney)
    sheet1.write(i, 7, restMoney)
    sheet1.write(i, 8, savedMoney)
    totalMoney = stockMoney + restMoney + savedMoney
    sheet1.write(i, 9, totalMoney)
    costShow.append(cost/timesM)
    totalMoneyShow.append(totalMoney/timesM)
    totalMoneyStupidShow.append(totalMoney1/timesM)
    totalMoneyZRBShow.append(ZRBM/timesM)



daysPast = datetime.datetime.strptime(date[-1][0:10], '%Y-%m-%d') - \
                    datetime.datetime.strptime(date[1][0:10], '%Y-%m-%d')

AnnualInterestRate = ((totalMoney / float(cost)) ** (1 / (float(daysPast.days) / 365.0)) - 1) * 100

out = "从%s到%s，购买%s的天数为%s天(共%s年)，初始资金%d，每月定投%d，最后总投入为%d。同期定投真融宝(年化收益0.1)得到%d。" \
      % (str(date[1][0:10]), str(date[-1][0:10]), name, str(daysPast), str(daysPast.days/365.0), inputMoney, monthIn, cost, ZRBM)
out0 = "20日线定投基金最终得到%d。年化收益率%f" % (totalMoney, AnnualInterestRate) + '%'
out1 = "无脑月初定投%d，最终得到%d。" \
      % (monthIn, totalMoney1)
out00 = "交易规则：初始资金%d，每月定投%d，每天检查，如果当天价格在二十日线以上，\
再判断剩余的钱是否足够买100股以上，足够的话，买买买。当天价格在二十日线以下就全仓卖掉" % (inputMoney, monthIn)
sheet1.write(0, 0, out)
sheet1.write(1, 0, out0)
sheet1.write(2, 0, out1)
sheet1.write(3, 0, out00)

for i in range(1, len(totalMoneyShow)):
    tr = totalMoneyShow[i] / totalMoneyShow[i-1]
    if tr < 0.95:
        print 'Big Fall!'
        print dateShow[i], tr

print out.decode('utf-8')
print out0.decode('utf-8')
print out1.decode('utf-8')

fig = plt.figure(figsize=(length, width), dpi=myDPI)
plt.title(name)
plt.xlabel('time')
plt.ylabel('data')
plt.grid(True)
upLimit = max([max(priceShow), max(totalMoneyShow), max(costShow)])
downLimit = min([min(priceShow), min(totalMoneyShow), min(costShow)])
plt.ylim(downLimit, upLimit)
plt.plot_date(dateShow, zoomPrice(priceShow, timesPrice), 'b-', label='Daily Closing Price')
plt.plot_date(dateShow, zoomPrice(MA20Show, timesPrice), 'k--', label='MA20', linewidth=2)
plt.scatter(buyDateShow, zoomPrice(buyShow, timesPrice), s=80, label='Buy', c='red', marker='*')
plt.scatter(sellDateShow, zoomPrice(sellShow, timesPrice), s=100, label='Sell', c='yellow', marker='.')
plt.plot_date(dateShow, costShow, 'c--', label='Cost(w)', linewidth=1)
plt.plot_date(dateShow, totalMoneyShow, 'm-', label='TotalMoney(w)', linewidth=2)
plt.plot_date(dateShow, totalMoneyStupidShow, 'g:', label='TotalMoneyStupid(w)', linewidth=2)
plt.plot_date(dateShow, totalMoneyZRBShow, 'k--', label='TotalMoneyZRB(w)', linewidth=1)
#plt.plot_date(dateShow, buyShow)

#plt.plot_date(dateShow, sellShow)
plt.legend(numpoints=1, fontsize=18)
#plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
plt.legend(loc='upper left')
#plt.show()
plt.savefig(name + '.png')

fig1 = plt.figure(figsize=(length, width), dpi=myDPI)
plt.title(name + '(Norm)(price*%.1f)' % round(gainNorm,1))
plt.xlabel('time')
plt.ylabel('data')
plt.grid(True)
priceShowNorm = norm(priceShow, gainNorm)
MA20ShowNorm = norm(MA20Show, gainNorm)
buyShowNorm = normWithMaxNMin(buyShow, max(priceShow), min(priceShow), gainNorm)
sellShowNorm = normWithMaxNMin(sellShow, max(priceShow), min(priceShow), gainNorm)
totalMoneyShowNorm = norm(totalMoneyShow)
costShowNorm = normWithMaxNMin(costShow, max(totalMoneyShow), min(totalMoneyShow))
totalMoneyStupidShowNorm = normWithMaxNMin(totalMoneyStupidShow, max(totalMoneyShow), min(totalMoneyShow))
totalMoneyZRBShowNorm = normWithMaxNMin(totalMoneyZRBShow, max(totalMoneyShow), min(totalMoneyShow))
# plt.ylim(downLimit, upLimit)
plt.plot_date(dateShow, priceShowNorm, 'b-', label='Daily Closing Price')
plt.plot_date(dateShow, MA20ShowNorm, 'k--', label='MA20', linewidth=2)
plt.scatter(buyDateShow, buyShowNorm, s=80, label='Buy', c='red', marker='*')
plt.plot_date(dateShow, costShowNorm, 'c--', label='Cost', linewidth=2)
plt.plot_date(dateShow, totalMoneyShowNorm, 'm-', label='TotalMoney', linewidth=1)
plt.plot_date(dateShow, totalMoneyStupidShowNorm, 'g:', label='TotalMoneyStupid(w)', linewidth=2)
plt.plot_date(dateShow, totalMoneyZRBShowNorm, 'k--', label='TotalMoneyZRB(w)', linewidth=1)
plt.scatter(sellDateShow, sellShowNorm, s=100, label='Sell', c='yellow', marker='.')
#plt.plot_date(dateShow, sellShow)
plt.legend(numpoints=1, fontsize=18)
#plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
plt.legend(loc='upper left')
plt.savefig(name + 'Norm_priceX%.1f.png'%gainNorm)
#plt.show()
wb.save(name + '_' + str(datetime.datetime.now().strftime('%d_%H')) + '.xls')