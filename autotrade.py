import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
from slacker import Slacker
import time, calendar
import requests

def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )
 
myToken = "token token"
def dbgout(message):
    """Printing the token and the stock information on both python shell and slack."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken,"#stock-info", strbuf)

def printlog(message, *args):
    """print the message logs on the python shell."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)
 
# CREAON API objects
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def check_creon_system():
    """Check if Creon has been initialized successfully"""
    # With admin privilege?
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False
 
    # Connected?
    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False
 
    # Is initialized (only to be used when account info is passed)
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True

def get_current_price(code):
    """return the current price, buy/sell price"""
    cpStock.SetInputValue(0, code)  # price info based on the code
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # current price
    item['ask'] =  cpStock.GetHeaderValue(16)        # buying price
    item['bid'] =  cpStock.GetHeaderValue(17)        # selling price    
    return item['cur_price'], item['ask'], item['bid']

def get_ohlc(code, qty):
    """return OHLC price info based on the quantity (qty)"""
    cpOhlc.SetInputValue(0, code)           # code
    cpOhlc.SetInputValue(1, ord('2'))        # 1:period, 2:qty
    cpOhlc.SetInputValue(4, qty)             # order qty
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0:date, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D:day
    cpOhlc.SetInputValue(9, ord('1'))        # 0:share value, 1: value after alteration
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count): 
        index.append(cpOhlc.GetDataValue(0, i)) 
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)]) 
    df = pd.DataFrame(rows, columns=columns, index=index) 
    return df

def get_stock_balance(code):
    """return the name of the stock and amount purchased"""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # account number
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:all, 1:stock, 2:options
    cpBalance.SetInputValue(0, acc)         # account number
    cpBalance.SetInputValue(1, accFlag[0])  # stock categories - first stock out of all
    cpBalance.SetInputValue(2, 50)          # number of requests (max 50)
    cpBalance.BlockRequest()     
    if code == 'ALL':
        dbgout('Account Name: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('Balance" : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('Total asset: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('Net gain: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('Number of Stocks: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # code
        stock_name = cpBalance.GetDataValue(0, i)   # name
        stock_qty = cpBalance.GetDataValue(15, i)   # quantity
        if code == 'ALL':
            dbgout(str(i+1) + ' ' + stock_code + '(' + stock_name + ')' 
                + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name, 
                'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0

def get_current_cash():
    """Amount that can be ordered with 100% balance"""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # account number
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:all, 1:stock, 2:options
    cpCash.SetInputValue(0, acc)              # account number
    cpCash.SetInputValue(1, accFlag[0])      # stock categories - first stock out of all
    cpCash.BlockRequest() 
    return cpCash.GetHeaderValue(9) # 100% balance oderable

def get_target_price(code):
    """Selling target price"""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open 
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]                                      
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        target_price = today_open + (lastday_high - lastday_low) * 0.5
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None
    
def get_movingaverage(code, window):
    """Moving average line"""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()         
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None    

def buy_etf(code):
    """Fill or Kill"""
    try:
        global bought_list      # global var for bought list
        if code in bought_list: # if bought then dont buy again
            #printlog('code:', code, 'in', bought_list)
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code) 
        target_price = get_target_price(code)    # buying target
        ma5_price = get_movingaverage(code, 5)   # 5 day moving avg
        ma10_price = get_movingaverage(code, 10) # 10 day moving avg
        buy_qty = 0        # reset quantity to purchase
        if ask_price > 0:  # if asking price or target price exists 
            buy_qty = buy_amount // ask_price  
        stock_name, stock_qty = get_stock_balance(code)  # search using code
        #printlog('bought_list:', bought_list, 'len(bought_list):',
        #    len(bought_list), 'target_buy_count:', target_buy_count)     
        if current_price > target_price and current_price > ma5_price \
            and current_price > ma10_price:  
            printlog(stock_name + '(' + str(code) + ') ' + str(buy_qty) +
                'EA : ' + str(current_price) + ' meets the buy condition!`')            
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]      # acconut number
            accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:all,1:stock,2:option                
            # FOK
            cpOrder.SetInputValue(0, "2")        # 2: buy
            cpOrder.SetInputValue(1, acc)        # account number
            cpOrder.SetInputValue(2, accFlag[0]) # category, 1st. 
            cpOrder.SetInputValue(3, code)       # code
            cpOrder.SetInputValue(4, buy_qty)    # buying qty
            cpOrder.SetInputValue(7, "2")        # order condition 0:basic, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")       # order price 1:normal, 3:market price
                                                 # 5:conditional, 12:IOC, 13:FOK 
            # Purchase order
            ret = cpOrder.BlockRequest() 
            printlog('FoK Purchase ->', stock_name, code, buy_qty, '->', ret)
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('request time limit; remaining time:', remain_time/1000)
                time.sleep(remain_time/1000) 
                return False
            time.sleep(2)
            printlog('balance :', buy_amount)
            stock_name, bought_qty = get_stock_balance(code)
            printlog('get_stock_balance :', stock_name, stock_qty)
            if bought_qty > 0:
                bought_list.append(code)
                dbgout("`buy_etf("+ str(stock_name) + ' : ' + str(code) + 
                    ") -> " + str(bought_qty) + "EA bought!" + "`")
    except Exception as ex:
        dbgout("`buy_etf("+ str(code) + ") -> exception! " + str(ex) + "`")

def sell_all():
    """Sell all shares with IOC"""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]       # Account Number
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:all, 1:stock, 2:options
        while True:    
            stocks = get_stock_balance('ALL') 
            total_qty = 0 
            for s in stocks:
                total_qty += s['qty'] 
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:                  
                    cpOrder.SetInputValue(0, "1")         # 1:sell, 2:buy
                    cpOrder.SetInputValue(1, acc)         # Account number
                    cpOrder.SetInputValue(2, accFlag[0])  # category, first
                    cpOrder.SetInputValue(3, s['code'])   # code
                    cpOrder.SetInputValue(4, s['qty'])    # sell qty
                    cpOrder.SetInputValue(7, "1")   # condition 0:normal, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # price 12:IOC, 13:FOK 
                    # Sell with IOC
                    ret = cpOrder.BlockRequest()
                    printlog('IOC Sell', s['code'], s['name'], s['qty'], 
                        '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('request time limit; remaining time:', remain_time/1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))

if __name__ == '__main__': 
    try:
        symbol_list = ['A122630', 'A252670', 'A233740', 'A250780', 'A225130',
             'A280940', 'A261220', 'A217770', 'A295000', 'A176950']
        bought_list = []     # list of bought
        target_buy_count = 5 # target of shares to buy
        buy_percent = 0.19   
        printlog('check_creon_system() :', check_creon_system())  # check connection
        stocks = get_stock_balance('ALL')      # view all owned stocks
        total_cash = int(get_current_cash())   # 100% balance check
        buy_amount = total_cash * buy_percent  # calculate amount per shaer
        printlog('100% balance :', total_cash)
        printlog('share % :', buy_percent)
        printlog('share amount :', buy_amount)
        printlog('start time:', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False;

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0,microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # End on Saturday or Sunday
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)
            if t_9 < t_now < t_start and soldout == False:
                soldout = True
                sell_all()
            if t_start < t_now < t_sell :  # AM 09:05 ~ PM 03:15 : Buy
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 5: 
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : Sell All
                if sell_all() == True:
                    dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :End program
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')
