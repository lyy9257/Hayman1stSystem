import CallBuyAndSell
import WatchPrice

import time
import datetime
import win32com.client
import xlrd

def TimeandPrice():
        now = time.localtime()
        Time = "%04d-%02d-%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
       
        print (Time)
        print (WatchPrice.NowPrice(1))
        print (' ')

def HowMuchBuy(NowPrice,Flag):
    MyDB = xlrd.open_workbook('MyDB.xlsx')
    worksheet = MyDB.sheet_by_index(0)
    k = 5
    j = 0
    
    while(k < 9):
        if(NowPrice >= int(worksheet.cell_value(20,k))):
            j += 1    
            k += 1
        else:
            k += 1
    
    if(Flag == 1):
        z = j * 0.25
        return z
    else:
        print("금일 종가 : " + str(NowPrice))
        z = j * 0.25
        print("매수 비중 : " + str(z))

def savetosell(StockAmount,flag):
    if(flag == 0):
        f = open("/StockAmount.txt", 'w')
        f.write(str(StockAmount))
    elif(flag == 1):
        f = open("/StockAmount.txt", 'r')
        k = f.readline()
        return k
    else:
        print("Wrong Flag!")

if __name__ == "__main__":

    while True:
        now = time.localtime()
        
        if now.tm_hour == 9 and now.tm_min == 0 and now.tm_sec == 3:
            try:
                Ratio = HowMuchBuy(WatchPrice.NowPrice(0),1)
                SellStockAmount = savetosell(CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1),1)
                print('매도수량 :',SellStockAmount)

                CallBuyAndSell.Sell('A233740', SellStockAmount)
            except:
                print("매수를 실행하지 않았습니다. 매도도 실행하지 않습니다.\n")

        elif now.tm_hour == 15 and now.tm_min == 19 and now.tm_sec == 20:
            WatchPrice.MakeExcelFile()

        elif now.tm_hour == 15 and now.tm_min == 20 and now.tm_sec == 3:
            HowMuchBuy(WatchPrice.NowPrice(0),2)
            Ratio = HowMuchBuy(WatchPrice.NowPrice(0),1)
            CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),0)        
            
            CallBuyAndSell.Buy('A233740', CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1))
            
            savetosell(CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1),0)
            print('매수수량 :',savetosell(CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1),1))
        
        elif now.tm_hour == 15 and now.tm_min == 21 and now.tm_sec == 3:
            print("프로그램을 종료합니다.")
            break;

        else :
            TimeandPrice()

        time.sleep(1)
