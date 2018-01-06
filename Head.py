# 메인 모듈
import CallBuyAndSell
import WatchPrice

import time
import datetime
import win32com.client
import xlrd

# 시간과 현재가격을 출력하는 함수
def TimeandPrice():

        # 시간 호출
        now = time.localtime()
        Time = "%04d-%02d-%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
        
        # 시간과 가격 출력
        print (Time)
        print (WatchPrice.NowPrice(1))
        print (' ')

# 매수 비중을 결정하는 함수
def HowMuchBuy(NowPrice,Flag):
    
    # DB 파일을 불러옴
    MyDB = xlrd.open_workbook('MyDB.xlsx')
    worksheet = MyDB.sheet_by_index(0)
    k = 5
    j = 0
    
    # MA와 현재가를 비교
    while(k < 9):

        # MA보다 현재가가 크면 비중 + 1
        if(NowPrice >= int(worksheet.cell_value(20,k))):
            j += 1    
            k += 1

        # MA보다 현재가가 작으면 비중 그대로.
        else:
            k += 1
    
    # 백분위로 환산
    if(Flag == 1):
        z = j * 0.25
        return z

    # 백분위로 환산하되, 사용자에게 보여주기 위해 Print 실시.
    else:
        print("금일 종가 : " + str(NowPrice))
        z = j * 0.25
        print("매수 비중 : " + str(z))

# 매매수량 기록 및 호출함수
def savetosell(StockAmount,flag):

    # 매수 후 매수수량 입력
    if(flag == 0):
        f = open("/StockAmount.txt", 'w')
        f.write(str(StockAmount))
   
    # 매도 때 매수한 수량 로딩
    elif(flag == 1):
        f = open("/StockAmount.txt", 'r')
        k = f.readline()
        return k

    else:
        print("Wrong Flag!")

# Main 함수
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

        elif now.tm_hour == 15 and now.tm_min == 21 and now.tm_sec == 30:
            WatchPrice.MakeExcelFile()

        elif now.tm_hour == 15 and now.tm_min == 23 and now.tm_sec == 3:
            HowMuchBuy(WatchPrice.NowPrice(0),2)
            Ratio = HowMuchBuy(WatchPrice.NowPrice(0),1)
            CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),0)        
            
            CallBuyAndSell.Buy('A233740', CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1))
            
            savetosell(CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1),0)
            print('매수수량 :',savetosell(CallBuyAndSell.ChoiseToAmountToBuy(Ratio,WatchPrice.NowPrice(0),1),1))
        
        elif now.tm_hour == 15 and now.tm_min == 24 and now.tm_sec == 3:
            print("프로그램을 종료합니다.")
            break;

        else :
            TimeandPrice()

        time.sleep(1)
     