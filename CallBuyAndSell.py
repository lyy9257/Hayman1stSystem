# 조건에 따른 매도 매수를 시행하는 모듈
import win32com.client

# 사이보스, 크레온 플러스의 연결 여부를 확인하는 함수
def CybosConnection():
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()

# 얼마나 살지 수량을 정하는 함수
def ChoiseToAmountToBuy(AccountRatio,NowPrice,flag):
    
    # 주문 초기화
    objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
    objTrade.TradeInit(0)
    AccountNumber = objTrade.AccountNumber[0] # 계좌번호
    objTrade5331A = win32com.client.Dispatch("CpTrade.CpTdNew5331A")

    objTrade5331A.SetInputValue(0, AccountNumber)
    objTrade5331A.SetInputValue(1, "01")
    objTrade5331A.SetInputValue(2, 'A233740')
    objTrade5331A.SetInputValue(3, '03')
    objTrade5331A.SetInputValue(4, NowPrice)
    objTrade5331A.SetInputValue(6, 2)
    objTrade5331A.BlockRequest()

    AbleToBuy = objTrade5331A.GetHeaderValue(17) # 현금 주문 가능수량
    Amount = objTrade5331A.GetHeaderValue(45) # 잔고 호출

    StockAmount = int(AbleToBuy * AccountRatio)
    
   # Flag 변수를 이용하여 출력값을 선택한다.
    if(flag == 0):
        print('계좌 :', AccountNumber)
        print('현재 잔고 :', Amount)

        print('주문 가능 수량 :', AbleToBuy)
        print('주문 수량 :', StockAmount)

    elif(flag == 1):
        return StockAmount

    else:
        print('잘못된 Flag입니다!')

# 매수 함수
def Buy(Code, HowMuchBuy):  

    # 주문 초기화
    objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
 
    # 주식 매수 주문
    AccountNumber = objTrade.AccountNumber[0] # 계좌번호
    accFlag = objTrade.GoodsList(AccountNumber, 1) # 주식상품 구분
    print(AccountNumber, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "2") # 2 : 매수
    objStockOrder.SetInputValue(1, AccountNumber) #  계좌번호
    objStockOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, Code) # 종목코드 - 입력받은 코드
    objStockOrder.SetInputValue(4, int(HowMuchBuy)) # 매수수량 : 위에서 정한 수량
    objStockOrder.SetInputValue(5, 14100) # 주문단가  - 14,100(보통가 주문에서만 유효)
    objStockOrder.SetInputValue(7, "0")
    objStockOrder.SetInputValue(8, "03") # 03 시장가로 주문 
 
    # 매수 주문 요청
    objStockOrder.BlockRequest()
 
    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()

def Sell(Code, HowMuchBuy):

    # 주문 초기화
    objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
    if (initCheck != 0):
        print("주문 초기화 실패")
        exit()

    # 주식 매도 주문
    acc = objTrade.AccountNumber[0] # 계좌번호
    accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "1") # 1 : 매도
    objStockOrder.SetInputValue(1, acc ) # 계좌번호
    objStockOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, Code) # 종목코드 - 입력받은 코드
    objStockOrder.SetInputValue(4, HowMuchBuy) # 매도수량 : 매수한 수량만큼
    objStockOrder.SetInputValue(5, 14100) # 주문단가 - 14,100원 (시장가에서는 필요없음)
    objStockOrder.SetInputValue(7, "0")
    objStockOrder.SetInputValue(8, "03") # 시장가
 
    # 매도 주문 요청
    objStockOrder.BlockRequest()
 
    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()