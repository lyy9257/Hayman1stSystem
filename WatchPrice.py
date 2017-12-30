#시세 Database를 뽑는 모듈
import win32com.client
import numpy as np
import pandas as pd
from pandas import Series, DataFrame
import xlrd

# 코드,종목명,현재가를 구하는 함수
def NowPrice(Flag):

    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, 'A233740')  # 종목 코드 - KODEX 코스닥 150 레버리지
    objStockMst.BlockRequest()
 
    # 현재가 정보 조회
    code = objStockMst.GetHeaderValue(0)  # 종목코드
    name = objStockMst.GetHeaderValue(1)  # 종목명
    time = objStockMst.GetHeaderValue(4)  # 시간
    CurrentPrice = objStockMst.GetHeaderValue(11)  # 종가

    # 코드,이름,현재가 리턴
    if Flag == 1:
        return code, name, CurrentPrice
    
    # 현재가 리턴
    else:
        return CurrentPrice

# DataBase 제작 함수
def MakeDataBase(k):

    # 사이보스 플러스를 이용하여 데이터를 끌어온다.
    data = []
    objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
 
    objStockChart.SetInputValue(0, 'A233740') # 종목 코드 - KODEX 코스닥 150 레버리지
    objStockChart.SetInputValue(1, '2') # 개수로 조회
    objStockChart.SetInputValue(4, 20) # 최근 20일 치
    objStockChart.SetInputValue(5, [0,2,3,4,5,8]) # 날짜,시가,고가,저가,종가,거래량
    objStockChart.SetInputValue(6, ord('D')) # 차트 주가 - 일간 차트 요청
    objStockChart.SetInputValue(9, '1') # 수정주가 사용
    objStockChart.BlockRequest()
 
    len = objStockChart.GetHeaderValue(3)
    j = 0

    # 끌어온 데이터의 가공
    for i in range(len):
        day = objStockChart.GetDataValue(0, j)
        open = objStockChart.GetDataValue(1, j)
        high = objStockChart.GetDataValue(2, j)
        low = objStockChart.GetDataValue(3, j)
        close = objStockChart.GetDataValue(4, j)
        
        data.append([day, open, high, low, close])
        j = j + 1

    data.reverse()

    if k == 0:
        return len
    else:
        return data

def MakeExcelFile():

    # 연결 여부 체크
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    
    # 데이터프레임 제작
    data1 = MakeDataBase(1)
    df_1st = DataFrame(data1,  columns=['날짜', '시가', '고가', '저가', '종가'])
    df_1st.sort_index(0,ascending=False)

    # 기존 데이터에 MA를 추가시켜준다.
    MovingAverage_3Days = round(df_1st['종가'].rolling(window = 3).mean())
    df_1st.insert(len(df_1st.columns), "3일이평", MovingAverage_3Days)

    MovingAverage_5Days = round(df_1st['종가'].rolling(window = 5).mean())
    df_1st.insert(len(df_1st.columns), "5일이평", MovingAverage_5Days)

    MovingAverage_10Days = round(df_1st['종가'].rolling(window = 10).mean())
    df_1st.insert(len(df_1st.columns), "10일이평", MovingAverage_10Days)

    MovingAverage_20Days = round(df_1st['종가'].rolling(window = 20).mean())
    df_1st.insert(len(df_1st.columns), "20일이평", MovingAverage_20Days)
    
    # 엑셀화
    df_1st.to_excel('./MyDB.xlsx',encoding = 'euc_KR')
    print('DB 제작 완료!')