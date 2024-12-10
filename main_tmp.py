import sys

# 추가할 경로
# additional_path = "C:/Users/boom1/PycharmProjects/pythonProject/WEB_SAA_TEST"
# sys.path.append(additional_path)
# additional_path = "C:/Users/boom1/PycharmProjects/pythonProject/webSAA"
# sys.path.append(additional_path)
# additional_path = "C:/Users/boom1/PycharmProjects/pythonProject/web_simulation"
# sys.path.append(additional_path)
# sys.path에 경로 추가


import pandas as pd
import resampled_mvo_tmp
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
import backtest_tmp

file = "C:/Users/이승기/Desktop/DATA/Python_project/RawData_.xlsx"


price = pd.read_excel(file, sheet_name="price", parse_dates=["Date"], index_col=0, header=0).dropna()

universe = pd.read_excel(file, sheet_name="universe",
                         names=None, dtype={'Date': datetime}, header=0)

universe['key'] = universe['symbol'] + " - " + universe['name']

assets = universe['symbol']

input_price = price[list(assets)]
input_universe = universe[universe['symbol'].isin(list(assets))].drop(['key'], axis=1)
input_universe = input_universe.reset_index(drop=True)  # index 깨지면 Optimization 배열 범위 초과 오류 발생

start_date = input_price.index[0]
start_date = datetime.combine(start_date, datetime.min.time())

end_date = input_price.index[-1]
end_date = datetime.combine(end_date, datetime.min.time())
constraint_range=[[0,100], [0,100], [0,100]]


daily = False
monthly = True
annualization = 12
freq = "monthly"
nPort = 100
nSim = 10

EF = resampled_mvo_tmp.simulation(input_price,
                              nSim, nPort,
                              input_universe,
                              constraint_range,
                                annualization, freq)

Target = 0.05

Opt_Weight = EF[abs(EF['EXP_RET']-Target) ==
                min(abs(EF['EXP_RET']-Target))].drop(columns=['EXP_RET','STDEV'])

Opt_Weight['Cash'] = 1- Opt_Weight.sum().sum()
#Opt_Weight = Opt_Weight.T.squeeze().tolist()

input_price = pd.concat([input_price.pct_change().dropna(),  pd.DataFrame({'Cash': [100] * len(input_price)}, index=input_price.index)], axis=1)


portfolio_port, allocation_f = backtest_tmp.simulation(input_price, Opt_Weight, 0, 'Monthly', 'Daily')
alloc = allocation_f.copy()
ret = (input_price.iloc[1:] / input_price.shift(1).dropna()) - 1
contribution = ((ret * (alloc.shift(1).dropna())).dropna() + 1).prod(axis=0) - 1

if monthly == True:
    portfolio_port = portfolio_port[portfolio_port.index.is_month_end == True]
drawdown = backtest_tmp.drawdown(portfolio_port)



input_price_N = input_price[
    (input_price.index >= portfolio_port.index[0]) &
    (input_price.index <= portfolio_port.index[-1])]

input_price_N = 100 * input_price_N / input_price_N.iloc[0, :]

portfolio_port.index = portfolio_port.index.date
drawdown.index = drawdown.index.date
input_price_N.index = input_price_N.index.date
alloc.index = alloc.index.date

result = pd.concat([portfolio_port,
                                     drawdown],axis=1)

# START_DATE = portfolio_port.index[0].strftime("%Y-%m-%d")
# END_DATE = portfolio_port.index[-1].strftime("%Y-%m-%d")
# Total_RET = round(float(portfolio_port[-1] / 100 - 1) * 100, 2)
# Anuuual_RET = round(float(((portfolio_port[-1] / 100) ** (
#         annualization / (len(portfolio_port) - 1)) - 1) * 100), 2)
# Anuuual_Vol = round(
#     float(np.std(portfolio_port.pct_change().dropna())
#           * np.sqrt(annualization) * 100), 2)
#
# MDD = round(float(min(drawdown) * 100), 2)
# Daily_RET = portfolio_port.pct_change().dropna()
import xlwings as xw

# 새 워크북 생성
wb = xw.Book()

# 첫 번째 시트(sheet1) 이름 설정
sheet1 = wb.sheets[0]
sheet1.name="Backtest"
#sheet1.name = "Backtest"  # sheet1 이름 변경
sheet1.range('A1').value = result  # 데이터 설정

# 두 번째 시트 추가 및 이름 설정
sheet2 = wb.sheets.add(name="Efficient Frontier")  # 새 시트 추가
sheet2.range('A1').value = EF  # 데이터 설정