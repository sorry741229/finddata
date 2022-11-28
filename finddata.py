import pandas as pd
import os, sys
from colorama import init, Fore, Back #字體顏色
init(autoreset=True)#字體顏色
#https://clay-atlas.com/blog/2020/03/22/python-%E8%BC%B8%E5%87%BA%E5%AD%97%E4%B8%B2%E5%9C%A8%E7%B5%82%E7%AB%AF%E6%A9%9F%E4%B8%AD%E9%A1%AF%E7%A4%BA%E9%A1%8F%E8%89%B2/

pd.set_option("display.max_rows", None)    #設定最大能顯示1000rows
pd.set_option("display.max_columns", None) #設定最大能顯示1000columns
pd.set_option('display.width', 500) # 設置打印寬度 
pd.set_option('display.max_colwidth', 180)
pd.set_option("display.colheader_justify","center") #抬頭對齊用
pd.set_option('display.unicode.ambiguous_as_wide', True) #抬頭對齊用
pd.set_option('display.unicode.east_asian_width', True) #抬頭對齊用


while True :
    print('')
    xlsx = input('請輸入要查詢的路徑: ')
    while True:
        num = input('請輸入要查詢的格式_1.品名規格 2.廠商品號,或按【R】重新輸入路徑: ')
        if num == 'r' or num == 'R':
            break
        elif num == '1' or num == '2' :
            while True :    
                if num == '1' :
                    print('')
                    name = input('請輸入品名規格_或按【C】重選查詢的格式: ')
                    print('')       
                    if name == 'c' or name == 'C':
                        break
                    ws = []
                    sheet = pd.read_excel(xlsx, sheet_name = None, index_col = 0,header=1)
                    for line in sheet:
                        try:
                            df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                            df = df.fillna(' ')
                            result = df.loc[(df['品名/規格'] == name),['品名/規格','廠商品號','單價','備註']]
                            if result.empty:
                                continue
                            else:
                                print('')
                                print(result, end = Fore.RED +' ----此為'+ line + '工作表的內容')
                                print('')
                                ws.append(result)
                        except :
                            pass

                elif num == '2' :
                    print('')
                    name = input('請輸入廠商品號_或按【C】重選查詢的格式: ')
                    print('')       
                    if name == 'c':
                        break
                    ws = []
                    sheet = pd.read_excel(xlsx, sheet_name = None, index_col = 0,header=1)
                    for line in sheet:
                        try:
                            df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                            df = df.fillna(' ')
                            df.loc[mask, 'a'] = df['b']
                            result = df.loc[(df['廠商品號'] == name),['品名/規格','廠商品號','單價','備註']]
                            result = result["廠商品號"].str.contains(name)
                            if result.empty:
                                continue
                            else:
                                print('')
                                print(result,Fore.YELLOW +' ----此為'+ line + '工作表的內容')
                                print('')
                                ws.append(result)
                        except :
                            pass

                print('')
                print('此檔案共有',len(ws),'個工作表,有找到要查詢的品名')

        else:
            print(Fore.RED +'輸入錯誤，或按R後重新輸入路徑')   # 收到錯誤訊息，顯示錯誤
            print('')
