import pandas as pd
import os, sys

pd.set_option("display.max_rows", None)    #設定最大能顯示1000rows
pd.set_option("display.max_columns", None) #設定最大能顯示1000columns
pd.set_option('display.width', 500) # 設置打印寬度 
pd.set_option('display.max_colwidth', 180)
pd.set_option("display.colheader_justify","center") #抬頭對齊用
pd.set_option('display.unicode.ambiguous_as_wide', True) #抬頭對齊用
pd.set_option('display.unicode.east_asian_width', True) #抬頭對齊用


while True :
    print('')
    xlsx = input('請輸入要查詢的Excel路徑: ')
    while True :
        name = input('請輸入品名或按【C】重新輸入Excel路徑: ')
        print('')       
        if name == 'c':
            break
        ws = []
        sheet = pd.read_excel(xlsx, sheet_name = None, index_col = 0,header=1)
        for line in sheet:
            try:
                df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                result = df.loc[(df['品名/規格'] == name),['品名/規格','單價','備註']]
                if result.empty:
                    continue
                else:
                    print(result)
                    print('')
                    ws.append(result)
            except :
                pass
        print('共有',len(ws),'筆要查詢的品名')