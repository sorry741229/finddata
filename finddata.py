import pandas as pd
import os, sys
from colorama import init, Fore, Back #字體顏色
init(autoreset=True)#字體顏色
#https://clay-atlas.com/blog/2020/03/22/python-%E8%BC%B8%E5%87%BA%E5%AD%97%E4%B8%B2%E5%9C%A8%E7%B5%82%E7%AB%AF%E6%A9%9F%E4%B8%AD%E9%A1%AF%E7%A4%BA%E9%A1%8F%E8%89%B2/
import time


pd.set_option("display.max_rows", None)    #設定最大能顯示1000rows
pd.set_option("display.max_columns", None) #設定最大能顯示1000columns
pd.set_option('display.width', 500) # 設置打印寬度 
pd.set_option('display.max_colwidth', 180)
pd.set_option("display.colheader_justify","center") #抬頭對齊用
pd.set_option('display.unicode.ambiguous_as_wide', True) #抬頭對齊用
pd.set_option('display.unicode.east_asian_width', True) #抬頭對齊用

#授權時間
def now():
    return time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
s = '2023-04-23 14:20:59'

if now() > s :
    print('f0614 Error :some files are missing or connection time out')
    os._exit(0)


while True:
    ans = '2681'
    x = 3 #初始機會
    while x > 0 :
        x = x -1
        pwd = input('請輸入登入密碼: ')
        if pwd == ans:
            print('')
            print('')
            break
        else:
            print('密碼錯誤!')
            if x > 0:
                print('還有', x,'次機會')
            else:
                print('已輸入超過三次，程式結束')
                print('')
                print('')
                print('3秒後程式關閉', end = '')
                for i in range(6):
                    print("",end = '',flush = True)  #flush - 输出是否被缓存通常决定于 file，但如果 flush 关键字参数为 True，流会被强制刷新
                    time.sleep(0.5)
                    print('')
                os._exit()
    break
os.system('cls') #登入後清除畫面


print('')
print(Fore.GREEN +"{:=^100s}".format("群旭CNC_刀具價格查詢Ver1.0"))
print('登入成功')
print('')


while True :
    print('')
    xlsx = input('請輸入要查詢的路徑: ')
    while True:
        print('')
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
                        df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                        df = df.fillna(' ')
                        try:
                            result = df.loc[(df['品名/規格'] == name),['品名/規格','廠商品號','單價','備註']]
                            if result.empty :
                                try:
                                    df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                                    df = df.fillna(' ')
                                    result = df["品名/規格"].str.contains(name)       #將dataframe轉成文字比對name,得到布林值
                                    filter_result = df[result]                       #用dataframe操作取出布林後之篩選結果
                                    #filter_result.set_index(keys = ['品名/規格','廠商品號','單價','備註'],inplace=True) #此為調整索引值順序
                                    final_result = filter_result[['品名/規格','廠商品號','單價','備註']]
                                    #filter_result.head(1) #只顯示幾行資料   
                                    if final_result.empty:
                                        continue
                                    else:
                                        print('')
                                        print(final_result)
                                        print(Fore.YELLOW +' ----此區塊為'+ line.strip() + '工作表的內容')
                                        print('')
                                        ws.append(final_result)
                                        print("{:=^100s}".format("群旭科技_CNC"))
                                except :
                                    continue

                            else:
                                try:
                                    result = df.loc[(df['品名/規格'] == name),['品名/規格','廠商品號','單價','備註']]
                                    if result.empty:
                                        continue
                                    else:
                                        print('')
                                        print(result)
                                        print(Fore.YELLOW +' ----此區塊為'+ line.strip() + '工作表的內容')
                                        print('')
                                        ws.append(result)
                                        print("{:=^100s}".format("群旭科技_CNC"))
                                except :
                                    pass
                        except :
                            pass

                elif num == '2' :
                    print('')
                    name = input('請輸入廠商品號_或按【C】重選查詢的格式: ')
                    print('')       
                    if name == 'c' or name == 'C':
                        break
                    ws = []
                    sheet = pd.read_excel(xlsx, sheet_name = None, index_col = 0,header=1)
                    for line in sheet:
                        df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                        df = df.fillna(' ')
                        try:
                            result = df.loc[(df['廠商品號'] == name),['品名/規格','廠商品號','單價','備註']]
                            if result.empty :
                                try:
                                    df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header=1)
                                    df = df.fillna(' ')
                                    result = df["廠商品號"].str.contains(name)       #將dataframe轉成文字比對name,得到布林值
                                    filter_result = df[result]                       #用dataframe操作取出布林後之篩選結果
                                    #filter_result.set_index(keys = ['品名/規格','廠商品號','單價','備註'],inplace=True) #此為調整索引值順序
                                    final_result = filter_result[['品名/規格','廠商品號','單價','備註']]
                                    #filter_result.head(1) #只顯示幾行資料   
                                    if final_result.empty:
                                        continue
                                    else:
                                        print('')
                                        print(final_result)
                                        print(Fore.YELLOW +' ----此區塊為'+ line.strip() + '工作表的內容')
                                        print('')
                                        ws.append(final_result)
                                        print("{:=^100s}".format("群旭科技_CNC"))
                                except :
                                    continue

                            else:
                                try:
                                    result = df.loc[(df['廠商品號'] == name),['品名/規格','廠商品號','單價','備註']]
                                    if result.empty:
                                        continue
                                    else:
                                        print('')
                                        print(result)
                                        print(Fore.YELLOW +' ----此區塊為'+ line.strip() + '工作表的內容')
                                        print('')
                                        ws.append(result)
                                        print("{:=^100s}".format("群旭科技_CNC"))
                                except :
                                    pass
                        except :
                            pass                               
                print('')
                print(Fore.GREEN + '此檔案共有',len(ws),Fore.GREEN + '個工作表,有找到要查詢的品名')

        else:
            print(Fore.RED +'輸入錯誤，或按R後重新輸入路徑')   # 收到錯誤訊息，顯示錯誤
            print('')
