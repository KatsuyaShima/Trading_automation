#NYダウの強さを計算
import pandas as pd
import time
import math
import xlwings as xw #僕が好きなエクセルを動かすライブラリです。一番安定してます。
import pyautogui as py #キーボード操作で有名なライブラリです。
import 楽天RSS資産確認_自動注文

#NaNの削除(上場した日からの取得)
def 上場初日計算(df_sheet_name):
    df_sheet_name=df_sheet_name.reset_index()
    for i in range(0,len(df_sheet_name)):
        if df_sheet_name.at[df_sheet_name.index[i],"出来高"]>0:
            return i
def 最後尾計算(df_sheet_name):
    df_sheet_name=df_sheet_name.reset_index()
    for i in range(0,len(df_sheet_name)):
        if df_sheet_name.at[df_sheet_name.index[i],"日付"]=="--------":
            return i

def IsStrongDow():
    #朝8:55にNYDOWの時系列データを更新
    楽天RSS資産確認_自動注文.start_marketspeed2()
    py.hotkey("win","m")
    #好きな場所にRSSの自動売買の設定をしたエクセルファイルをおいてください。
    wb = xw.Book(r'C:\Users\katuy\Desktop\株価時系列取得.xlsm')
    wb.app.activate(steal_focus=False)

    #下は重要な行。エクセルのアドインを読み込みます。xlwingsだと通常起動ではアドインを読み込まないことが多いです。
    addin_path=r'C:\Users\katuy\AppData\Local\MarketSpeed2\Bin\rss\MarketSpeed2_RSS_32bit.xll'
    wb.app.api.RegisterXLL(addin_path)
    time.sleep(5)

    #画面中央をクリック　またもや画面中央なのか確認をお願いします。
    py.click(960,540)
    time.sleep(1)
    #エクセルのメニューボタンをaltを押して動かしてゆきます。
    py.press("alt")
    time.sleep(0.1)
    #マーケットスピート2のメニューを選ぶ。実装する際にはエクセルを起動してaltを押した後、メニュー欄を確認してください。
    py.press("y")
    py.press("3")
    time.sleep(0.1)
    #マーケットスピード2に接続する。自動的に接続する設定にしていると、注意が必要です。僕は能動的に接続するほうが安心だと思います。
    py.press("y")
    py.press("1")
    time.sleep(20)
    #接続に10秒ぐらいかかる

    #銘柄売買用シートの指定
    sheet = wb.sheets['時系列データ取得']
    stock_code = ["DJIA"]
    stock_code.append("EOF")

    for k in range(0,len(stock_code)):
        df_sheet_name=pd.DataFrame()
        #値のクリア
        sheet.range('D3').value = None
        sheet.range('E3').value = None
        sheet.range('F3').value = None

        # セルへ書き込む(更新)
        sheet.range('D3').value = stock_code[k]
        sheet.range('E3').value = "D" #日足
        sheet.range('F3').value = 245 #1年間の営業日
            
        if k <1:
            time.sleep(0.5)

        #一旦セーブ
        wb.save(r'C:\Users\katuy\Desktop\株価時系列取得.xlsm')

        #取得の間隔
        time.sleep(0.5)
        
        #開く
        df_sheet_name = pd.read_excel(r'C:\Users\katuy\Desktop\株価時系列取得.xlsm', sheet_name='時系列データ取得',header=5,index_col=0)

        u = 上場初日計算(df_sheet_name)

        o=最後尾計算(df_sheet_name)

        df_sheet_name=df_sheet_name[u:o]
        #売買代金(千万円)の作成
        Closelist=[]
        
        for y in range(0,len(df_sheet_name)):
            if math.isnan(df_sheet_name.at[df_sheet_name.index[y],'出来高'])==False:
                Closelist.append(round(df_sheet_name.at[df_sheet_name.index[y],"終値"]*df_sheet_name.at[df_sheet_name.index[y],"出来高"]/10000000,2))#1000万で割る
            else:
                Closelist.append("")
        
        df_sheet_name=df_sheet_name.reset_index()
        df_sheet_name=df_sheet_name.set_axis(['datetime', 'volume', 'first','high','low','close'], axis='columns')
        df_sheet_name["平均売買代金(千万円)"]=Closelist
        if k>0:
            #書き出し
            df_sheet_name.to_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code[k-1]}.csv',encoding="shift-jis")

    #アプリ終了
    app = xw.apps.active
    wb.save()
    wb.close()
    app.kill()
    楽天RSS資産確認_自動注文.stop_marketspeed2()

    time.sleep(3)

    df_NY_dow=pd.read_csv(r"C:/Users/katuy/Desktop/株式Python/出力csv/DJIA.csv",encoding="cp932")

    last_index=len(df_NY_dow)-1

    #25日平均線を計算する
    Twentyfive_days_ave=0
    for i in range(0,25):
        Twentyfive_days_ave=df_NY_dow.at[df_NY_dow.index[last_index-i],"close"]+Twentyfive_days_ave
    Twentyfive_days_ave=Twentyfive_days_ave/25
    #本日の終値の25日平均移動線からの乖離率を計算
    kairi=round(df_NY_dow.at[df_NY_dow.index[last_index],"close"]/Twentyfive_days_ave,3)

    #5日平均線を計算する
    five_days_ave=0
    for i in range(0,5):
        five_days_ave=df_NY_dow.at[df_NY_dow.index[last_index-i],"close"]+five_days_ave
    five_days_ave=five_days_ave/5
    #本日の終値の5日平均移動線からの乖離率を計算
    kairi_five=round(df_NY_dow.at[df_NY_dow.index[last_index],"close"]/five_days_ave,3)

    #print(kairi)
    #print(kairi_five)

    #25日平均移動線からの乖離率が-2%未満かつ、5日乖離率が-1.5%未満ならば、現在の保有銘柄をこの寄付で成り行き売りし、その日は取引中止
    path = r"C:\Users\katuy\Desktop\株式Python\SuspendTrading_days.txt"
    if kairi<=0.980 and kairi_five<=0.985:
        f = open(path,'w')
        f.write(str(1))#取引中止期間を1日に設定
        f.close()
        楽天RSS資産確認_自動注文.send_line_notify("\n ダウ相場が激弱なので、\n 本日の取引は中止します。")
    else:
        with open(path) as f:
            SuspendTrading_days0 = f.read()
            SuspendTrading_days0 = int(SuspendTrading_days0)
            #取引中止期間が1日以上なら1日減らして書き換える
            if SuspendTrading_days0>=1:
                SuspendTrading_days=int(SuspendTrading_days0)-1
                
                f = open(path,'w')
                f.write(str(SuspendTrading_days))
                f.close()
        楽天RSS資産確認_自動注文.send_line_notify("\n 現在の相場は正常です。")
#IsStrongDow()