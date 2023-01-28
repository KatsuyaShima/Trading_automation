import pandas as pd
import subprocess #マーケットスピードを起動させるために使います。
import time
import os,math
import xlwings as xw #僕が好きなエクセルを動かすライブラリです。一番安定してます。
import pyautogui as py #キーボード操作で有名なライブラリです。
import datetime

#以下の時間内(ザラ場：9時～15時)では、銘柄を厳選したものだけに省略して、ザラ場内でスコアを再配点
base0 = datetime.time(8, 0, 0)
base9 = datetime.time(9, 0, 0)
base15 = datetime.time(15, 0, 0)
dt_now = datetime.datetime.now()
now = dt_now.time()
#now = datetime.time(8, 30, 0)

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

def start_rss():
    py.hotkey("win","m")#邪魔なウィンドウは最小化 pyautoguiを使うときは必ず使うことを推奨。
    #この下の行は重要です。binがあるところをディレクトリに設定しないと動きませんでした。
    os.chdir(r'C:\Users\katuy\AppData\Local\MarketSpeed2\Bin')
    global market_speed
    #EXEファイルを探して、下で指定してください。
    market_speed_path = r"C:\Users\katuy\AppData\Local\MarketSpeed2\Bin\MarketSpeed2.exe"
    market_speed = subprocess.Popen(market_speed_path)
    #market_speed.wait()
    time.sleep(30)
    #10秒程度だと、朝一では少ない。バージョンアップの確認が入ります。時にはもっと時間がかかるかもしれません。
    #画面中央のクリックを入れないと不安定
    py.click(960,540)
    #画面中央をクリック 画面中央なのか確認をお願いします。
    time.sleep(1)
    py.typewrite(r"PASSWORD")  #パスワード　ログインIDは記録させておくため入力しない
    time.sleep(2)
    py.press("Enter")
    time.sleep(60)
    #こちらも10秒では少ない

# RSSを停止します
def stop_rss():
    global market_speed
    market_speed.kill()

def excel_move():
    start_rss()
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
    time.sleep(10)
    #接続に10秒ぐらいかかる

    #9時～15時の間なら、選別された銘柄だけを再取得
    if (now >= base9 and now <= base15):
        #print("ok")
        ALL_Stock_code = pd.read_csv('/Users/katuy/Desktop/株式Python/出力csv3/df_filtered.csv', encoding="shift-jis")
        ALL_Stock_code = ALL_Stock_code.loc[:,['コード']]
    else:
        ALL_Stock_code = pd.read_csv('/Users/katuy/Desktop/All_Stock_data_20220930.csv', encoding="shift-jis")
        ALL_Stock_code = ALL_Stock_code.loc[:,['コード']]
    #銘柄売買用シートの指定
    sheet = wb.sheets['時系列データ取得']
    stock_code = ["N225","TSGM","DJIA"]
    for i in range(0,len(ALL_Stock_code)):
        stock_code.append(ALL_Stock_code.at[ALL_Stock_code.index[i],'コード'])
    #stock_code = stock_code[:b]
    stock_code.append("EOF")

    stock_code2=[]
    #前回途切れた最後の証券コードを取得
    path2 = r"C:\Users\katuy\Desktop\株式Python\Code_count.txt"
    isempty = os.stat(path2).st_size == 0#ファイルが空か判定
    #ファイルが空かつ、8時～15時以外なら実行
    if (now < base0 or base15 < now) and isempty == False:
        with open(path2) as f:
            Codes = f.read()
        Last_code=Codes[-4:]
    #print(Last_code)
        for t in range(0,len(stock_code)):
            if str(Last_code) == str(stock_code[t]):
                stock_code2=stock_code[t:]
    else:
        stock_code2=stock_code
    #print(stock_code2)

    #print('Code:'+str(stock_code[i])+'\t'+str(i+1)+' of '+str(len(ALL_Stock_code)))
    try:
        for k in range(0,len(stock_code2)):
            df_sheet_name=pd.DataFrame()
            #値のクリア
            sheet.range('D3').value = None
            sheet.range('E3').value = None
            sheet.range('F3').value = None

            # セルへ書き込む(更新)
            sheet.range('D3').value = stock_code2[k]
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
            #print(u)

            o=最後尾計算(df_sheet_name)

            df_sheet_name=df_sheet_name[u:o]
            #last_index = df_sheet_name.index[-1]
            #df_sheet_name=df_sheet_name.drop(last_index,axis=0)
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
                df_sheet_name.to_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code2[k-1]}.csv',encoding="shift-jis")
                #print(df_sheet_name)
                #print("a")
                #8時～15時以外なら実行
                if now < base0 or base15 < now:
                    #途中で途切れたコードを記憶して、そこからやり直すために、コードを記録していく
                    path2 = r"C:\Users\katuy\Desktop\株式Python\Code_count.txt"
                    f = open(path2,mode="a")
                    f.write(str(stock_code2[k]))
                    f.close()

                    #実行中を記録(他のスケジューラが動かないようにするため)
                    path3 = r"C:\Users\katuy\Desktop\株式Python\IsJikkoutyu.txt"
                    f = open(path3,mode="w")
                    f.write("YES")
                    f.close()
    #エラー処理
    except TypeError:
        #実行中を解除
        if now < base0 or base15 < now:
            path3 = r"C:\Users\katuy\Desktop\株式Python\IsJikkoutyu.txt"
            f = open(path3,mode="w")
            f.write("NO")
            f.close()
        time.sleep(5)
        #アプリ終了
        app = xw.apps.active
        wb.save()
        wb.close()
        app.kill()
        stop_rss()

    time.sleep(5)
    #実行中を解除
    if now < base0 or base15 < now:
        path3 = r"C:\Users\katuy\Desktop\株式Python\IsJikkoutyu.txt"
        f = open(path3,mode="w")
        f.write("NO")
        f.close()
    #アプリ終了
    app = xw.apps.active
    wb.save()
    wb.close()
    app.kill()
    stop_rss()
#excel_move()