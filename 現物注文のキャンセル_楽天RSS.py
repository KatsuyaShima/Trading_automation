#大引け前の14:00に注文キャンセルを行うコード
import pandas as pd
import time
import random
import xlwings as xw #僕が好きなエクセルを動かすライブラリです。一番安定してます。
import pyautogui as py #キーボード操作で有名なライブラリです。
import 楽天RSS資産確認_自動注文

def Cancel_Order():
    #以下、注文用のエクセル起動
    楽天RSS資産確認_自動注文.start_marketspeed2()
    py.hotkey("win","m")
    #好きな場所にRSSの自動売買の設定をしたエクセルファイルをおいてください。
    wb = xw.Book(r'C:\Users\katuy\Desktop\注文キャンセル用.xlsm')
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
    time.sleep(10.0)
    #接続に10秒ぐらいかかる

    #銘柄売買用シートの指定
    sheet = wb.sheets['注文キャンセル']
    sheet.range('C3').value = None
    sheet.range('D3').value = None

    #一旦セーブ
    wb.save(r'C:\Users\katuy\Desktop\注文キャンセル用.xlsm')
    time.sleep(2)

    #銘柄売買用シートの指定
    #sheet = wb.sheets['資産・買付余力']
    df_sheet_name=pd.DataFrame()
    #開く
    df_sheet_name = pd.read_excel(r'C:\Users\katuy\Desktop\注文キャンセル用.xlsm', sheet_name='注文キャンセル',header=5,index_col=0).reset_index()
    print(df_sheet_name)
    print(len(df_sheet_name))

    #発注ONにする
    #画面中央をクリック
    py.click(960,540)
    time.sleep(1)
    py.press("alt")
    time.sleep(0.1)
    #マーケットスピート2のメニューを選ぶ。実装する際にはエクセルを起動してaltを押した後、メニュー欄を確認してください。
    py.press("y")
    py.press("3")
    time.sleep(0.1)
    #発注ON
    py.press("y")
    py.press("2")
    time.sleep(3)
    py.doubleClick(1020,615)
    time.sleep(3)

    #注文が一つ以上あれば実行
    if len(df_sheet_name)>=2:
        Tyuumon_code_list=[]
        #通常注文状況を確認して発注中であれば、その行の注文番号をリストに追加する
        for i in range(0,len(df_sheet_name)):
            if (df_sheet_name.at[df_sheet_name.index[i],"通常注文状況"]=="執行待ち" 
            or df_sheet_name.at[df_sheet_name.index[i],"通常注文状況"]=="執行中"):
                Tyuumon_code_list.append(df_sheet_name.at[df_sheet_name.index[i],"注文番号"])
        #print(Tyuumon_code_list)

        if len(Tyuumon_code_list)>0:
            #上のリストに保存した注文番号をある分だけ入力する
            for i in range(0,len(Tyuumon_code_list)):
                #注文銘柄が1つの場合
                if i<=0:
                    random_number=random.randint(0,100)
                #注文銘柄が2つあれば、値をリセットして再入力
                if i >=1:
                    time.sleep(3)
                    sheet.range('C3').value = None
                    sheet.range('D3').value = None
                    random_number=random.randint(101,200)

                sheet.range('C3').value = random_number
                sheet.range('D3').value = Tyuumon_code_list[i]
            楽天RSS資産確認_自動注文.send_line_notify("\n 注文を"+str(len(Tyuumon_code_list))+"件、\n 大引前キャンセルしました。")
        else:
            楽天RSS資産確認_自動注文.send_line_notify("\n 注文が1件もないので、\n 大引前キャンセルはしません。")

    time.sleep(5)
    #アプリ終了
    app = xw.apps.active
    wb.save()
    wb.close()
    app.kill()
    楽天RSS資産確認_自動注文.stop_marketspeed2()
#Cancel_Order()