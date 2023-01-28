#楽天RSS注文・資産確認
import pandas as pd
import subprocess #マーケットスピードを起動させるために使います。
import time
import os,random
import xlwings as xw #僕が好きなエクセルを動かすライブラリです。一番安定してます。
import pyautogui as py #キーボード操作で有名なライブラリです。
import requests,datetime

base8 = datetime.time(8, 0, 0)
base9 = datetime.time(9, 0, 0)
base10 = datetime.time(10, 0, 0)
base14 = datetime.time(14, 0, 0)
base15 = datetime.time(15, 0, 0)
dt_now = datetime.datetime.now()
now = dt_now.time()
#now = datetime.time(8, 30, 0)

def start_marketspeed2():
    py.hotkey("win","m")#邪魔なウィンドウは最小化 pyautoguiを使うときは必ず使うことを推奨。
    #この下の行は重要です。binがあるところをディレクトリに設定しないと動きませんでした。
    os.chdir(r'C:\Users\katuy\AppData\Local\MarketSpeed2\Bin')
    global market_speed
    #EXEファイルを探して、下で指定してください。
    market_speed_path = r"C:\Users\katuy\AppData\Local\MarketSpeed2\Bin\MarketSpeed2.exe"
    market_speed = subprocess.Popen(market_speed_path)
    #market_speed.wait()
    time.sleep(20)
    #10秒程度だと、朝一では少ない。バージョンアップの確認が入ります。時にはもっと時間がかかるかもしれません。
    #画面中央のクリックを入れないと不安定
    py.click(960,540)
    #画面中央をクリック 画面中央なのか確認をお願いします。
    time.sleep(1)
    py.typewrite(r"Rakuten_Robotto")  #パスワード　ログインIDは記録させておくため入力しない
    time.sleep(2)
    py.press("Enter")
    time.sleep(60)
    #こちらも10秒では少ない

# RSSを停止します
def stop_marketspeed2():
    global market_speed
    market_speed.kill()

def excel_move():
    start_marketspeed2()
    py.hotkey("win","m")
    #好きな場所にRSSの自動売買の設定をしたエクセルファイルをおいてください。
    wb = xw.Book(r'C:\Users\katuy\Desktop\資産確認用.xlsm')
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

    #一旦セーブ
    wb.save(r'C:\Users\katuy\Desktop\資産確認用.xlsm')
    time.sleep(2)

    #銘柄売買用シートの指定
    #sheet = wb.sheets['資産・買付余力']
    df_sheet_name=pd.DataFrame()
    #開く
    df_sheet_name = pd.read_excel(r'C:\Users\katuy\Desktop\資産確認用.xlsm', sheet_name='資産・買付余力',header=2,index_col=0).reset_index()
    #現物買付可能額を保留(この後、保有銘柄数が2つ以上なら、入れ替えを行うため。)
    Kaitsuke_kanougaku=df_sheet_name.at[df_sheet_name.index[0],"現物買付可能額"]
    #print(Kaitsuke_kanougaku)

    #保有銘柄が2つある場合
    if len(df_sheet_name)>2:
        #print(len(df_sheet_name))
        df_last_index = df_sheet_name[-1:]#最終行を別のデータフレームに保留
        df_sheet_name = df_sheet_name[:-1]#時価評価額を降順に並び替える為に、一旦最終行を削除
        df_sheet_name = df_sheet_name.sort_values("時価評価額",ascending=False)
        #最終行を元に戻す
        df_sheet_name=pd.concat([df_sheet_name,df_last_index])
        print(df_sheet_name)

        Code_name2=df_sheet_name.at[df_sheet_name.index[1],'銘柄名称']#2行目の銘柄名を取得
        Suuryo2=df_sheet_name.at[df_sheet_name.index[1],'保有数量']#2行目の保有数を取得
        kounyu_tanka2=df_sheet_name.at[df_sheet_name.index[1],'平均取得価額']#2行目の保有数を取得
        df_code_shutoku=pd.read_csv(f'/Users/katuy/Desktop/All_Stock_data_20220930.csv',encoding="shift-jis")
        for i in range(0,len(df_code_shutoku)):
            if df_code_shutoku.at[df_code_shutoku.index[i],"銘柄名"]==Code_name2:
                Code2=df_code_shutoku.at[df_code_shutoku.index[i],"コード"]
        MeigaraSuu=2#銘柄保有数を2にする
    elif len(df_sheet_name)==2:#保有銘柄数が1つの場合
        Code2=0
        Suuryo2=0
        kounyu_tanka2=0
        MeigaraSuu=1#銘柄保有数を1にする(売るときだけ参照するので、保有していない場合は無視)
        

    #資産総額を取得
    if df_sheet_name.at[df_sheet_name.index[0],"時価評価額"]!="--------":
        if MeigaraSuu == 1:
            Sisan_Sougaku=df_sheet_name.at[df_sheet_name.index[0],"時価評価額"]+Kaitsuke_kanougaku
        elif MeigaraSuu == 2:
            Sisan_Sougaku=df_sheet_name.at[df_sheet_name.index[0],"時価評価額"]+df_sheet_name.at[df_sheet_name.index[1],"時価評価額"]+Kaitsuke_kanougaku
        #平均取得価額を取得
        kounyu_tanka=df_sheet_name.at[df_sheet_name.index[0],'平均取得価額']
        Code_name=df_sheet_name.at[df_sheet_name.index[0],'銘柄名称']
        Suuryo=df_sheet_name.at[df_sheet_name.index[0],'保有数量']
        KounyuFlag=0
        
    else:#保有銘柄が0の場合の処理(KounyuFlag=1)
        Sisan_Sougaku=Kaitsuke_kanougaku
        kounyu_tanka=105#エラー対策の為に適当にいれておく
        Code_name="トヨタ自動車"#エラー対策の為に適当にいれておく
        Suuryo=100
        KounyuFlag=1
        Code2=0
        Suuryo2=0
        kounyu_tanka2=0
        MeigaraSuu=0

    time.sleep(5)
    #アプリ終了
    app = xw.apps.active
    wb.save()
    wb.close()
    app.kill()
    stop_marketspeed2()
    return KounyuFlag,Sisan_Sougaku,kounyu_tanka,Code_name,Suuryo,Kaitsuke_kanougaku,Code2,Suuryo2,MeigaraSuu,kounyu_tanka2

#一日に一回のみ売る(午後の集計でリセット)
def Nariyuki_Sasine_Uri(Code,Suuryo,IsNariyuki,Kabuka,Code2,Suuryo2,MeigaraSuu):
    path = r"C:\Users\katuy\Desktop\株式Python\IsAlreadySoldToday.txt"
    with open(path) as f:
        IsSold = f.read()

    #1日に2回は売らせない対策(今のところ、朝8:50に一度しか動かさない為、この対策は必要ない)
    if IsSold == "NO":
        #今から売る銘柄をcsvに上書き保存しておく(次の購入候補から除外する為)
        path3 = r"C:\Users\katuy\Desktop\株式Python\Zenkai_Hoyuu.txt"
        f = open(path3,'w')
        f.write(str(Code))
        f.close()

        #以下、注文用のエクセル起動
        start_marketspeed2()
        py.hotkey("win","m")
        #好きな場所にRSSの自動売買の設定をしたエクセルファイルをおいてください。
        wb = xw.Book(r'C:\Users\katuy\Desktop\注文用_売り.xlsm')
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
        sheet = wb.sheets['現物注文']
        sheet.range('A4').value = None
        sheet.range('B4').value = None
        sheet.range('C4').value = None
        sheet.range('D4').value = None
        sheet.range('E4').value = None

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

        for i in range(0,MeigaraSuu):
            if i <=0:
                random_ID=random.randint(1,100)
                #発注ID
            if i >= 1:#もし2週目に入ったら、各値リセットして二番目の保有銘柄の各値に修正
                sheet.range('A4').value = None
                sheet.range('B4').value = None
                sheet.range('C4').value = None
                sheet.range('D4').value = None
                sheet.range('E4').value = None
                time.sleep(3)
                random_ID = random.randint(101,200)
                Code = Code2
                Suuryo = Suuryo2
                IsNariyuki = 0 #成行売り

            sheet.range('A4').value = random_ID
            #証券コード
            sheet.range('B4').value = Code
            #数量
            sheet.range('C4').value = Suuryo
            #成行：0 / 指値：1
            sheet.range('D4').value = IsNariyuki
            #成り行きと指値によって処理を分ける
            if IsNariyuki == 0:
                #価格_成り行き注文は指値しない
                sheet.range('E4').value = None
            else:
                #価格_指値で株価指定
                sheet.range('E4').value = Kabuka
            
            if i < 1: #1週目
                IsUri=1 # 1なら売り
                Sisan=0#エラー対策
                Kaitsuke=0#エラー対策
                main(Sisan,Kaitsuke,Code,Suuryo,Kabuka,IsUri,IsNariyuki)
            if i >= 1: #2週目
                send_line_notify(str(Code2)+"を成行売りしました。")
            time.sleep(5)
        #アプリ終了
        app = xw.apps.active
        wb.save()
        wb.close()
        app.kill()
        stop_marketspeed2()
        f = open(path,'w')
        f.write("YES")
        f.close()

def Uri():
    #買いと売りで、エクセルを二回呼び出さないようにする為
    KounyuFlag,Sisan_Sougaku,kounyu_tanka,Code_name,Suuryo,Kaitsuke_kanougaku,Code2,Suuryo2,MeigaraSuu,kounyu_tanka2=excel_move()

    #相場の強弱を判断
    path100 = r"C:\Users\katuy\Desktop\株式Python\SuspendTrading_days.txt"
    with open(path100) as f:
        SuspendTrading_days = f.read()
        SuspendTrading_days = int(SuspendTrading_days)

    #保有銘柄がある場合は以下の条件に一致すれば、売り注文の確認処理を行う
    df_kouho=pd.read_csv(r'C:\Users\katuy\Desktop\株式Python\df_BUY_FINAL.csv',encoding="Shift-jis")
    list=[9999,0,0,0,0,0,0,0,0]
    df_insert = pd.DataFrame(list).T
    df_insert = df_insert.set_axis(["コード","Score (高値からの下落率)","Score (底値からの上昇率)","Score (5日乖離率)","Score (20日乖離率)","Score (直近5日間での上昇率)","Score (流動性)","Score (直近1の強さ)","合計点数"], axis=1)
    #print(df_insert)

    #候補が0の場合はkコード9999を追加
    if len(df_kouho)<1:
        df_kouho=pd.concat([df_kouho,df_insert])
    
    #銘柄名とコードの変換(楽天RSSからは保有銘柄コードが表示されない為)
    df_code_shutoku=pd.read_csv(f'/Users/katuy/Desktop/All_Stock_data_20220930.csv',encoding="shift-jis")
    for i in range(0,len(df_code_shutoku)):
        if df_code_shutoku.at[df_code_shutoku.index[i],"銘柄名"]==Code_name:
            Code=df_code_shutoku.at[df_code_shutoku.index[i],"コード"]
    df_code_jikeiretsu=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{Code}.csv',encoding="shift-jis")
    #print(Code)
    #print(df_code_jikeiretsu)

    #現在の保有銘柄の保有日数の取得
    path2 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
    with open(path2) as f:
        DaysFromBought = f.read()
    DaysFromBought=int(DaysFromBought)#現在の保有日数

    #8時～9時までの間にのみ売りを発動(8:50だけの一日一回のみ売る)
    if (base8 <= now and now <= base9):
    #保有していたら、売りの処理を実行
        if KounyuFlag==0 :
            #次の関数で買えるようにKounyuFlagを1に設定しなくてもよい(タスクマネージャーでまた動かす)
            today=datetime.date.today()
            #最新日の行
            last_index=len(df_code_jikeiretsu)-1

            #利確モードかチェック
            #当日の5日平均移動線の乖離率を求める(高値ベース)
            average_five_today=df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index],"high"]
            #print(average_five_today)
            for r in range(1,5):
                average_five_today=average_five_today+df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index-r],"close"]
            average_five_today=average_five_today/5
            #print(average_five_today)

            #前日の5日平均移動線の乖離率を求める(高値ベース)
            average_five_lastday=df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index-1],"high"]
            for r in range(2,6):
                average_five_lastday=average_five_lastday+df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index-r],"close"]
            average_five_lastday=average_five_lastday/5

            #当日の75日平均移動線の乖離率を求める(高値ベース)
            average_seventyfive_today=df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index],"high"]
            for r in range(1,75):
                average_seventyfive_today=average_seventyfive_today+df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index-r],"close"]
            average_seventyfive_today=average_seventyfive_today/75
            #print(average_seventyfive_today)

            #前日の高値の5日平均移動線からの乖離率が10%を越えていた場合
            if df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index-1],"high"]/average_five_lastday > 1.1:
                #翌日の寄付で成行売り注文
                send_line_notify("\n 保有銘柄："+str(Code)+"\n 利確(2日目)\n 寄付成行売りで発注します。")
                IsNariyuki=0 #成行:0
                Kabuka=10000#エラー対策
                Nariyuki_Sasine_Uri(Code,Suuryo,IsNariyuki,Kabuka,Code2,Suuryo2,MeigaraSuu)
                
            #当日の高値の5日平均移動線からの乖離率が10%を越えていた場合
            elif df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index],"high"]/average_five_today > 1.1:
                #当日の75日平均移動線乖離率がマイナスの場合(まだ上昇余地がある場合は保留_高値ベース)
                if df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index],"high"]/average_seventyfive_today < 0:
                    #保有予定日数を残り1日にリセットする
                    path3 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
                    f = open(path3,'w')
                    f.write(str(1))
                    f.close()
                    #LINEメッセージ
                    send_line_notify("\n 保有銘柄："+str(Code)+"\n 利確(初日_75日線下)\n 次の寄付で成行売りします。")
                else:
                    send_line_notify("\n 保有銘柄："+str(Code)+"\n 利確(初日_75日線上)\n 寄付成行売りで発注します。")
                    IsNariyuki=0 #成行:0
                    Kabuka=10000#エラー対策
                    Nariyuki_Sasine_Uri(Code,Suuryo,IsNariyuki,Kabuka,Code2,Suuryo2,MeigaraSuu)

            #取引制限がある場合は成り行き売り
            elif SuspendTrading_days>=1:
                send_line_notify("\n 保有銘柄："+str(Code)+"\n NYダウ相場が激弱なので、\n 寄付成行売りで発注します。")
                IsNariyuki=0 #成行:0
                Kabuka=10000#エラー対策
                Nariyuki_Sasine_Uri(Code,Suuryo,IsNariyuki,Kabuka,Code2,Suuryo2,MeigaraSuu)

            else:#利確モードではない場合
                #損切モードかチェック
                if (df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index],"close"]/kounyu_tanka)<=0.95:
                    #取得価額が直近の終値と比較して、終値ベースで-5%以上の含み損になった場合に、翌日の寄付で成行売り注文
                    send_line_notify("\n 保有銘柄："+str(Code)+"\n 損切5%"+"\n 寄付成行売りで発注します。")
                    IsNariyuki=0 #成行：0
                    Kabuka=10000#エラー対策
                    Nariyuki_Sasine_Uri(Code,Suuryo,IsNariyuki,Kabuka,Code2,Suuryo2,MeigaraSuu)

                #利確モードでも損切モードでもない場合、保有銘柄がスコア1位で無ければ成行で売る
                #スコアと保有日数をチェックして買い替えるか判断
                elif Code!=df_kouho.at[df_kouho.index[0],"コード"]:
                    if DaysFromBought<=0:
                        send_line_notify("\n 保有銘柄："+str(Code)+"\n スコアが1位ではないです。\n 保有日数が3営業日を超えた為\n 当該銘柄の売却を行います。")
                        IsNariyuki=0 #成行：0
                        Kabuka=10000#エラー対策
                        Nariyuki_Sasine_Uri(Code,Suuryo,IsNariyuki,Kabuka,Code2,Suuryo2,MeigaraSuu)
                    elif DaysFromBought==1:
                        send_line_notify("\n 保有銘柄："+str(Code)+"\n 売却予定日まであと："+str(DaysFromBought)+"日"+"\n スコアが1位ではないです。"+"\n 次の営業日の寄り付きで\n 成行売りの可能性アリです。")
                    else :
                        send_line_notify("\n 保有銘柄："+str(Code)+"\n 売却予定日まであと："+str(DaysFromBought)+"日"+"\n スコアが1位ではないです。")

                #スコアが1位の場合は現状維持
                else:
                    #保有銘柄の保有日数の取得
                    path2 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
                    with open(path2) as f:
                        DaysFromBought = f.read()
                    DaysFromBought=int(DaysFromBought)#現在の保有日数

                    send_line_notify("\n 保有銘柄："+str(Code)+"\n 売却予定日まであと："+str(DaysFromBought)+"日"+"\n スコアが1位、現状維持です。")

    #そのまま買いの関数に値を横流し
    return KounyuFlag,Sisan_Sougaku,kounyu_tanka,Kaitsuke_kanougaku

def Suuryou_keisan():
    KounyuFlag,Sisan_Sougaku,kounyu_tanka,Kaitsuke_kanougaku=Uri()
    path2 = r"C:\Users\katuy\Desktop\株式Python\Zenkai_Hoyuu.txt"
    with open(path2) as f:
        Code = f.read()
    #print("a"+str(Code))
    #購入銘柄を取得
    df_kouho=pd.read_csv(r'C:\Users\katuy\Desktop\株式Python\df_BUY_FINAL.csv',encoding="Shift-jis")
    #購入候補が1つ以上あれば買いを実行

    cols=["コード","Score1 (高値からの下落率)","Score2 (底値からの上昇率)","Score3 (5日乖離率)","Score4 (20日乖離率)","Score5 (直近5日間での上昇率)","Score6 (流動性)","Score7 (直近の流れの強さ)","Score8 (直近1の強さ)","合計点数"]
    df_kouho2=pd.DataFrame(columns=cols)
    #現在保有している銘柄を除外する
    Code2=int(Code)
    length=len(df_kouho)
    if length>=1:
        for i in range(0,length):
            if df_kouho.at[df_kouho.index[i],"コード"]==Code2:
                df_kouho2=pd.concat([df_kouho2,df_kouho.drop(df_kouho.index[i])],ignore_index=True)
                length-=1
            else:
                df_kouho2=pd.concat([df_kouho2,df_kouho[i:i+1]],ignore_index=True)
                    #print("test")
            #print(df_kouho2)
            #前の保有銘柄を除外した表からスコア1位,2位の銘柄を選択して購入候補に
            if len(df_kouho2)>=1:
                kouho=df_kouho2.at[df_kouho2.index[0],"コード"]
                df_data=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{kouho}.csv',encoding="shift-jis")
                last_gyo1=len(df_data)
                Zenjitu_owarine1=df_data.at[df_data.index[last_gyo1-1],"close"]
                Yuuyo1=int(Zenjitu_owarine1*1.02)

                tanngen_suu1=Sisan_Sougaku/Yuuyo1
                tanngen_suu2=round(tanngen_suu1,-2)
                tanngen_suu2=int(tanngen_suu2)
                if tanngen_suu2>tanngen_suu1:
                    tanngen_suu2=tanngen_suu2-100

                if len(df_kouho2)>=2:
                    kouho2=df_kouho2.at[df_kouho2.index[1],"コード"]
                    df_data2=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{kouho2}.csv',encoding="shift-jis")
                    last_gyo2=len(df_data2)
                    Zenjitu_owarine2=df_data2.at[df_data2.index[last_gyo2-1],"close"]
                    Yuuyo2=int(Zenjitu_owarine2*1.02)
                    #Rank1位購入後に余った金額で、Rank2が100株以上買えそうか計算
                    tanngen_suu3=(Sisan_Sougaku-(tanngen_suu2*Yuuyo1))/Yuuyo2
                    tanngen_suu4=round(tanngen_suu3,-2)
                    tanngen_suu4=int(tanngen_suu4)
                    if tanngen_suu4>tanngen_suu3:
                        tanngen_suu4=tanngen_suu4-100

                    if tanngen_suu4>=100:
                        Rank2_buy=1 #1なら、スコア2位の銘柄kouho2をYuuyo2円で、tangen_suu4株購入
                else:
                    Rank2_buy=0
                    kouho2=0
                    Yuuyo2=0
                    tanngen_suu4=0
            else:
                KounyuFlag=0
                kouho=0
                Yuuyo1=0
                tanngen_suu2=0
                Sisan_Sougaku=1#資産と買付可能額が一致しないようにする
                Kaitsuke_kanougaku=0#資産と買付可能額が一致しないようにする
                Rank2_buy=0
                kouho2=0
                Yuuyo2=0
                tanngen_suu4=0  
            #print(df_data)

    #購入候補が一つもない時はKounyuFlagをオフにする
    else:
        KounyuFlag=0
        kouho=0
        Yuuyo1=0
        tanngen_suu2=0
        Sisan_Sougaku=1#資産と買付可能額が一致しないようにする
        Kaitsuke_kanougaku=0#資産と買付可能額が一致しないようにする
        Rank2_buy=0
        kouho2=0
        Yuuyo2=0
        tanngen_suu4=0
        
    return KounyuFlag,kouho,Yuuyo1,tanngen_suu2,Sisan_Sougaku,Kaitsuke_kanougaku,Rank2_buy,kouho2,Yuuyo2,tanngen_suu4

def Chuumon():
    #証券コードと購入価格と数量を渡す
    Kaitsuke,Code,Kabuka1,Suuryo,Sisan,Kaitsuke_kanougaku,Rank2_buy,Rank2,Kabuka2,Suuryo2=Suuryou_keisan()

    #相場の強弱を判断
    path100 = r"C:\Users\katuy\Desktop\株式Python\SuspendTrading_days.txt"
    with open(path100) as f:
        SuspendTrading_days = f.read()
        SuspendTrading_days = int(SuspendTrading_days)

    #print(Kaitsuke,SuspendTrading_days,Sisan,Kaitsuke_kanougaku)
    #買付フラグがオン&取引制限がない場合は、購入処理を実行する。
    if Sisan==Kaitsuke_kanougaku and Kaitsuke ==1 and SuspendTrading_days<=0:

        #以下、注文用のエクセル起動
        start_marketspeed2()
        py.hotkey("win","m")
        #好きな場所にRSSの自動売買の設定をしたエクセルファイルをおいてください。
        wb = xw.Book(r'C:\Users\katuy\Desktop\注文用.xlsm')
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

        #銘柄売買用シートの指定
        sheet = wb.sheets['現物注文']
        sheet.range('A4').value = None
        sheet.range('B4').value = None
        sheet.range('C4').value = None
        sheet.range('D4').value = None

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

        if Rank2_buy == 1:
            kaisuu=2
        else:
            kaisuu=1

        for i in range(0,kaisuu):
            if i <=0:
                Random_ID=random.randint(1,100)
                #発注ID
            if i >=1:
                #2週目があるなら、書き換えて注文
                sheet.range('A4').value = None
                sheet.range('B4').value = None
                sheet.range('C4').value = None
                sheet.range('D4').value = None
                time.sleep(3)
                Random_ID=random.randint(101,200)
                Code=Rank2
                Suuryo=Suuryo2
                Kabuka1=Kabuka2

            sheet.range('A4').value = Random_ID
            #証券コード
            sheet.range('B4').value = Code
            #数量
            sheet.range('C4').value = Suuryo
            #価格
            sheet.range('D4').value = Kabuka1
            time.sleep(5)

            IsUri=0 #0なら買い
            IsNariyuki=10 #エラー対策
            main(Sisan,Kaitsuke,Code,Suuryo,Kabuka1,IsUri,IsNariyuki)

        #アプリ終了
        app = xw.apps.active
        wb.save()
        wb.close()
        app.kill()
        stop_marketspeed2()

        #新しい銘柄を買ったので、次に売るまでの保有日数を3営業日の期間にリセットしておく
        path3 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
        f = open(path3,'w')
        f.write(str(3))#残りの保有日数を3日にセット
        f.close()

    else:
        if (base8 <= now and now < base9):#8時～9時
            send_line_notify("\n <寄付> \n 新規買いは行いません。")
        elif (base9 <= now and now <= base10):#9時～10時
            send_line_notify("\n <朝 9:30> \n 新規買いは行いません。")
        elif (base14 <= now and now <= base15):#14時～15時
            send_line_notify("\n <昼 14:00> \n 新規買いは行いません。")

def main(Sisan,Kaitsuke,Code,Suuryo,Kabuka,IsUri,IsNariyuki):
    #8時～15時までの間なら、売買約定関連のメッセージを送信
    if (base8 <= now and now <= base15):
        if IsUri==0:#買いの場合に実行
            if Kaitsuke ==1:
                if(base8 <= now and now < base9):#8時～9時
                    send_line_notify("\n <寄付>"+str(Code)+"を "+str(Kabuka)+"円以下で、\n "+str(Suuryo)+"株"+"買いました。")
                elif (base9<=now and now<=base10):
                    send_line_notify("\n <朝 9:30>"+str(Code)+"を "+str(Kabuka)+"円以下で、\n "+str(Suuryo)+"株"+"買いました。")
                else:
                    send_line_notify("\n <昼 14:00>"+str(Code)+"を "+str(Kabuka)+"円以下で、\n "+str(Suuryo)+"株"+"買いました。")
        elif IsNariyuki==0:#成り行き売りの場合に実行
            send_line_notify("\n"+str(Code)+"を"+str(Suuryo)+"株"+"、\n 成行で売りました。")
        else:#指値売りの場合に実行
            send_line_notify("\n"+str(Code)+"を "+str(Kabuka)+"円で "+str(Suuryo)+"株"+"、\n 売り注文を出しました。")

def send_line_notify(notification_message):
    #LINEへ通知
    line_notify_token ="QIgTWDInx8qvVQKiCpsSBWBPP7y5hLTFhNmtXnn9AU2"
    line_notify_api = 'https://notify-api.line.me/api/notify'
    headers = {'Authorization': f'Bearer {line_notify_token}'}
    data = {'message': f' {notification_message}'}
    requests.post(line_notify_api, headers = headers, data = data)
#Chuumon()