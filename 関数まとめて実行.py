#関数まとめて実行/Windowsタスクスケジューラで平日に動くように設定
import YahooF2_GetStockData
import 楽天RSS_時系列データ取得
import 銘柄フィルター2種
import 四季報情報取得
import スコア表の作成
import 楽天RSS資産確認_自動注文,現物注文のキャンセル_楽天RSS,NYダウの強さを計算
import jpholiday,datetime,random,time

#祝日を判定する
today = datetime.date.today()
result = jpholiday.is_holiday(today)
base8= datetime.time(8, 0, 0)
base9 = datetime.time(9, 0, 0)
base14 = datetime.time(14, 0, 0)
base15 = datetime.time(15, 0, 0)
dt_now = datetime.datetime.now()
now = dt_now.time()
#now = datetime.time(8, 30, 0)

path4= r"C:\Users\katuy\Desktop\株式Python\IsJikkoutyu.txt"
with open(path4) as f:
    IsJikkoutyu = f.read()

def Morning_call():
    #8時～9時に実行
    if (base8 < now and now < base9):
        line_kaomoji=["(´艸｀*)","('ω')ノ","(~_~)","(≧▽≦)","(^^♪","(# ﾟДﾟ)","|:3ミ"]
        ramdom_number=random.randint(0,len(line_kaomoji)-1)
        if result:#祝日の場合に実行
            today_name=jpholiday.is_holiday_name(today)
            楽天RSS資産確認_自動注文.send_line_notify("\nXol、おはよう"+str(line_kaomoji[ramdom_number])+"\n今日は"+str(today_name)+"です。"+"\nなので取引はありません。")
        else:
            楽天RSS資産確認_自動注文.send_line_notify("\nXol、おはよう"+str(line_kaomoji[ramdom_number]))

path3 = r"C:\Users\katuy\Desktop\株式Python\Code_count_IsSuccess.txt"
with open(path3) as f:
    Code_count_IsSuccess = f.read()

#本日が祝日でないなら自動取引を実行
if result == False:
    #8時から9時の間にスコア表はわざと更新しない(8:55に前日のデータを確認して、売りの楽天RSS資産確認_自動注文だけ動かす)
    if (now < base8 or base9 < now) and Code_count_IsSuccess == "NO" and IsJikkoutyu =="NO":
        data="Rakuten" #Yahoo：Yahoofi2 / #Rakuten：楽天RSS
        ryaku="ALL" #ALL：データ取得から最後まで/No_data2：Score()のみ

        #<Yahooから株価時系列取得>_ときどき止まる
        if data=="Yahoo" and ryaku=="ALL":
            How_YEARS=3 #取得する時系列データの期間(年数)
            Sakujyo=0#途中で止まった時に再開する用(最後の出力を足し算で入力していく/例：1232+343+299+100...)
            YahooF2_GetStockData.Get_N225_data(How_YEARS)
            YahooF2_GetStockData.Get_All_stock_data(How_YEARS,Sakujyo)

        #<楽天RSSから株価時系列取得>_休日祝日はかなり不安定
        if data=="Rakuten" and ryaku=="ALL":
            print("1 of 5: 楽天時系列データ収集中...")
            楽天RSS_時系列データ取得.excel_move()

        if ryaku=="No_data" or ryaku=="ALL":
            #高値からの下落率・平均取引額でフィルター
            print("2 of 5: フィルター2種作成中...")
            銘柄フィルター2種.Sort_ratio(0)

            #直近の強さを計算 / 疑義注記&重要事象銘柄・四季報の評価が悪い銘柄・決算日の近い銘柄を除外
            print("3 of 5: 四季報データ収集中...")
            四季報情報取得.GoodStock(-10,0)#引数に、高値からの下落率をマイナスで入力

            print("4 of 5: スコア表作成中...")
            #下落率・上昇率・乖離率・取引額・直近の強さの5要素で点数を決定し、上位からソート(四季報情報取得ファイル(購入銘柄1種)：df_BUY_FINAL.csv)
            スコア表の作成.Score(0)
        
        #スコア表を作成したあとの一番最後に実行する(失敗すると、下の処理が重複してしまうため)
        #ザラ場外(8時～15時外)なら、1日1回の売り制限&保有銘柄の保有日数をリセットしておく
        if (now <= base8 or base15 <= now):
            path = r"C:\Users\katuy\Desktop\株式Python\IsAlreadySoldToday.txt"
            f = open(path,'w')
            f.write("NO")
            f.close()

            #一通りうまくいけば、コードのカウントファイルを削除
            path0 = r"C:\Users\katuy\Desktop\株式Python\Code_count.txt"
            f = open(path0,mode="w")
            f.write("")
            f.close()
            #取得成功ファイルの書き換え
            path3 = r"C:\Users\katuy\Desktop\株式Python\Code_count_IsSuccess.txt"
            f = open(path3,mode="w")
            f.write("YES")
            f.close()

else:#祝日(平日のみ)の8時～9時の間に、LINEに休場のメッセージを飛ばす。
    Morning_call()

#本日が祝日でないなら自動取引を実行
if result == False:
    #LINEでおはよう
    Morning_call()
    
    #8時～9時でのみNYダウの前日の強さを取得する
    if (base8 < now and now < base9):
        NYダウの強さを計算.IsStrongDow()

    #8時～15時でのみ売買の注文処理を実行
    if (base8 < now and now < base15) and IsJikkoutyu == "NO":
        print("5 of 5: 注文処理判別中...")
        #点数が1位のものに楽天RSSから注文確定
        楽天RSS資産確認_自動注文.Chuumon()

        #Code_count_IsSuccessをNOに戻す(スコア表を算出して、買うことが出来る状態にする)
        path3 = r"C:\Users\katuy\Desktop\株式Python\Code_count_IsSuccess.txt"
        f = open(path3,mode="w")
        f.write("NO")
        f.close()
        
        #14時～15時でのみ全注文の取消しを実行(大引けで売買はしたくない為)
        if base14 <= now and now <= base15:
            time.sleep(60)#14:00で買い注文するパターンがあるかもしれないの少し待つ(約定までの猶予)
            現物注文のキャンセル_楽天RSS.Cancel_Order()