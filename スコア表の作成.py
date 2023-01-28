#銘柄別の点数化&購入銘柄の決定
import pandas as pd
import datetime,requests,os,math,numpy
import 楽天RSS資産確認_自動注文

today=datetime.datetime.today().date()

def 前回決算からの経過日数計算(back_dates):
    df=pd.read_csv(f'C:/Users/katuy/Desktop/株式Python/df_good_stock.csv',encoding="cp932")

    date_list=[]
    date_list=df["決算日"].to_list()

    #日付の型を揃える
    for i in range(0,len(date_list)):
        date_list[i]="20"+date_list[i]
        date_list[i]=date_list[i].replace('/', '-')
        tdatetime = datetime.datetime.strptime(date_list[i], '%Y-%m-%d')
        date_list[i] = datetime.date(tdatetime.year, tdatetime.month, tdatetime.day)

    sabun=[]
    for i in range(0,len(date_list)):
        sabun.append((today-date_list[i]).days)
    df["前回決算からの経過日数"]=sabun
    #print(df)

    cols=["コード","直近1","直近2","直近3","直近4","直近5","高値からの変動率","底値からの変動率","5日乖離率","25日乖離率","75日乖離率","直近5日間での上昇率","70日平均取引額(億)","四季報","四季報2","四季報フラグ","決算日","前回決算からの経過日数"]
    #cols = ['コード', '直近1', '直近2','直近3', '直近4', '直近5','高値からの変動率','底値からの変動率']
    df2=pd.DataFrame(columns=cols)

    days=5
    #直近5日の取引額が低いものは除く
    for i in range(0,len(df)):
        df5 = pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{df.at[df.index[i],"コード"]}.csv',encoding="shift-jis")
        if back_dates>0:
            df5=df5[:-int(back_dates)]#※日数いじり
        amount=0
        for k in range(len(df5)-days,len(df5)):
            amount=amount+df5.at[df5.index[k],'平均売買代金(千万円)']
        if (amount/days)>0 and (0<=df.at[df.index[i],"前回決算からの経過日数"] and df.at[df.index[i],"前回決算からの経過日数"]<=180):#(千万円) & 決算日の直前直後のものは値動きが不適切なので省く 10<=X<=80
            df2=df2.append(df.iloc[i])

    df2=df2.reset_index().drop("index",axis=1)
    df2.to_csv(f'C:/Users/katuy/Desktop/株式Python/df_BUY.csv',encoding="cp932",index=False)
    #print(df2)
    return df2

#各銘柄のスコア決定(配点要素：直近の流れ　弱→強　/ 高値から変動率 / 底値からの変動率 / 5日乖離率 / 70日平均取引額 / 一瞬だけ高値の銘柄は弾く)
def Score(back_dates):
    Takane_ratio_list=[]
    Sokone_ratio_list=[]
    Sokone_ratio_list2=[]
    Fivedays=[]
    Twentydays=[]
    Seventy_five_days=[]
    Fivedays_up_ratio=[]
    Torihiki_volume=[]
    Sokone_min=0
    df2=前回決算からの経過日数計算(back_dates)
    #print(df2)
    cols = ["コード","直近1","直近2","直近3","直近4","直近5","高値からの変動率","底値からの変動率","5日乖離率","25日乖離率","75日乖離率","直近5日間での上昇率","70日平均取引額(億)","四季報","四季報2","四季報フラグ","決算日","前回決算からの経過日数"]
    df_score=pd.DataFrame(columns=cols)

    length=len(df2)
    #print(length)
    #NaNの除外
    for i in range(0,length):
        if (math.isnan(df2.at[df2.index[i],"25日乖離率"])
        or  math.isnan(df2.at[df2.index[i],"5日乖離率"])
        or  math.isnan(df2.at[df2.index[i],"75日乖離率"])
        or  math.isnan(df2.at[df2.index[i],"直近5日間での上昇率"])
        or  math.isnan(df2.at[df2.index[i],"70日平均取引額(億)"])
        or  math.isnan(df2.at[df2.index[i],"直近1"])
        or  math.isnan(df2.at[df2.index[i],"直近2"])
        or  math.isnan(df2.at[df2.index[i],"直近3"])
        or  math.isnan(df2.at[df2.index[i],"直近4"])
        or  math.isnan(df2.at[df2.index[i],"直近5"])):
            df_score=pd.concat([df_score,df2.drop(df2.index[i])],ignore_index=True)
            length-=1
        else:
            df_score=pd.concat([df_score,df2[i:i+1]],ignore_index=True)
            #print(i)
    #print(df_score)
    #高値からの変動率の計算(相対的)
    Ten=45
    for i in range(0,len(df_score)):
        Takane_ratio_list.append(int(round((pow(-df_score.at[df_score.index[i],"高値からの変動率"],3)/pow(-df_score.at[df_score.index[0],"高値からの変動率"],3))*Ten,0)))
    
    #底値変動率のリストを作成
    for i in range(0,len(df_score)):
        Sokone_ratio_list.append(int(df_score.at[df_score.index[i],"底値からの変動率"]))
    #print(Sokone_ratio_list)

    #最小の底値からの変動率を求める
    for i in range(0,len(Sokone_ratio_list)):
        Sokone_ratio_list2.append(Sokone_ratio_list[i])
    length2=len(Sokone_ratio_list2)
   
    c=0
    for i in range(0,length2):
        if Sokone_ratio_list2[i-c]==0:
            del Sokone_ratio_list2[i-c]
            c+=1
    Sokone_min=min(Sokone_ratio_list2)

    #底値からの変動率の計算(相対的)
    Ten=40
    for i in range(0,len(Sokone_ratio_list)):
        if Sokone_ratio_list[i] !=0:
            Sokone_ratio_list[i]=int(pow(1/(Sokone_ratio_list[i]/Sokone_min),0.3)*Ten)#15_0.6
        else:
            Sokone_ratio_list[i]=Ten

    #5日乖離率のリストを作成
    for i in range(0,len(df_score)):
        Fivedays.append(df_score.at[df_score.index[i],"5日乖離率"])
    #print(Fivedays)
    
    #5日乖離率の計算(絶対的)
    Ten=3
    for i in range(0,len(Fivedays)):
        if Fivedays[i]>-3:
            Fivedays[i]=int(-round(Fivedays[i]*Ten,2))
        else:
            Fivedays[i]=9

    #25日乖離率のリストを作成
    for i in range(0,len(df_score)):
        Twentydays.append(df_score.at[df_score.index[i],"25日乖離率"])
    #print(Fivedays)
    
    #25日乖離率の計算(絶対的)
    Ten=4
    for i in range(0,len(Twentydays)):
        #print(i)
        Twentydays[i]=-int(abs(round(Twentydays[i]*Ten,0)))

    #75日乖離率のリストを作成
    for i in range(0,len(df_score)):
        Seventy_five_days.append(df_score.at[df_score.index[i],"75日乖離率"])
    #print(Fivedays)
    
    Ten=3
    up_ratio=3
    #75日乖離率の計算(絶対的)
    for i in range(0,len(Seventy_five_days)):
        if Seventy_five_days[i]>up_ratio:
            Seventy_five_days[i]=int(-round((Seventy_five_days[i]-up_ratio)*Ten,2))
        else:
            Seventy_five_days[i]=0

    #直近5日での上昇率のリストを作成
    for i in range(0,len(df_score)):
        Fivedays_up_ratio.append(df_score.at[df_score.index[i],"直近5日間での上昇率"])
    #print(Fivedays)

    #直近5日での上昇率の計算(絶対的)
    Ten=3
    for i in range(0,len(Fivedays_up_ratio)):
            Fivedays_up_ratio[i]=-int(abs(round(Fivedays_up_ratio[i]*Ten,0)))
    Ten=0

    volume_max=max(df_score["70日平均取引額(億)"])
    Ten=10
    #70日平均取引額(相対的)
    for i in range(0,len(df_score)):
        Torihiki_volume.append(int(round(pow((df_score.at[df_score.index[i],"70日平均取引額(億)"]/volume_max),0.3)*Ten,2)))
    Ten=0

    Sabun1=[]
    Sabun2=[]
    Sabun3=[]
    tyokkin1_Storength=[]

    down_ratio = -3
    #直近1の強さ
    for i in range(0,len(df_score)):
        tyokkin1_Storength.append(int(round(df_score.at[df_score.index[i],"直近1"],0)))
    for i in range(0,len(tyokkin1_Storength)):
        if tyokkin1_Storength[i]<down_ratio:
            tyokkin1_Storength[i]=round(tyokkin1_Storength[i]-down_ratio,2)*3
        else:
            tyokkin1_Storength[i]=0
    #print(tyokkin1_Storength)

    #直近の流れ 弱→強(絶対的)
    for i in range(0,len(df_score)):
        Sabun1.append(round((df_score.at[df_score.index[i],"直近1"])-(df_score.at[df_score.index[i],"直近2"]),2))
        Sabun2.append(round((df_score.at[df_score.index[i],"直近2"])-(df_score.at[df_score.index[i],"直近3"]),2))
        Sabun3.append(round((df_score.at[df_score.index[i],"直近3"])-(df_score.at[df_score.index[i],"直近4"]),2))
    
    df_score2=pd.DataFrame()
    df_score2["コード"]=df_score["コード"]
    df_score2["直近の流れの強さ1"]=Sabun1
    df_score2["直近の流れの強さ2"]=Sabun2
    df_score2["直近の流れの強さ3"]=Sabun3
    
    
    count_list=[]
    #直近の強さが+S%以上回復(最大で3週連続回復で、3Pt×15=45Pt)
    hiritu=0
    for i in range(0,len(df_score2)):
        count=0
        if df_score2.at[df_score2.index[i],"直近の流れの強さ1"]>hiritu:
            count+=3
        if df_score2.at[df_score2.index[i],"直近の流れの強さ2"]>hiritu:
            count+=2
        if df_score2.at[df_score2.index[i],"直近の流れの強さ3"]>hiritu:
            count+=1
        count_list.append(count)
    df_score2["カウント"]=count_list
    #print(df_score2)

    Count_Score=[]
    #カウント数のリストを作成
    for i in range(0,len(df_score2)):
        Count_Score.append(pow(df_score2.at[df_score2.index[i],"カウント"],1))
    #print(Fivedays)
    
    #直近の流れの強さの評価(絶対的)
    Ten=5
    for i in range(0,len(Fivedays)):
            Count_Score[i]=Count_Score[i]*Ten
    Ten=0

    Double=[]
    #高値からの下落率×底値からの上昇率
    for i in range(0,len(df_score)):
        Double.append(int(round(Takane_ratio_list[i]*Sokone_ratio_list[i]/30,2)))
    #print(Double)

    df_final_score=pd.DataFrame()
    df_final_score["コード"]=df_score["コード"]
    df_final_score["Score (高値からの下落率)"]=Takane_ratio_list
    df_final_score["Score (底値からの上昇率)"]=Sokone_ratio_list
    #df_final_score["Score (高値×底値)"]=Double
    df_final_score["Score (5日乖離率)"]=Fivedays
    df_final_score["Score (25日乖離率)"]=Twentydays
    df_final_score["Score (75日乖離率)"]=Seventy_five_days
    df_final_score["Score (直近5日間での上昇率)"]=Fivedays_up_ratio
    df_final_score["Score (流動性)"]=Torihiki_volume
    #df_final_score["Score (直近の流れの強さ)"]=Count_Score
    df_final_score["Score (直近1の強さ)"]=tyokkin1_Storength

    #print(df_final_score)

    Goukei_Score_list=[]
    for i in range(0,len(df_final_score)):
        Goukei_Score=(df_final_score.at[df_final_score.index[i],"Score (高値からの下落率)"]
        +df_final_score.at[df_final_score.index[i],"Score (底値からの上昇率)"]
        #+df_final_score.at[df_final_score.index[i],"Score (高値×底値)"]
        +df_final_score.at[df_final_score.index[i],"Score (5日乖離率)"]
        +df_final_score.at[df_final_score.index[i],"Score (25日乖離率)"]
        +df_final_score.at[df_final_score.index[i],"Score (75日乖離率)"]
        +df_final_score.at[df_final_score.index[i],"Score (直近5日間での上昇率)"]
        +df_final_score.at[df_final_score.index[i],"Score (流動性)"]
        #+df_final_score.at[df_final_score.index[i],"Score (直近の流れの強さ)"]
        +df_final_score.at[df_final_score.index[i],"Score (直近1の強さ)"])
        Goukei_Score_list.append(Goukei_Score)

    Goukei_Score_list_tmp=Goukei_Score_list
    #print(Goukei_Score_list)
    Goukei_max=max(Goukei_Score_list_tmp)
    for i in range(0,len(Goukei_Score_list)):
        Goukei_Score_list[i]=int(round(Goukei_Score_list[i]/Goukei_max*100,1))
    #print(Goukei_Score_list)

    df_final_score["合計点数"]=Goukei_Score_list
    df_final_score=df_final_score.sort_values("合計点数",ascending=False).reset_index().drop("index",axis=1)
    
    df_final_score = df_final_score.sort_values("Score (高値からの下落率)",ascending=False).reset_index().drop("index",axis=1)
    print(df_final_score)

    #高値からの下落率偏差値計算
    takene_hensati=[]
    for i in range(0,len(df_final_score)):
        takene_hensati.append(df_final_score.at[df_final_score.index[i],"Score (高値からの下落率)"])
    np_takane_hensati=numpy.array(takene_hensati)
    mean = numpy.mean(np_takane_hensati)
    std=numpy.std(np_takane_hensati)

    deviation_value=[]
    for i in range(0,len(df_final_score)):
        deviation=(takene_hensati[i]-mean)/std
        deviation_value.append(50+deviation*10)
    #print(takene_hensati)
    #print(deviation_value)

    #高値からの下落率の偏差値が48以下のものは削除
    for y in range(0,len(df_final_score)):
        if deviation_value[y]<55:
            df_final_score=df_final_score[:y]
            break

    df_final_score=df_final_score.reset_index().drop("index",axis=1)

    print(df_final_score)
    df_final_score = df_final_score.sort_values("合計点数",ascending=False).reset_index().drop("index",axis=1)
    
    #70点以下のものは削除(相対評価)
    for y in range(0,len(df_final_score)):
        if df_final_score.at[df_final_score.index[y],"合計点数"]<=70:
            df_final_score=df_final_score[:y]
            break

    df_final_score = df_final_score.reset_index().drop("index",axis=1)

    Goukei_Score_list2=[]
    for i in range(0,len(df_final_score)):
        Goukei_Score=(df_final_score.at[df_final_score.index[i],"Score (高値からの下落率)"]
        +df_final_score.at[df_final_score.index[i],"Score (底値からの上昇率)"]
        #+df_final_score.at[df_final_score.index[i],"Score (高値×底値)"]
        +df_final_score.at[df_final_score.index[i],"Score (5日乖離率)"]
        +df_final_score.at[df_final_score.index[i],"Score (25日乖離率)"]
        +df_final_score.at[df_final_score.index[i],"Score (75日乖離率)"]
        +df_final_score.at[df_final_score.index[i],"Score (直近5日間での上昇率)"]
        +df_final_score.at[df_final_score.index[i],"Score (流動性)"]
        #+df_final_score.at[df_final_score.index[i],"Score (直近の流れの強さ)"]
        +df_final_score.at[df_final_score.index[i],"Score (直近1の強さ)"])
        Goukei_Score_list2.append(Goukei_Score)

    Goukei_Score3=[]
    for i in range(0,len(df_final_score)):
        Goukei_Score3.append(Goukei_Score_list2[i]-df_final_score.at[df_final_score.index[i],"Score (高値からの下落率)"]+45)

    df_final_score["合計点数2"]=Goukei_Score3
    df_final_score = df_final_score.sort_values("合計点数2",ascending=False).reset_index().drop("index",axis=1)

    print(df_final_score)
    #合計点数2が70点以下のものは削除(絶対評価)
    for y in range(0,len(df_final_score)):
        if df_final_score.at[df_final_score.index[y],"合計点数2"]<=70:
            df_final_score=df_final_score[:y]
            break
    print(df_final_score)

    base1 = datetime.time(9, 0, 0)
    base2 = datetime.time(15, 0, 0)
    dt_now = datetime.datetime.now()
    now = dt_now.time()
    #now = datetime.time(8, 30, 0)

    #ザラ場の外で実行した場合はスコア表をLINEで通知
    #if (now <= base1 or now >= base2):
        #main(df_final_score)#ここをコメントアウトで、スコア表だけ表示!!

    df_final_score.to_csv(f'C:/Users/katuy/Desktop/株式Python/df_BUY_FINAL.csv',encoding="cp932",index=False)
    df_final_score.to_excel(f'C:/Users/katuy/Desktop/株式Python/df_BUY_FINAL.xlsx',encoding="cp932",index=False)

def main(df_final):
    KounyuFlag,Sisan_Sougaku,kounyu_tanka,Code_name,Suuryo,Kaituke_kanougaku,Code2,Suuryo2,MeigaraSuu,kounyu_tanka_Rank2=楽天RSS資産確認_自動注文.excel_move()
    
    Hizuke=[]
    Rank1=[]
    Code=0

    if len(df_final)>0:
        Rank1.append(df_final.at[df_final.index[0],"コード"])
    else:
        print("Nothing")

    #保有銘柄名とコードの変換(楽天RSSからは保有銘柄コードが表示されない為)
    df_code_shutoku=pd.read_csv(f'/Users/katuy/Desktop/All_Stock_data_20220930.csv',encoding="shift-jis")
    for i in range(0,len(df_code_shutoku)):
        if df_code_shutoku.at[df_code_shutoku.index[i],"銘柄名"]==Code_name:
            Code=df_code_shutoku.at[df_code_shutoku.index[i],"コード"]

    #日付を取得する為に、N225.csvを参照して取得
    df_Hizuke=pd.read_csv(r'C:\Users\katuy\Desktop\株式Python\出力csv\N225.csv',encoding="cp932")
    Hizuke.append(df_Hizuke.at[df_Hizuke.index[-1],"datetime"])
    df_original_Rank=pd.read_csv(f'C:/Users/katuy/Desktop/株式Python/df_Rank1_list.csv',encoding="cp932")
    df_original_Rank = df_original_Rank.set_axis(["日付","コード"], axis=1)
    
    if len(df_final)>0:
        df_Rank1=pd.DataFrame()
        df_Rank1["日付"]=Hizuke
        df_Rank1["コード"]=Rank1

        df_Rank1 = pd.concat([df_Rank1,df_original_Rank]).reset_index().drop("index",axis=1)

        #同じ日付が存在する時は、追加しない
        if df_Rank1.at[df_Rank1.index[0],"日付"]!=df_original_Rank.at[df_original_Rank.index[0],"日付"]:
            df_Rank1.to_csv(f'C:/Users/katuy/Desktop/株式Python/df_Rank1_list.csv',encoding="cp932",mode="w",index=False)

    else:
        #dataがないので手動で作成
        list1=[]
        list1.append(Hizuke[-1])
        list1.append("----")
        df_insert = pd.DataFrame(list1).T
        df_insert = df_insert.set_axis(["日付","コード"], axis=1)

        df_insert = pd.concat([df_insert,df_original_Rank]).reset_index().drop("index",axis=1)

        #同じ日付が存在する時は、追加しない
        if df_insert.at[df_insert.index[0],"日付"]!=df_original_Rank.at[df_original_Rank.index[0],"日付"]:
            df_insert.to_csv(f'C:/Users/katuy/Desktop/株式Python/df_Rank1_list.csv',encoding="cp932",mode="w",header=False,index=False)

    #保有銘柄がトヨタ自動車でなければ、保有銘柄があるということ
    if Code_name!="トヨタ自動車":
        kounyu_tanka2=int(round(kounyu_tanka,0))
        #保有銘柄の損益率と損益額の計算
        df_code_jikeiretsu=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{Code}.csv',encoding="shift-jis")
        last_index=len(df_code_jikeiretsu)-1
        Latest_close=df_code_jikeiretsu.at[df_code_jikeiretsu.index[last_index],"close"]
        Rieki=round(((Latest_close/kounyu_tanka)-1)*100,2)
        Sonekigaku=int(round(Suuryo*(Latest_close-kounyu_tanka),0))
    Kabushiki=int(Sisan_Sougaku-Kaituke_kanougaku)

    if len(df_final)>0:

        #スコア表をLINE通知
        df2=pd.DataFrame()
        df2["コード"]=df_final["コード"]
        df2["スコア"]=df_final["合計点数"]

    #処理の終了時刻を計算
    t_delta=datetime.timedelta(hours=9)
    JST=datetime.timezone(t_delta,"JST")
    now = datetime.datetime.now(JST)
    d=now.strftime("%Y年%m月%d日 %X")

    send_line_notify("\n"+str(d)+"\n 本日の集計が無事に完了。")
    send_line_notify("\n 資産額："+str(Sisan_Sougaku)+"円"+"\n 株式："+str(Kabushiki)+"円"+"\n 現金："+str(Kaituke_kanougaku)+"円")
    
    if Code_name!="トヨタ自動車" and MeigaraSuu !=0:
        send_line_notify("\n 保有銘柄："+str(Code)+"\n 保有数量："+str(Suuryo)+"株"+"\n 損益率："+str(Rieki)+"%"+"\n 含み損益額："+str(Sonekigaku)+"円"+"\n 現在値："+str(Latest_close)+"円"+"\n 取得単価："+str(kounyu_tanka2)+"円")
        #その日の1位と保有銘柄が一致していれば保有期間を3日にリセット
        if len (df_final)>=1:
            if Code == df_final.at[df_final.index[0],"コード"]:
                #保有銘柄の保有日数の取得
                path0 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
                with open(path0) as f:
                    DaysFromBought0 = f.read()
                DaysFromBought0=int(DaysFromBought0)#現在の保有日数

                #保有予定日数を3日間（初期状態）にリセットする
                path3 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
                f = open(path3,'w')
                f.write(str(3))#残りの保有日数を3日にセット
                f.close()

                #保有銘柄の保有日数の取得
                path4 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
                with open(path4) as f:
                    DaysFromBought = f.read()
                DaysFromBought=int(DaysFromBought)#現在の保有日数
                send_line_notify("\n 保有銘柄のスコアが1位なので、"+"\n 残りの保有営業日数を\n "+str(DaysFromBought0)+"日から"+str(DaysFromBought)+"日に延長しました。")
            #保有銘柄が1位じゃない場合/残りの保有日数を1日減らして、LINE通知
        else:
            #利確・損切が発動せずに、保有日数が3日以上経過する&スコア1位じゃなくなった時に売る
            path5 = r"C:\Users\katuy\Desktop\株式Python\DaysFromBought.txt"
            isempty = os.stat(path5).st_size == 0#ファイルが空か判定

            #ファイルの中身を確認して、保有日数を1日減らす
            if isempty==False:
                with open(path5) as f:
                    DaysFromBought0 = f.read()
                DaysFromBought=int(DaysFromBought0)-1
                
                f = open(path5,'w')
                f.write(str(DaysFromBought))
                f.close()

                if DaysFromBought>=1:
                    send_line_notify("\n 残りの保有営業日数を\n "+str(DaysFromBought0)+"日から"+str(DaysFromBought)+"日に減らしました。")
                else:
                    send_line_notify("\n 残りの保有営業日数を\n "+str(DaysFromBought0)+"日から"+str(DaysFromBought)+"日に減らしました。"+"\n 次の寄付で成行売りします。")
            else:#なにかしらの問題でファイルの中身が空の場合は2にリセットする
                f = open(path5,'w')
                f.write(str(3))#保有日数は3日を最小にする
                f.close()
            
        #保有銘柄数が2つの場合
        if MeigaraSuu>=2:
            df_code_jikeiretsu2=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{Code2}.csv',encoding="shift-jis")
            last_index2=len(df_code_jikeiretsu2)-1
            Latest_close2=df_code_jikeiretsu2.at[df_code_jikeiretsu2.index[last_index2],"close"]
            Rieki2=round(((Latest_close2/kounyu_tanka_Rank2)-1)*100,2)
            Sonekigaku2=int(round(Suuryo2*(Latest_close2-kounyu_tanka_Rank2),0))
            send_line_notify("\n ※サブ"+"\n 保有銘柄："+str(Code2)+"\n 保有数量："+str(Suuryo2)+"株"+"\n 損益率："+str(Rieki2)+"%"+"\n 含み損益額："+str(Sonekigaku2)+"円"+"\n 現在値："+str(Latest_close2)+"円"+"\n 取得単価："+str(kounyu_tanka_Rank2)+"円")
    else:
        #保有銘柄がないので、前回保有銘柄を9999(存在しない証券コード)にする：これで前回持っていた銘柄がランク1位なら、1日開けて買うことが出来るようにする
        path3 = r"C:\Users\katuy\Desktop\株式Python\Zenkai_Hoyuu.txt"
        f = open(path3,'w')
        f.write("9999")
        f.close()
        send_line_notify("\n 現在の保有銘柄はありません。")

    if len(df_final)>0:
        send_line_notify("\n 最新のスコア表を共有します↓")
        send_line_notify("\n "+str(df2))
    else:
        send_line_notify("\n 本日のスコア表はありません。\n ※及第点以上の銘柄がないです。")

    #毎日ファイルに日付と資産額を書き込み
    isempty = os.stat(r"C:\Users\katuy\Desktop\株式Python\資産推移.csv").st_size == 0#ファイルが空か判定

    day=datetime.datetime.now()
    date=[day.date()]
    Sisan_Sougaku1=[Sisan_Sougaku]

    cols=["日付","資産額"]
    df=pd.DataFrame(columns=cols)
    df["日付"]=date
    df["資産額"]=Sisan_Sougaku1
    
    if isempty==False:
        df0=pd.read_csv(r"C:\Users\katuy\Desktop\株式Python\資産推移.csv",encoding="utf-8")
        last_index=len(df0)-1
        if df0.at[df0.index[last_index],"日付"]!=str(day.date()):
            df.to_csv(r"C:\Users\katuy\Desktop\株式Python\資産推移.csv", mode='a',header=False,index=False)
    else:
        df.to_csv(r"C:\Users\katuy\Desktop\株式Python\資産推移.csv", mode='a',index=False)

def send_line_notify(notification_message):
    #LINEへ通知
    line_notify_token ="ご自身のトークンを入力してください"
    line_notify_api = 'https://notify-api.line.me/api/notify'
    headers = {'Authorization': f'Bearer {line_notify_token}'}
    data = {'message': f' {notification_message}'}
    requests.post(line_notify_api, headers = headers, data = data)
#Score(0)