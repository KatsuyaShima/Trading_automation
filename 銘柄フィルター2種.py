#出来高と下落率でフィルター
import pandas as pd
import datetime,time

#上場銘柄データの読み込み
ALL_Stock_code = pd.read_csv('/Users/katuy/Desktop/All_Stock_data_20220930.csv', encoding="shift-jis")
ALL_Stock_code = ALL_Stock_code.loc[:,['コード']]
#print(ALL_Stock_code)

#出力データ保存用のディレクトリの再生成
"""shutil.rmtree(f'/Users/katuy/Desktop/株式Python/出力csv3')
os.mkdir(f'/Users/katuy/Desktop/株式Python/出力csv3')"""

stock_code0=[]
for i in range(0,len(ALL_Stock_code)):
    stock_code0.append(ALL_Stock_code.at[ALL_Stock_code.index[i],'コード'])
#時系列データの日数でフィルターする(100日以上のdataがある)
stock_code=[]
for i in range(0,len(ALL_Stock_code)):
    df = pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code0[i]}.csv',encoding="shift-jis")
    if len(df)>100:
        stock_code.append(ALL_Stock_code.at[ALL_Stock_code.index[i],'コード'])

base1 = datetime.time(9, 0, 0)
base2 = datetime.time(15, 0, 0)
dt_now = datetime.datetime.now()
now = dt_now.time()
#now = datetime.time(8, 30, 0)

#関数1_過去の指定期間内での銘柄の最高値(終値)の算出
def Compare_All_stock_data(back_dates):
    rate_of_change_max=[]
    rate_of_change_min=[]
    for j in range(0,len(stock_code)):
        #print(stock_code[j])
        df = pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code[j]}.csv',encoding="shift-jis")
        
        #最新日から、何日分のデータを取るか1年52週・平日は約245日
        df=df.tail(245)
        if back_dates>0:
            df=df[:-int(back_dates)]#※日数いじり
        #print(df)

        date=df['datetime']
        last_date=len(date)-1
        #過去15ヵ月の最高値(終値)から、直近終値の変動率を下位順にソート
        for i in range(0,len(date)):
            if i<1:
                max_close=df.at[df.index[i],'close']
            elif df.at[df.index[i],'close']>max_close:
                max_close=df.at[df.index[i],'close']
        #過去15ヵ月の最高値(終値)から、直近終値の変動率を下位順にソート
        for i in range(0,len(date)):
            if i<1:
                min_close=df.at[df.index[i],'close']
            elif df.at[df.index[i],'close']<min_close:
                min_close=df.at[df.index[i],'close']
        #print(max_close)
        #print(df.at[df.index[last_date],'close'])
        rate_of_change_max.append(round(((df.at[df.index[last_date],'close']/max_close)-1)*100,2))
        rate_of_change_min.append(round(((df.at[df.index[last_date],'close']/min_close)-1)*100,2))
    #print(rate_of_change)
    df_rate=pd.DataFrame()
    df_rate['コード']=stock_code
    df_rate['高値からの変動率']=rate_of_change_max
    df_rate['底値からの変動率']=rate_of_change_min
        
    #print(df_tmp)
    if (now <= base1 or base2 <= now):
        df_rate.to_csv(f'C:/Users/katuy/Desktop/株式Python/出力csv3/df_rate.csv',encoding="shift-jis")
    return df_rate

def Sort_ratio(back_dates):
    df_rate=Compare_All_stock_data(back_dates)
    #高値からの変動率で下位順に上から200銘柄取得
    df_rate=df_rate.sort_values('高値からの変動率',ignore_index=True)
    #df_rate=df_rate.drop('Unnamed: 0',axis=1)
    #print(df_rate)
    df_rate2=pd.DataFrame()
    df_rate2=df_rate2.append(df_rate.iloc[:100])

    #下落率上位のリスト
    Geraku_list=[]
    Geraku_list2=[]
    
    for j in range(0,len(df_rate2)):
        Geraku_list.append(int(df_rate2.at[df_rate2.index[j],'コード']))
    #print(Geraku_list)
    
    amount_volume=[]
    #売買代金上位の銘柄に絞る(日数=days)
    days=70
    for w in range(0,len(Geraku_list)):
        df_Geraku = pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{Geraku_list[w]}.csv',encoding="shift-jis")
        if back_dates>0:
            df_Geraku=df_Geraku[:-int(back_dates)]#※日数いじり

        amount=0
        for f in range(len(df_Geraku)-days,len(df_Geraku)):
            amount=amount+df_Geraku.at[df_Geraku.index[f],'平均売買代金(千万円)']
            #print(amount)
        if (amount/days)>=10:#(千万円)
            amount_volume.append(round((amount/days)/10,2))
            Geraku_list2.append(Geraku_list[w])
    #print(Geraku_list2)下落率と出来高でフィルターした銘柄
    df_rate4=pd.DataFrame()
    df_rate4['コード']=Geraku_list2
    df_rate4['取引額(億)']=amount_volume

    #9時～15時のザラ場意外の時間帯にだけ、df_filteredを更新(じゃないと、このデータを参照している楽天RSS時系列データの更新と次のスクリプトの四季報取得でずれてしまう)
    if (now <= base1 or base2 <= now):
        df_rate2.to_csv(f'C:/Users/katuy/Desktop/株式Python/出力csv3/df_SokoneRatio.csv',encoding="shift-jis")
        df_rate4.to_csv(f'C:/Users/katuy/Desktop/株式Python/出力csv3/df_filtered.csv',encoding="shift-jis")
        time.sleep(1)

#Sort_ratio(0)#引数はさかのぼる日数/最新日を基準にするなら0