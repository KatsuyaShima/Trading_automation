#購入候補の銘柄リストの作成プログラム
import sys
from yahoo_finance_api2 import share
from yahoo_finance_api2.exceptions import YahooFinanceError
import pandas as pd
import numpy as np
import time,math
import datetime

#出力データ保存用のディレクトリの再生成
"""shutil.rmtree(f'/Users/katuy/Desktop/株式Python/出力csv')
os.mkdir(f'/Users/katuy/Desktop/株式Python/出力csv')"""

#関数1
def Get_N225_data(a):#aには所得するデータの年数が入る
    time.sleep(1)
    my_share = share.Share('^N225') #取得したい証券コード入力
    symbol_data = None
    try:
        symbol_data = my_share.get_historical(
        share.PERIOD_TYPE_YEAR, a,#取得したい過去の期間指定
        share.FREQUENCY_TYPE_DAY, 1)#取得したい型指定(週足/日足/分足など)
    except YahooFinanceError as e:
        print(e.message)
        sys.exit(1)
    #↓APIから取得してきたデータをdataflameのdfに格納する
    df_N225 = pd.DataFrame(symbol_data)
    #dfにtimestamp列の表記方法を変更した日時をdatetime列として追加する
    df_N225["datetime"] = pd.to_datetime(df_N225.timestamp, unit='ms')
    #dateにdfのdatetime列を代入して日時のデータだけ別途保存する
    date=df_N225["datetime"]

    #N225の小数点以下丸める
    for o in range(0,len(date)):
        df_N225.at[df_N225.index[o],'close']=round(df_N225.at[df_N225.index[o],'close'])

    #計算結果格納用配列
    zenjitu_ratio = [0]
    zenjitu_ratio1=[]
    sougou_ratio = []

    #前日比計算
    for l in range(0, len(date)-1):
        zenjitu_ratio.append(round(((df_N225.at[df_N225.index[l+1],'close']/df_N225.at[df_N225.index[l],'close']-1)*100),2))
    #前日比を累計利益率の計算の為に再計算する（例:値上がり率10%→1.1倍）
    for m in range(0, len(date)):
        zenjitu_ratio1.append((zenjitu_ratio[m]*0.01+1))
    tmp=zenjitu_ratio1[0]
    #累計利益率計算
    for n in range(0,len(date)):
        tmp=np.multiply(tmp,zenjitu_ratio1[n])
        sougou_ratio.append(round((tmp-1)*100,2))

    #各出力結果を時系列データに組み込み
    df_N225["前日比"] = zenjitu_ratio
    df_N225["累計利益率"] = sougou_ratio
    
    #使わない列データを削除
    df_N225 = df_N225.drop(['open','high','low'], axis=1)
    #ローカルに保存
    df_N225.to_csv(f'/Users/katuy/Desktop/N225_chart.csv',encoding="shift-jis")

#関数2
def Get_All_stock_data(a,b):#aには所得するデータの年数が入る/bには削る銘柄数
    #上場銘柄データの読み込み
    ALL_Stock_code = pd.read_csv('/Users/katuy/Desktop/All_Stock_data_20220930.csv', encoding="shift-jis")
    ALL_Stock_code = ALL_Stock_code.loc[:,['コード']]
    ALL_Stock_code = ALL_Stock_code.drop(range(b))
    for i in range(0,len(ALL_Stock_code)):
        time.sleep(1)
        my_share = share.Share(str(ALL_Stock_code.at[ALL_Stock_code.index[i],'コード']) +'.T') #取得したい証券コード入力
        symbol_data = None
        try:
            symbol_data = my_share.get_historical(
            share.PERIOD_TYPE_YEAR, a,#取得したい過去の期間指定
            share.FREQUENCY_TYPE_DAY, 1)#取得したい型指定(週足/日足/分足など)
        except YahooFinanceError as e:
            print(e.message)
            sys.exit(1)
        #↓APIから取得してきたデータをdataflameのdfに格納する
        df = pd.DataFrame(symbol_data)
        #dfにtimestamp列の表記方法を変更した日時をdatetime列として追加する
        if not 'timestamp' in df.columns :
            continue
        df["datetime"] = pd.to_datetime(df.timestamp, unit='ms')
        #dateにdfのdatetime列を代入して日時のデータだけ別途保存する
        date=df["datetime"]
        
        #日付表記変更
        for o in range(0, len(date)):
            df.at[df.index[o],'datetime'] = str(df.at[df.index[o],'datetime'])#df.atでindexと列名である一つの要素を指定している。そしてstrftimeで型を変更。
            #print(type(df.at[df.index[o],'datetime']))
        
        #計算結果格納用配列
        zenjitu_ratio = [0]
        zenjitu_ratio1=[]
        sougou_ratio = []
        Nikkei_ratio = []
        dealAmount1 = []

        if i<1:
            stock_code = []

        stock_code.append(ALL_Stock_code.at[ALL_Stock_code.index[i],'コード'])
        print('Code:'+str(stock_code[i])+'\t'+str(i+1)+' of '+str(len(ALL_Stock_code)))
        #前日比計算
        for l in range(0, len(date)-1):
            zenjitu_ratio.append(round(((df.at[df.index[l+1],'close']/df.at[df.index[l],'close']-1)*100),2))
        #前日比を累計利益率の計算の為に再計算する（例:値上がり率10%→1.1倍）
        for m in range(0, len(date)):
            zenjitu_ratio1.append((zenjitu_ratio[m]*0.01+1))
        tmp=zenjitu_ratio1[0]
        #累計利益率計算
        for n in range(0,len(date)):
            tmp=np.multiply(tmp,zenjitu_ratio1[n])
            sougou_ratio.append(round((tmp-1)*100,2))

        df_Nikkei=pd.read_csv(f'/Users/katuy/Desktop/N225_chart.csv', encoding="shift-jis")
        #日経225との値動きの差
        #df_Nikkei=Get_N225_data(3)#日経225のチャート取得→値動きの強さの比較用

        #日経225の前日比
        for q in range(0,len(date)):
            Nikkei_ratio.append(df_Nikkei.at[df_Nikkei.index[q],'前日比'])
        
        #print(df)
        #平均売買代金計算
        for p in range(0,len(date)):
                if math.isnan(df.at[df.index[p],'volume'])==False:
                    dealAmount = round(df.at[df.index[p],'close']*df.at[df.index[p],'volume']/10000000)#1000万で割る
                else:
                    dealAmount = ''
                dealAmount1.append(dealAmount)

        #各出力結果を時系列データに組み込み
        df['日経前日比'] = Nikkei_ratio
        df["前日比"] = zenjitu_ratio
        df["累計利益率"] = sougou_ratio
        df["平均売買代金(千万円)"] = dealAmount1
        
        #使わない列データを削除
        df = df.drop(['timestamp','open','high','low'], axis=1)
        df.to_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code[i]}.csv',encoding="shift-jis")

        df_shuusei=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code[i]}.csv',encoding="shift-jis")
        #print(df_shuusei)

        for e in range(0,len(df)):
            df_shuusei.at[df.index[e],"datetime"],date_sakujyo=df_shuusei.at[df.index[e],"datetime"].split()
            df_shuusei.at[df.index[e],"datetime"] = df_shuusei.at[df.index[e],"datetime"].replace("-","/")
        #print(df_shuusei)
        df_shuusei.to_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code[i]}.csv',encoding="shift-jis")