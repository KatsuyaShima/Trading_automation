#前日の日経平均の終値から、5～14日の範囲内で、最も近い終値で線を結んでいくプログラム_作業停止中(一応、削除禁止)
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
import pandas as pd
import re
from webdriver_manager.chrome import ChromeDriverManager

def senmusubi(back_dates):
    #上場銘柄データの読み込み
    ALL_Stock_code = pd.read_csv('/Users/katuy/Desktop/株式Python/出力csv3/df_filtered.csv', encoding="shift-jis")
    ALL_Stock_code = ALL_Stock_code.loc[:,['コード','取引額(億)']]

    stock_code=[]
    stock_code2=[]
    stock_code_volume=[]
    stock_code2_volume=[]

    df_a=pd.DataFrame()
    for i in range(0,len(ALL_Stock_code)):
        stock_code.append(ALL_Stock_code.at[ALL_Stock_code.index[i],'コード'])
    for i in range(0,len(ALL_Stock_code)):
        stock_code_volume.append(ALL_Stock_code.at[ALL_Stock_code.index[i],'取引額(億)'])

    #時系列データの日数でフィルターする(100日以上のdataがある)
    for j in range(0,len(ALL_Stock_code)):
        df_a=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code[j]}.csv',encoding="shift-jis")
        if len(df_a)>100:
            stock_code2.append(stock_code[j])
            stock_code2_volume.append(stock_code_volume[j])

    #print(stock_code2_volume)
    df=pd.read_csv(r'C:/Users/katuy/Desktop/株式Python/出力csv/N225.csv',encoding="shift-jis")
    if back_dates>0:
        df=df[:-int(back_dates)]#※日数いじり
    #df=pd.read_csv(r'C:/Users/katuy/Desktop/N225_chart.csv',encoding="shift-jis")

    #1.直近の日経平均から、3日以内で終値が近いものがあるか検索。その後、営業日：5～10日(1,2週間程度)以内で最も変動が少ない株価(終値ベース)で線を結び日付を取得する。
    Firstdate_list=[]
    Lastdate_list=[]
    Hendou_list=[]
    #取得する線の本数
    for u in range(0,5):
        df_sub1=pd.DataFrame()
        df_sub2=pd.DataFrame()
        if u < 1:
            last_a=len(df)
        #取得する日足データの間隔
        for i in range(0,1):
            last=last_a-(i+1)
            sub_close=[]
            sub_date=[]
            #線を結ぶ日足の長さ
            for j in range(10,15):
                sub_date.append(str(df.at[df.index[last-j],'datetime'][:10]))
                sub_close.append((abs(df.at[df.index[last-j],'close']-df.at[df.index[last],'close']))/df.at[df.index[last],'close'])
            df_sub1['開始日']=sub_date
            df_sub1['終了日']=df.at[df.index[last],'datetime'][:10]
            df_sub1['差分']=sub_close
            df_sub2=df_sub2.append(df_sub1)

        df_sub2 = df_sub2.sort_values('差分')
        df_sub2 = df_sub2.reset_index(drop=True)
        #print(df_sub2)
        Firstdate_list.append(df_sub2.at[df_sub2.index[0],'開始日'])
        Lastdate_list.append(df_sub2.at[df_sub2.index[0],'終了日'])
        Hendou_list.append(round(df_sub2.at[df_sub2.index[0],'差分']*100,2))
        for k in range(0,len(df)):
            if df_sub2.at[df_sub2.index[0],'開始日']==df.at[df.index[k],'datetime'][:10]:
                last_a = k+1
        
    #print(Firstdate_list)
    #print(Lastdate_list)

    #求めた3つの線で日経の変動率を算出

    df3=pd.DataFrame(data={'開始日':Firstdate_list,'終了日':Lastdate_list,'N225_変動率':Hendou_list})
    df3 = df3.iloc[::-1]
    Firstdate_list.reverse()
    Lastdate_list.reverse()
    stock_code3=[]
    Latest_close1=[]
    Latest_close2=[]
    Latest_close3=[]
    Latest_close4=[]
    Latest_close5=[]
    Sokone_list=[]
    Takane_list=[]
    average_close_5_list=[]
    average_close_20_list=[]
    average_close_75_list=[]
    Fivedays_up_list=[]

    df_stock_Sokone = pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv3/df_SokoneRatio.csv',encoding="shift-jis")
    #同じ期間で個別銘柄の変動率を算出
    for i in range(0,len(stock_code2)):
        df_stock=pd.DataFrame()
        df3_1=pd.DataFrame()
        First_close_list=[]
        End_close_list=[]
        average_close_5=0
        average_close_20=0
        average_close_75=0
        Fivedays_up_ratio=0


        df_stock = pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv/{stock_code2[i]}.csv',encoding="shift-jis")

        if back_dates>0:
            df_stock=df_stock[:-int(back_dates)]#※日数いじり

        for k in range(0,len(df3)):
            for j in range(0,len(df_stock)):
                if df3.at[df.index[k],'開始日']==df_stock.at[df_stock.index[j],'datetime']:
                    First_close_list.append(df_stock.at[df_stock.index[j],'close'])
                if df3.at[df.index[k],'終了日']==df_stock.at[df_stock.index[j],'datetime']:
                    End_close_list.append(df_stock.at[df_stock.index[j],'close'])
        
        #5日平均移動線からの乖離率        
        for r in range(1,6):
            #print(df_stock.at[df_stock.index[len(df_stock)-r],'close'])
            average_close_5=average_close_5+df_stock.at[df_stock.index[len(df_stock)-r],'close']
        average_close_5=average_close_5/5
        average_close_5=round(((df_stock.at[df_stock.index[len(df_stock)-1],'close']/average_close_5)-1)*100,2)
        #print(average_close_5)

        #25日平均移動線からの乖離率        
        for r in range(1,26):
            #print(df_stock.at[df_stock.index[len(df_stock)-r],'close'])
            average_close_20=average_close_20+df_stock.at[df_stock.index[len(df_stock)-r],'close']
        average_close_20=average_close_20/25
        average_close_20=round(((df_stock.at[df_stock.index[len(df_stock)-1],'close']/average_close_20)-1)*100,2)

        #75日平均移動線からの乖離率        
        for r in range(1,76):
            #print(df_stock.at[df_stock.index[len(df_stock)-r],'close'])
            average_close_75=average_close_75+df_stock.at[df_stock.index[len(df_stock)-r],'close']
        average_close_75=average_close_75/75
        average_close_75=round(((df_stock.at[df_stock.index[len(df_stock)-1],'close']/average_close_75)-1)*100,2)

        df_fivedays_min_close=10000000
        #直近5日間での最安値終値からの上昇率
        for r in range(1,6):
            if df_stock.at[df_stock.index[len(df_stock)-r],'close'] < df_fivedays_min_close:
                df_fivedays_min_close=df_stock.at[df_stock.index[len(df_stock)-r],'close']
        Fivedays_up_ratio=round(((df_stock.at[df_stock.index[len(df_stock)-1],'close']/df_fivedays_min_close)-1)*100,2)



        for u in range(0,len(df_stock_Sokone)):
            if df_stock_Sokone.at[df_stock_Sokone.index[u],'コード']==stock_code2[i]:
                Takane_list.append(df_stock_Sokone.at[df_stock_Sokone.index[u],'高値からの変動率'])
                Sokone_list.append(df_stock_Sokone.at[df_stock_Sokone.index[u],'底値からの変動率'])
        
        """print(stock_code2[i])
        print(len(Firstdate_list))
        print(len(Lastdate_list))
        print(len(First_close_list))
        print(len(End_close_list))"""

        First_close_list.reverse()
        End_close_list.reverse()

        Close_ratio=[]
        #print(stock_code2[i])
        df3_1=pd.DataFrame(data={'開始日':Firstdate_list,'終了日':Lastdate_list,'開始日の終値':First_close_list,'終了日の終値':End_close_list})
        for q in range(0,len(df3_1)):
            Close_ratio.append(round((df3_1.at[df3_1.index[q],'終了日の終値']-df3_1.at[df3_1.index[q],'開始日の終値'])/df3_1.at[df3_1.index[q],'開始日の終値']*100,2))
        df3_1['個別変動率']=Close_ratio
        stock_code3.append(stock_code2[i])

        Latest_close1.append(Close_ratio[len(df3_1)-1])
        Latest_close2.append(Close_ratio[len(df3_1)-2])
        Latest_close3.append(Close_ratio[len(df3_1)-3])
        Latest_close4.append(Close_ratio[len(df3_1)-4])
        Latest_close5.append(Close_ratio[len(df3_1)-5])
        average_close_5_list.append(average_close_5)
        average_close_20_list.append(average_close_20)
        average_close_75_list.append(average_close_75)
        Fivedays_up_list.append(Fivedays_up_ratio)

    """print(len(stock_code3))
    print(len(Latest_close1))
    print(len(Latest_close2))
    print(len(Latest_close3))
    print(len(Latest_close4))
    print(len(Latest_close5))
    print(len(Takane_list))
    print(len(Sokone_list))
    print(len(average_close_5_list))
    print(len(average_close_20_list))
    print(len(Fivedays_up_list))
    print(len(stock_code2_volume))"""

    df_sort1=pd.DataFrame(data={'コード':stock_code3,'直近1':Latest_close1,'直近2':Latest_close2,'直近3':Latest_close3,'直近4':Latest_close4,'直近5':Latest_close5,'高値からの変動率':Takane_list,'底値からの変動率':Sokone_list,'5日乖離率':average_close_5_list,'25日乖離率':average_close_20_list,'75日乖離率':average_close_75_list,'直近5日間での上昇率':Fivedays_up_list,'70日平均取引額(億)':stock_code2_volume})
    df_sort1=df_sort1.sort_values('直近1',ascending=False,ignore_index=True)
    for i in range(0,len(df_sort1)):
        if df_sort1.at[df.index[i],'直近1']<0:
            break

    N225_array=[]
    N225_array=Hendou_list
    N225_array.insert(0,'N225')
    Length=len(N225_array)
    N225_array.insert(Length,'0.0')
    N225_array.insert(Length+1,'0.0')
    #print(N225_array)
    df_insert = pd.DataFrame(N225_array).T
    df_insert = df_insert.set_axis(['コード', '直近1', '直近2','直近3', '直近4', '直近5','高値からの変動率','底値からの変動率'], axis=1)
    return df_sort1

def shikihou(back_dates):
    #検索用銘柄のファイルの読み込み
    Shikihou_Stock_code=senmusubi(back_dates)
    Code=[]
    for i in range(0,len(Shikihou_Stock_code)):
        Code.append(Shikihou_Stock_code.at[Shikihou_Stock_code.index[i],'コード'])

    options = Options()
    options.add_argument('--user-agent=hogehoge')
    options.add_argument('--headless')
    #ログイン処理
    driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
    #driver = webdriver.Chrome(r'C:\Users\katuy\Desktop\株式Python\chromedriver',options=options)
 
    #楽天証券のサイトへアクセス
    driver.get('https://www.rakuten-sec.co.jp/ITS/V_ACT_Login.html')
    #待機
    driver.implicitly_wait(10)

    driver.get_cookies()
    elem_id = driver.find_element(By.XPATH,"/html/body/div[2]/div/div[2]/div[1]/div[1]/div/form/div/div[1]/input[1]")
    elem_id.send_keys("ID")
    elem_pass = driver.find_element(By.XPATH,"/html/body/div[2]/div/div[2]/div[1]/div[1]/div/form/div/div[1]/input[2]")
    elem_pass.send_keys("PASSWORD")
    elem_login_btn = driver.find_element(By.XPATH,"/html/body/div[2]/div/div[2]/div[1]/div[1]/div/form/ul/li[1]/button")
    elem_login_btn.click()

    driver.implicitly_wait(10)
    #四季報のページへ移動
    #elem_kokunai_kabushiki = driver.find_element(By.XPATH,"/html/body/header/div[4]/div/ul/li[2]/a/span")
    elem_kokunai_kabushiki = driver.find_element(By.XPATH,"/html/body/header/div/div[5]/div/ul/li[2]/a/span")#_11/22_うまく動作しない
    elem_kokunai_kabushiki.click()

    shikihou=[]
    shikihou2=[]
    shikihou_flag=[]
    kessann_date=[]
    #Code=Code[90:]
    #検索窓(検索する銘柄をファイルから入力)
    for j in range(0,int(len(Code))):
        print(str(Code[j])+" "+str(j+1)+" of "+ str(len(Code)))
        if j <1:
            elem_stock = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr/td/form/div[3]/div[1]/table/tbody/tr[1]/td[1]/nobr/input")
            elem_kensaku_btn = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr/td/form/div[3]/div[1]/table/tbody/tr[1]/td[2]/a/img")
        else:
            elem_stock = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr/td[1]/form[1]/div[1]/div[1]/table/tbody/tr[1]/td[1]/nobr/input")
            elem_kensaku_btn = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr/td[1]/form[1]/div[1]/div[1]/table/tbody/tr[1]/td[2]/a/img")

        elem_stock.clear()
        elem_stock.send_keys(str(Code[j]))
        elem_kensaku_btn.click()
        
        #更新待機
        driver.implicitly_wait(10)

        #四季報へ
        elem_kensaku_btn2 = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr/td[1]/form[2]/div[2]/table[2]/tbody/tr/td[1]/table/tbody/tr/td/div/div/ul/li[2]/a/nobr")
        elem_kensaku_btn2.click()
        driver.implicitly_wait(10)
        
        #iframe内の情報をスクレイピング
        iframe = driver.find_element(By.ID,"J010101-004-1")
        driver.switch_to.frame(iframe)

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        data=soup.text
        data.strip()
        output=data.split('\n')
        output = list(filter(None, output))
        if "この銘柄には四季報データが存在しません。" not in data:
            for k in range(0,len(output)):
                if output[k]== '解説記事':
                    output[k+1] = "".join(output[k+1].split())
                    shikihou.append(output[k+1])
                    p = r'【(.+)】'  # アスタリスクに囲まれている任意の文字（アスタリスクを除く）
                    r = str(re.findall(p, output[k+1]))
                    shikihou_flag.append(r)
                    output[k+2] = "".join(output[k+2].split())
                    shikihou2.append(output[k+2])
            #iframeから出る
            driver.switch_to.default_content()
            driver.implicitly_wait(10)

            #決算日の取得
            elem_kensaku_btn3 = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr/td[1]/form[2]/div[2]/table[2]/tbody/tr/td[1]/table/tbody/tr/td/div/div/ul/li[5]/a/nobr")
            elem_kensaku_btn3.click()

            #iframe内の情報をスクレイピング
            iframe = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr/td[1]/form[2]/div[2]/div[2]/iframe")
            driver.switch_to.frame(iframe)

            elem_kensaku_btn3 = driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[3]/ul/li[2]/a/div")
            elem_kensaku_btn3.click()

            #決算日の体裁を整える
            soup2 = BeautifulSoup(driver.page_source, 'html.parser')
            data2=soup2.text
            data2.strip()
            output2=data2.split('\n')
            output2 = list(filter(None, output2))

            for y in range(0,len(output2)):
                output2[y]=output2[y].strip()
            
            count = 0
            output3 = []
            while count < len(output2):
                if output2[count] != '':
                    output3.append(output2[count])
                count += 1

            #print(output3)

            if "この銘柄には四季報データが存在しません。" not in data:
                for k in range(0,len(output3)):
                    if output3[k]== '適時開示履歴':
                        output3[k+1] = "".join(output3[k+1].split())
                        kessann_date.append(output3[k+1][:-5])
            else:
                kessann_date.append("None")    
        else:
            shikihou_flag.append("None")
            shikihou.append("None")
            shikihou2.append("None")
            kessann_date.append("None")
        
        #print(kessann_date[j])
        #iframeから出る
        driver.switch_to.default_content()
        driver.implicitly_wait(10)

        #次の検索欄へ
        driver.switch_to.default_content()
        driver.implicitly_wait(10)

    
    for o in range(0,len(shikihou_flag)):
        tbl = str.maketrans('\'', ' ', ' []')
        shikihou_flag[o] = shikihou_flag[o].translate(tbl)
    #print("driverを停止中...")
    driver.close()
    #print("書き出し中...")
    df_shikihou=pd.DataFrame()
    
    df_shikihou["コード"]=Code
    df_shikihou["決算日"]=kessann_date
    df_shikihou["四季報フラグ"]=shikihou_flag
    df_shikihou["四季報"]=shikihou
    df_shikihou["四季報2"]=shikihou2

    df_shikihou.to_csv(f'C:/Users/katuy/Desktop/株式Python/出力csv3/df_shikihou.csv',encoding="cp932")
    return df_shikihou,Shikihou_Stock_code

def Gattai(back_dates):
    kessan_list=[]
    shikihou_list=[]
    shikihou2_list=[]
    shikihou_flag=[]
    #エクセルに出力する前に四季報と繋げる
    df_shikihou,df_Code=shikihou(back_dates)

    for i in range(0,len(df_shikihou)):
        kessan_list.append(df_shikihou.at[df_shikihou.index[i],'決算日'])
        shikihou_list.append(df_shikihou.at[df_shikihou.index[i],'四季報'])
        shikihou2_list.append(df_shikihou.at[df_shikihou.index[i],'四季報2'])
        shikihou_flag.append(df_shikihou.at[df_shikihou.index[i],'四季報フラグ'])
    df_Code["決算日"]=kessan_list
    df_Code["四季報フラグ"]=shikihou_flag
    df_Code["四季報"]=shikihou_list
    df_Code["四季報2"]=shikihou2_list
    df_Code = df_Code.sort_values('高値からの変動率',ignore_index=True)

    df_Code.to_csv(f'C:/Users/katuy/Desktop/株式Python/df_shikihou.csv',encoding="cp932",index=False)
    df_Code.to_excel(f'C:/Users/katuy/Desktop/株式Python/df_shikihou.xlsx',encoding="cp932",index=False)
    return df_Code

def GoodStock(Minusrate,back_dates):
    df_Shikihou_jundre=pd.read_csv(f'/Users/katuy/Desktop/株式Python/出力csv3/df_shikihou_junre_better.csv',encoding="cp932")
    df_Shikihou_jundre=df_Shikihou_jundre.drop("Unnamed: 0",axis=1)
    
    cols = ['コード', '直近1', '直近2','直近3', '直近4', '直近5','高値からの変動率','底値からの変動率']
    df_good_stock=pd.DataFrame(columns=cols)
    df_GoodStock = Gattai(back_dates)
    #四季報の条件に合う銘柄のみを新しい配列に追加(条件を縛りすぎると見落とすので、各条件に遊びを設けている)
    for i in range(0,len(df_GoodStock)):
        #for k in range(0,len(df_Shikihou_jundre)):
        if (#良い四季報フラグを参照して、一致している銘柄のみ追加_過去の四季報を見ると、良い評価は関係ない
            #df_GoodStock.at[df_GoodStock.index[i],'四季報フラグ']==df_Shikihou_jundre.at[df_Shikihou_jundre.index[k],'四季報フラグ'] 
            df_GoodStock.at[df_GoodStock.index[i],'直近1']>-20 #0%以上が望ましい(直近が日経平均よりも強い値動き)
        and df_GoodStock.at[df_GoodStock.index[i],'底値からの変動率']<50 #10%未満が望ましい(あまり下で買っている人がいないので売られにくい)
        and df_GoodStock.at[df_GoodStock.index[i],'高値からの変動率']<Minusrate #-85%以下が望ましい(大きく下がっているので、爆発しやすい)
        #以下、継続懸念銘柄の除外(倒産リスクがある銘柄は買いが入りにくい為)
        and not any(map(df_GoodStock.at[df_GoodStock.index[i],"四季報"].__contains__, ("疑義注記","重要事象")))
        and not any(map(df_GoodStock.at[df_GoodStock.index[i],"四季報2"].__contains__, ("疑義注記","重要事象")))):
            df_good_stock=df_good_stock.append(df_GoodStock.iloc[i])
    df_good_stock=df_good_stock.reset_index()
    df_good_stock=df_good_stock.drop("index",axis=1)
    #print(df_good_stock)
    df_good_stock.to_csv(f'C:/Users/katuy/Desktop/株式Python/df_good_stock.csv',encoding="cp932",index=False)
    df_good_stock.to_excel(f'C:/Users/katuy/Desktop/株式Python/df_good_stock.xlsx',encoding="cp932",index=False)
#GoodStock(-10,0)