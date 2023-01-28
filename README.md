# Trading_automation

# リポジトリ名
このソフトは楽天証券の提供サービスであるマーケットスピード II RSSを用いて、株取引の完全自動化を行うものです。
※セキュリティの観点から、私個人のID・PASSWORDは伏せています。ご自身のものをお使いください。

## Dependency
使用言語 Python  
使用ライブラリ  
yahoo-finance-api2 == 0.0.12  
pandas == 1.3.4  
numpy == 1.21.4  
requests == 2.26.0  
selenium == 4.5.0  
beautifulsoup4 == 4.10.0  
webdriver-manager == 3.8.5  
xlwings == 0.28.1  
PyAutoGUI == 0.9.53  
jpholiday == 0.1.8  

標準ライブラリ  
sys,math,time,datetime,os,re,subprocess,random

## Setup
<必要な手続き>  
楽天証券口座の取得  
<インストールするもの>  
マーケットスピード II RSS 
Excel  
<OS>  
Windows10 (MacOSは不可)  

## Usage
1.マーケットスピード II RSSから時系列データを取得する。  
2.銘柄のテクニカル情報から、フィルタリングを行う。  
3.直近の強さを市場のベンチマークと比較する。  
4.楽天証券HPにアクセスし、国内株式>四季報の情報から、条件に合った内容の銘柄を選別する。(楽天証券アカウント必須)  
5.銘柄のスコア表を作成する。  
6.マーケットスピード II RSSにアクセスし、スコア表の上位銘柄を買う。(売買時間指定)  
7.その日の売買結果、資産をLINEで通知。  

## License
This software is released under the MIT License, see LICENSE.

## Authors
作者：Katsuya Shima  

## References
楽天RSSの起動方法を参照  https://www.stockinvestment.blog/?p=946
