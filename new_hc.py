#このプログラムは店別品番別実績を自動ダウンロードを行う

#----------------------------------------------------------------------------------------------


import time
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ChromeOptions
import datetime
import os
import glob
import shutil
from operator import itemgetter
import tes
import datetime
import pandas as pd
import re
import openpyxl as pyxl

#このプログラムは店別品番別実績を自動ダウンロードを行う



#ーーーーーーーーーー前回データの削除ーーーーーーーーーーーーー
folders = [0,1,2,3,4,5,6]
no = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]

w_day = datetime.datetime.today()

wd_no = w_day.weekday()#曜日Noを指定

main_dr = 'C:/Users/fun-f/Desktop/myfile'

print(wd_no)

to_file_path = str(main_dr) + '/' + str(wd_no)#drpathの指定

  #ーーーーーー曜日別商品実績ファイルクリアーーーーーーーーーー
  
if wd_no == 0:# 月曜日⇒0 火曜日⇒ 1 水曜日⇒ 2 木曜日⇒ 3 金曜日⇒ 4 土曜日⇒ 5 日曜日⇒ 6
  cl_sheet = pd.read_excel('C:/Users/fun-f/Desktop/myfile/クリアBOOK.xlsx')

  cl_df =pd.DataFrame(cl_sheet)
  for fd in folders:
    print(fd)
    for i in no:
      
      del_path = 'C:/Users/fun-f/Desktop/myfile/'+str(fd)+'/'+str(i)+'商品実績.xlsx'
      print(del_path)
      cl_df.to_excel(del_path)
      
  #ーーーーーーーーー実績ファイルクリアーーーーーーーーーーーーー
  
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/0/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/1/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/2/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/3/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/4/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/5/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/6/実績/実績.xlsx')

  print('success')#削除完了
  
else:  
  print('Non Success!!')#削除ファイルなし

#ーーーーーーーー前回ダウンロードファイル削除ーーーーーーーーーー
dr_files = 'C:/Users/fun-f/Desktop/myfile/dataf'
dr_read = os.listdir(dr_files)

print(dr_read)

for file_name in dr_read:
  del_f_path = dr_files + '/' + file_name#削除ファイルパスの設定
  os.remove(del_f_path)#dataf内のファイルの削除
  
  
dr_files_2 = 'C:/Users/fun-f/Desktop/myfile/売上実績'  

dr_read_2 = os.listdir(dr_files_2)

print(dr_read_2)

for file_name_2 in dr_read_2:
  if file_name_2.endswith('.csv'):
    del_f_path2 = dr_files_2 + '/' + file_name_2#削除ファイルパスの設定
    os.remove(del_f_path2)#dataf内のファイルの削除
  
#ーーーーーーー今日の日付設定ーーーーーーーーー

fold = 'C:/Users/fun-f/Downloads'


todaytime = datetime.date.today()
tod = '{0:20%y%m%d}'.format(todaytime)#今日の日付(西暦)


#ーーーーーーー販売NETスクレイピングーーーーーーーーーーー

url = 'http://tri.hanbai-net.com/system/Login.aspx'
#driver = webdriver.Chrome('C:/Users/fun-f/Downloads/chromedriver.exe')#旧
#driver = webdriver.Chrome('C:/Users/fun-f/Desktop/myfile/chromedriver.exe')
driver = webdriver.Chrome("C:/Users/fun-f/chromedriver.exe")#2021 0724

driver.get(url)

id_1 = 'tenpo'
id_2 = 'tenpo'

loginid_1 = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtUserCode"]')
loginid_2 = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtPassword"]')

loginid_1.send_keys(id_1)#ユーザーIDを入力
loginid_2.send_keys(id_2)#パスワードを入力



driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnLogin"]').click() 
#ログインボタンをクリック

driver.get('http://tri.hanbai-net.com/system/00000000.aspx')

driver.find_element_by_xpath('//*[@id="Menu1"]/ul/li[7]').click()

driver.find_element_by_xpath('//*[@id="Menu1:submenu:57"]/li[9]/a').click()
'//*[@id="Menu1:submenu:58"]/li[9]/a'#変更前

driver.get('http://tri.hanbai-net.com/system/30021901.aspx?id=010199')


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtCond02"]').clear()#日付クリア

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtCond02"]').send_keys(str(tod))#日付入力

#----------全店------------

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

time.sleep(5)

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '全店.csv')
                    shutil.move('全店.csv','C:/Users/fun-f/Desktop/myfile/dataf')                        
time.sleep(1)                    

#--------柏---------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[3]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[3]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '柏.csv')
                    shutil.move('柏.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------千葉-----------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[4]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[4]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '千葉.csv')
                    shutil.move('千葉.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(2)
#----------横浜------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[5]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[5]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力


time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '横浜.csv')
                    shutil.move('横浜.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------伊勢崎------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[9]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[9]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '伊勢崎.csv')
                    shutil.move('伊勢崎.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------岐阜------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[10]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[10]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '岐阜.csv')
                    shutil.move('岐阜.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)                    
#----------長町------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[11]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[11]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '長町.csv')
                    shutil.move('長町.csv','C:/Users/fun-f/Desktop/myfile/dataf')                    
time.sleep(1)
#----------船橋------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[12]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[12]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '船橋.csv')
                    shutil.move('船橋.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------富士見------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[13]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[13]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '富士見.csv')
                    shutil.move('富士見.csv','C:/Users/fun-f/Desktop/myfile/dataf')                    
time.sleep(1)
#----------レイク------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[15]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[15]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], 'レイク.csv')
                    shutil.move('レイク.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)                    
#----------海老名------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[17]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[17]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '海老名.csv')
                    shutil.move('海老名.csv','C:/Users/fun-f/Desktop/myfile/dataf')  
time.sleep(1)                                      
#----------むさし------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[18]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[18]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], 'むさし.csv')
                    shutil.move('むさし.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------平塚------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[19]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[19]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '平塚.csv')
                    shutil.move('平塚.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------名取------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[20]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[20]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '名取.csv')
                    shutil.move('名取.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------大高------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[21]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[21]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '大高.csv')
                    shutil.move('大高.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------東郷町------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[22]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[22]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '東郷町.csv')
                    shutil.move('東郷町.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------太田------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[23]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[23]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '太田.csv')
                    shutil.move('太田.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)                    
#----------水戸------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[24]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[24]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '水戸.csv')
                    shutil.move('水戸.csv','C:/Users/fun-f/Desktop/myfile/dataf')                    
time.sleep(1)
#----------EXPO------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[25]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[25]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], 'EXPO.csv')
                    shutil.move('EXPO.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)
#----------川崎------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[26]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[26]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '川崎.csv')
                    shutil.move('川崎.csv','C:/Users/fun-f/Desktop/myfile/dataf')
time.sleep(1)                    
#----------新三郷------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[27]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[27]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '新三郷.csv')
                    shutil.move('新三郷.csv','C:/Users/fun-f/Desktop/myfile/dataf')    
time.sleep(1)  

#----------幕張------------


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]').click()#店舗名指定上段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[28]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]').click()#店舗名指定下段

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[28]').click()#店舗選択


driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '品番売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '幕張.csv')
                    shutil.move('幕張.csv','C:/Users/fun-f/Desktop/myfile/dataf')    
time.sleep(1)                    
        
print("SUCCESS!!") 
                  
        
print("SUCCESS!!") 


driver.find_element_by_xpath('//*[@id="Menu1"]/ul/li[7]').click()

driver.find_element_by_xpath('//*[@id="Menu1:submenu:57"]/li[14]/a').click()
'//*[@id="Menu1:submenu:58"]/li[14]/a'#変更前


driver.get('http://tri.hanbai-net.com/system/30026401.aspx?id=010199')#売上集計＊

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()

time.sleep(1)

driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()#CSV出力

time.sleep(3)#一時待機

filelists = []
for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '売上集計':
            filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename(file[0], '売上実績.csv')
                    shutil.move('売上実績.csv','C:/Users/fun-f/Desktop/myfile/売上実績')
                    
print("SUCCESS!!")      
                               
driver.close()

#ーーーーーショッパー抜きのP率実績ーーーーーーーー

import os
import shutil
import datetime
import requests
import schedule
import pyautogui
import time
import path


#店舗リスト・パス

kasiwa = 'C:/Users/fun-f/Desktop/myfile/dataf/柏.csv'
tiba = 'C:/Users/fun-f/Desktop/myfile/dataf/千葉.csv'
yokohama = 'C:/Users/fun-f/Desktop/myfile/dataf/横浜.csv'
isesaki = 'C:/Users/fun-f/Desktop/myfile/dataf/伊勢崎.csv'
gihu = 'C:/Users/fun-f/Desktop/myfile/dataf/岐阜.csv'
nagamachi = 'C:/Users/fun-f/Desktop/myfile/dataf/長町.csv'
hunabasi = 'C:/Users/fun-f/Desktop/myfile/dataf/船橋.csv'
hujimi = 'C:/Users/fun-f/Desktop/myfile/dataf/富士見.csv'
reiku = 'C:/Users/fun-f/Desktop/myfile/dataf/レイク.csv'
ebina = 'C:/Users/fun-f/Desktop/myfile/dataf/海老名.csv'
musasi = 'C:/Users/fun-f/Desktop/myfile/dataf/むさし.csv'
hiratuka = 'C:/Users/fun-f/Desktop/myfile/dataf/平塚.csv'
natori = 'C:/Users/fun-f/Desktop/myfile/dataf/名取.csv'
otaka = 'C:/Users/fun-f/Desktop/myfile/dataf/大高.csv'
togocyo = 'C:/Users/fun-f/Desktop/myfile/dataf/東郷町.csv'
ota = 'C:/Users/fun-f/Desktop/myfile/dataf/太田.csv'
mito = 'C:/Users/fun-f/Desktop/myfile/dataf/水戸.csv'
expo = 'C:/Users/fun-f/Desktop/myfile/dataf/EXPO.csv'
kawasaki = 'C:/Users/fun-f/Desktop/myfile/dataf/川崎.csv'
sinmisato = 'C:/Users/fun-f/Desktop/myfile/dataf/新三郷.csv'
makuhari = 'C:/Users/fun-f/Desktop/myfile/dataf/幕張.csv'
all_sp = 'C:/Users/fun-f/Desktop/myfile/dataf/全店.csv'

no = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]


#no = [1]

shops_d = {1:"柏",2:"千葉",3:"横浜",4:"伊勢崎",5:"岐阜",6:"長町",7:"船橋",8:"富士見",9:"レイクタウン",10:"海老名",11:"むさし村山",12:"平塚",13:"名取",14:"大高",15:"東郷町",16:"太田",17:"水戸",18:"EXPO",19:"川崎",20:"新三郷",21:"幕張新都心"}

#shops_d = {1:kasiwa,2:tiba,3:yokohama,4:isesaki,5:gihu,6:nagamachi,7:hunabasi,8:hujimi,9:reiku,10:ebina,11:musasi,12:hiratuka,13:natori,14:otaka,15:togocyo,16:ota,17:mito,18:expo,19:kawasaki,20:sinmisato}

shops_l = [kasiwa,tiba,yokohama,isesaki,gihu,nagamachi,hunabasi,hujimi,reiku,ebina,musasi,hiratuka,natori,otaka,togocyo,ota,mito,expo,kawasaki,sinmisato,makuhari,all_sp]#店舗リスト

output_file = "C:/Users/fun-f/Desktop/myfile/Set率集計.xlsx"


#ーーーー曜日Noとto_file_pathの設定ーーーーーーーーー

w_day = datetime.datetime.today()

wd_no = w_day.weekday()#曜日Noを指定

main_dr = 'C:/Users/fun-f/Desktop/myfile'

to_file_path = str(main_dr) + '/' + str(wd_no)#drpathの指定

print(to_file_path)


#-----柏------

ln = 0 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'1商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'1商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[0],encoding='SHIFT-JIS')
#ーーーーーーショッパー別ーーーーーーー

  #sp_s = sp[sp['商品名'] == 'ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS']
  #sp_s_v = sp_s['数量４'].values
  #sp_m = sp[sp['商品名'] == 'ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞM']
  #sp_m_v = sp_m['数量４'].values
  #sp_l = sp[sp['商品名'] == 'ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞL']
  #sp_l_v = sp_l['数量４'].values
  #sp_ll = sp[sp['商品名'] == 'ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞLL']
  #sp_ll_v = sp_ll['数量４'].values
  #spttl = sp_s_v + sp_m_v + sp_l_v
  
#ーーーーーーーーーーーーーーーーーーー
sp_s = sp[sp['部門コード'] == 99]#Pショッパー除外点数
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv = round(((ttl-spttl)/noc),2)
pv_1 = (pv).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'1商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p2_2 = (pv_1[0])
p1 = (p2_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])

#-----千葉------

ln = 1 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'2商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'2商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)
pv2 = round(((ttl-spttl)/noc),2)
pv2_2 = (pv2).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'2商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p2_2 = (pv2_2[0])
p2 = (p2_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])

ln = 0

#-----横浜------

ln = 2 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'3商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'3商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)
pv3 = round(((ttl-spttl)/noc),2)
pv3_3 = (pv3).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'3商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p3_2 = (pv3_3[0])
p3 = (p3_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----伊勢崎------

ln = 3 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'4商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'4商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)
pv4 = round(((ttl-spttl)/noc),2)
pv4_4 = (pv4).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'4商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p4_2 = (pv4_4[0])
p4 = (p4_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----岐阜------

ln = 4 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'5商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'5商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['商品名'] == 'ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS']
sp_s_v = sp_s['合計数量'].values

sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv5 = round(((ttl-spttl)/noc),2)
pv5_5 = (pv5).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'5商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p5_2 = (pv5_5[0])
p5 = (p5_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----長町------

ln = 5 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'6商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'6商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv6 = round(((ttl-spttl)/noc),2)
pv6_6 = (pv6).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'6商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p6_2 = (pv6_6[0])
p6 = (p6_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----船橋------

ln = 6 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'7商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'7商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv7 = round(((ttl-spttl)/noc),2)
pv7_7 = (pv7).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'7商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p7_2 = (pv7_7[0])
p7 = (p7_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----富士見------

ln = 7 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'8商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'8商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv8 = round(((ttl-spttl)/noc),2)
pv8_8 = (pv8).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'8商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p8_2 = (pv8_8[0])
p8 = (p8_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----レイク------

ln = 8 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'9商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'9商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv9 = round(((ttl-spttl)/noc),2)
pv9_9 = (pv9).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'9商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p9_2 = (pv9_9[0])
p9 = (p9_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----海老名------

ln = 9 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'10商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'10商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv10 = round(((ttl-spttl)/noc),2)
pv10_10 = (pv10).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'10商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p10_2 = (pv10_10[0])
p10 = (p10_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----むさし------

ln = 10 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'11商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'11商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv11 = round(((ttl-spttl)/noc),2)
pv11_11 = (pv11).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'11商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p11_2 = (pv11_11[0])
p11 = (p11_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----平塚------

ln = 11 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'12商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'12商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv12 = round(((ttl-spttl)/noc),2)
pv12_12 = (pv12).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'12商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p12_2 = (pv12_12[0])
p12 = (p12_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----名取------

ln = 12 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'13商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'13商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv13 = round(((ttl-spttl)/noc),2)
pv13_13 = (pv13).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'13商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p13_2 = (pv13_13[0])
p13 = (p13_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----大高------

ln = 13 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'14商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'14商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv14 = round(((ttl-spttl)/noc),2)
pv14_14 = (pv14).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'14商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p14_2 = (pv14_14[0])
p14 = (p14_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----東郷町------


ln = 14 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'15商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'15商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv15 = round(((ttl-spttl)/noc),2)
pv15_15 = (pv15).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'15商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p15_2 = (pv15_15[0])
p15 = (p15_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----太田------

ln = 15 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'16商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'16商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv16 = round(((ttl-spttl)/noc),2)
pv16_16 = (pv16).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'16商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p16_2 = (pv16_16[0])
p16 = (p16_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----水戸------

ln = 16 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'17商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'17商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv17 = round(((ttl-spttl)/noc),2)
pv17_17 = (pv17).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'17商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p17_2 = (pv17_17[0])
p17 = (p17_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----EXPO------

ln = 17 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'18商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'18商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)
 
pv18 = round(((ttl-spttl)/noc),2)
pv18_18 = (pv18).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'18商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p18_2 = (pv18_18[0])
p18 = (p18_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----川崎------

ln = 18 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'19商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'19商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)
 
pv19 = round(((ttl-spttl)/noc),2)
pv19_19 = (pv19).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'19商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p19_2 = (pv19_19[0])
p19 = (p19_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----新三郷------

ln = 19 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'20商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'20商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv20 = round(((ttl-spttl)/noc),2)
pv20_20 = (pv20).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'20商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p20_2 = (pv20_20[0])
p20 = (p20_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0


#-----幕張------

ln = 20 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

data_f = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)
#sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
#sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
ttl = data_qyt.sum().values
ttl_amt = data_amt.sum()
ttl_amt_1 = (ttl_amt).values

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)
#sp_s = df_1.filter(like='9998998001',axis=0).values


df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'21商品実績.xlsx')

sp_dt = pd.read_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'21商品実績.xlsx')
sp_df_1 = pd.DataFrame(sp_dt)

#ーーーーーショッパー数量抽出ーーーーーーー

sp = pd.read_csv(shops_l[ln],encoding='SHIFT-JIS')
sp_s = sp[sp['部門コード'] == 99]
sp_s_v = sp_s['合計数量'].sum()#ショッパー合計数

spttl = (sp_s_v)

pv21 = round(((ttl-spttl)/noc),2)
pv21_21 = (pv21).values



df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt,bg,noc,pv], axis=1)#データ結合
df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'21商品実績.xlsx')

print(noc)
print(bg)
print(df_1)
print(ttl)
print(spttl)
print(pv)

bg_2 = (bg).values


mg1 = (str(bg_2[0]) +'/'+ str(ttl_amt_1))

p21_2 = (pv21_21[0])
p21 = (p21_2)


noc_1 = (noc).values
noc_2 = str(noc_1[0])
ln = 0

#-----全店------

data_f = pd.read_csv(shops_l[21],encoding='SHIFT-JIS')
data_f_1 = pd.DataFrame(data_f)
data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values)
data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4])
data_nm = pd.DataFrame(data_f_1['商品名'].values)
data_qyt = pd.DataFrame(data_f_1['合計数量'].values)
data_amt = pd.DataFrame(data_f_1['合計金額'].values)

df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)
print(data_cd)
print(data_nm)
print(data_qyt)
print(data_amt)
print(data_itcd)

df_1.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'0商品実績.xlsx')

df_1.to_excel('C:/Users/fun-f/Desktop/myfile/0商品実績.xlsx')

tenpo_val = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv',encoding='SHIFT-JIS')
tenpo_val.to_excel('C:/Users/fun-f/Desktop/myfile/'+str(wd_no)+'/'+'実績/実績.xlsx')
Set_list = []

for i_no in no :
  
  #print(i_no)
   
  data1 = pd.read_excel('C:/Users/fun-f/Desktop/myfile/' + str(wd_no) + '/' + str(i_no + 1) + '商品実績.xlsx')

  data2 = pd.read_excel('C:/Users/fun-f/Desktop/myfile/' + str(wd_no) + '/実績/実績.xlsx')

  df_data1 = pd.DataFrame(data1)
  #for data in data2 :
    #print(data)

  datalists = []#データを格納

  dellists = []#削除データをっ格納

  for i in df_data1.values :
    if i[2] == 98 :
      dellists.append(i)
      
      
    elif i[2] == 13 : 
      dellists.append(i)
      
    elif i[2] == 14 :   
      dellists.append(i)
      
    else :
      datalists.append(i) 
      

  df_filedata = pd.DataFrame(datalists)#商品リスト

    #print(df_filedata)


  cunt_1 = df_filedata[4].values

  print(cunt_1.sum())
    

  df_data2 = pd.DataFrame(data2)

  df_data2_value = df_data2["売上客数"].values#　客数
  df_data2_value2 = df_data2["客単価"].values# 客単価

  print(df_data2_value)
  print(df_data2_value2)
  print(i_no)

  Set = (cunt_1.sum() / df_data2_value[int(i_no)]) #+ "%" str('{: .2f}'.format(Set)) + "%"
  customer_unit_price = df_data2_value2[int(i_no)]
  customer_unit_price_def = "¥"+ '{: ,}'.format(customer_unit_price)

  Set_ans = str('{: .2f}'.format(Set)) + "%"
  
  shop_name = shops_d[i_no + 1]
  
  wb = pyxl.load_workbook(output_file)
  
  ws = wb.active
  
  ws_act = wb["Sheet1"]
  
  select_cell_A = "A" + str(i_no + 2)
  ws_act[str(select_cell_A)].value = shop_name
  
  select_cell_B = "B" + str(i_no + 2)
  ws_act[str(select_cell_B)].value = Set_ans
  
  select_cell_C = "C" + str(i_no + 2)
  ws_act[str(select_cell_C)].value = customer_unit_price_def
  


  Set_ans = str(shops_d[int(i_no + 1)]) + str('{: .2f}'.format(Set)) + "%"
  
  #print(re.search(r'\d+',Set_ans).group())
  
  Set_list.append(Set_ans)
  
  wb.save(output_file)
  wb.close

Set_list.sort()
#print(Set_list)

file_1 = pd.read_excel(output_file)

df_file_1 = pd.DataFrame(file_1)

p = df_file_1.sort_values("P率",ascending=False)
p_value = p.values

print(p)



#print(df_file_1)
p_1 = p_value[0]
p_2 = p_value[1]
p_3 = p_value[2]
p_4 = p_value[3]
p_5 = p_value[4]
p_6 = p_value[5]
p_7 = p_value[6]
p_8 = p_value[7]
p_9 = p_value[8]
p_10 = p_value[9]
p_11 = p_value[10]
p_12 = p_value[11]
p_13 = p_value[12]
p_14 = p_value[13]
p_15 = p_value[14]
p_16 = p_value[15]
p_17 = p_value[16]
p_18 = p_value[17]
p_19 = p_value[18]
p_20 = p_value[19]
p_21 = p_value[20]


#TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン
TOKEN = 'TNKXBcpEMmK4JAmRPaOyVABkA5GWIJkIHQOnsyfu4MD'#FUNの部屋トークン
api_url = 'https://notify-api.line.me/api/notify'
headers = {'Authorization' : 'Bearer ' + TOKEN}
#message = ('\n'+'柏'+'\n'+'【売上予算/実績】'+'\n' + str(mg1) +'\n' +'【P率】' +str(p1) +'\n'+ '【客数】'+ str(noc_2) +str(p2)+str(p3))
message = ('\n'+'今日のP率(※ショッパー抜き)'+'\n'+'1位' +str(p_1[0])+'\n'+"P率"+str(p_1[1])+" 客単価"+str(p_1[2])
           +'\n'+'\n'+'2位' +str(p_2[0])+'\n'+"P率"+str(p_2[1])+" 客単価"+str(p_2[2])+'\n'+'\n'+'3位' +str(p_3[0])+'\n'+"P率"+str(p_3[1])+" 客単価"+str(p_3[2])+'\n'+'\n'+'4位' +str(p_4[0])+'\n'+"P率"+str(p_4[1])+" 客単価"+str(p_4[2])+'\n'+'\n'+'5位' +str(p_5[0])+'\n'+"P率"+str(p_5[1])+" 客単価"+str(p_5[2])+'\n'+'\n'+'6位' +str(p_6[0])+'\n'+"P率"+str(p_6[1])+" 客単価"+str(p_6[2])+'\n'+'\n'+'7位' +str(p_7[0])+'\n'+"P率"+str(p_7[1])+" 客単価"+str(p_7[2])+'\n'+'\n'+'8位' +str(p_8[0])+'\n'+"P率"+str(p_8[1])+" 客単価"+str(p_8[2])+'\n'+'\n'+'9位' +str(p_9[0])+'\n'+"P率"+str(p_9[1])+" 客単価"+str(p_9[2])+'\n'+'\n'+'10位' +str(p_10[0])+'\n'+"P率"+str(p_10[1])+" 客単価"+str(p_10[2])+'\n'+'\n'+'11位' +str(p_11[0])+'\n'+"P率"+str(p_11[1])+" 客単価"+str(p_11[2])+'\n'+'\n'+'12位' +str(p_12[0])+'\n'+"P率"+str(p_12[1])+" 客単価"+str(p_12[2])+'\n'+'\n'+'13位' +str(p_13[0])+'\n'+"P率"+str(p_13[1])+" 客単価"+str(p_13[2])+'\n'+'\n'+'14位' +str(p_14[0])+'\n'+"P率"+str(p_14[1])+" 客単価"+str(p_14[2])+'\n'+'\n'+'15位' +str(p_15[0])+'\n'+"P率"+str(p_15[1])+" 客単価"+str(p_15[2])+'\n'+'\n'+'16位' +str(p_16[0])+'\n'+"P率"+str(p_16[1])+" 客単価"+str(p_16[2])+'\n'+'\n'+'17位' +str(p_17[0])+'\n'+"P率"+str(p_17[1])+" 客単価"+str(p_17[2])+'\n'+'\n'+'18位' +str(p_18[0])+'\n'+"P率"+str(p_18[1])+" 客単価"+str(p_18[2])+'\n'+'\n'+'19位' +str(p_19[0])+'\n'+"P率"+str(p_19[1])+" 客単価"+str(p_19[2])+'\n'+'\n'+'20位' +str(p_20[0])+'\n'+"P率"+str(p_20[1])+" 客単価"+str(p_20[2])+'\n'+'\n'+'21位' +str(p_21[0])+'\n'+"P率"+str(p_21[1])+" 客単価"+str(p_21[2])+'\n'+'\n'+'詳細はOneDriveの【シフト管理】売上実績ファイルを参照下さい！'+'\n'+'\n'+'質問や不明点あれば古内までご連絡下さい！'+'\n'+'\n'+'よろしくお願いいたします。')
#(+'\n'+'岐阜'+str(p5)+'\n'+'長町'+str(p6)+'\n'+'船橋'+str(p7)+'\n'+'富士見'+str(p8)+'\n'+'レイク'+str(p9)+'\n'+'海老名')
#(+str(p10)+'\n'+'むさし'+str(p11)+'\n'+'平塚'+str(p12)+'\n'+'名取'+str(p13)+'\n'+'大高'+str(p14)+'\n'+'東郷町'+str(p15)+'\n'+'太田'+str(p16)+'\n'+'水戸'+str(p17)+'\n'+'EXPO'+str(p18)+'\n'+'川崎'+str(p19)+'\n'+'新三郷'+str(p20)+'\n'+'詳細はOneDriveの【シフト管理】売上実績ファイルを参照下さい！')
payload = {'message': message}

requests.post(api_url, headers=headers, params=payload)   
print("SUCCESSFULL!!")



