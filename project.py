import csv
import os
import pandas as pd
import xlrd

dir_path = os.path.dirname(os.path.realpath(__file__))
filename1 = os.path.join(dir_path, '20201203課程.xlsx')
filename2 = os.path.join(dir_path, '20201206 108學年下學期課程.xlsx')
filename3 = os.path.join(dir_path, '條件一覽 - 110學年轉系條件.csv')
filename4 = os.path.join(dir_path, '條件一覽 - 轉系統整.csv')
filename5 = os.path.join(dir_path, '條件一覽 - 109學年雙主條件.csv')
filename6 = os.path.join(dir_path, '條件一覽 - 雙主統整.csv')
filename7 = os.path.join(dir_path, '條件一覽 - 109學年輔系條件.csv')
filename8 = os.path.join(dir_path, '條件一覽 - 輔系統整.csv')


print('請問是否有目標之科系？（ 請輸入 y（是） or n（否）)')
question = input()
while(question != 'y' and question != 'n'):
    print('請重新輸入')
    question = input()
if (question == 'n'):
    print('請先前往此網址，完成心理測驗：https://www.hollandexam.com/hollandQuiz.aspx')
    print('請依照心理測驗，輸入相對應型別的選項（數字)： 1.實做型 2.研究型 3.藝術型 4.社交型 5.企業型 6.常規型')
    hollandtype = input()
    hollandtest={
      '1':['1.工學院','2.規設院','3.電資院'],
      '2':['4.理學院','5.生技院','6.醫學院'],
      '3':['2.規設院','7.文學院'],
      '4':['8.社科院','6.醫學院'],
      '5':['9.管學院','8.社科院'],
      '6':['9.管學院','8.社科院']
       }
    print('這些是你的目標學院：')
    print(hollandtest[hollandtype])
    determination = '1'
    while (determination == '1'):
        print('請輸入想查詢的學院(數字)')
        department1 = input()
        #####
        str_combine=''
        for i in hollandtest[hollandtype]:
            str_combine+=i
        while (department1 not in str_combine):   #####
            print('請重新輸入')
            department1 = input()
        print("你輸入的數字是:{}".format(department1))
        print("是否重新輸入？(是請輸入1,不是請輸入0)")
        determination = input()
    department={
      '1':['1.機械系','2.化工系','3.材料系','4.資源系','5.土木系','6.水利系','7.系統系','8.能源學程','9.航太系','10.工科系','11.環工系','12.測量系','13.醫工系'],
      '2':['14.建築系','15.都計系','16.工設系'],
      '3':['17.電機系','18.資訊系'],
      '4':['19.數學系','20.物理系','21.化學系','22.光電科學系','23.地科系'],
      '5':['24.生科系','25.生技系'],
      '6':['26.護理系','27.醫學系','28.醫技系','29.物治系','30.職治系','31.藥學系','32.牙醫系'],
      '7':['33.中文系','34.外文系','35.台文系','36.歷史系'],
      '8':['37.法律系','38.政治系','39.經濟系','40.心理系'],
      '9':['41.企管系','42.統計系','43.會計系','44.工資系','45.交管系']
      }
    print('這些是此學院的科系：')
    print(department[department1])

elif(question == 'y'):
    determination = '1'
    while (determination == '1'):
        print('請選擇想選之科系學院(數字)')
        print('1.工學院 2.規設院 3.電資院\n4.理學院 5.生技院 6.醫學院\n7.文學院 8.社科院 9.管學院')
        department1 = input()
        while (int(department1) > 9 or int(department1) < 1):
            print('請重新輸入')
            department1 = input()
        print("你輸入的數字是:{}".format(department1))
        print("是否重新輸入？(是請輸入1,不是請輸入0)")
        determination = input()

    department = {
       '1': ['1.機械系', '2.化工系', '3.材料系', '4.資源系', '5.土木系', '6.水利系', '7.系統系', '8.能源學程', '9.航太系', '10.工科系', '11.環工系',
             '12.測量系', '13.醫工系'],
       '2': ['14.建築系', '15.都計系', '16.工設系'],
       '3': ['17.電機系', '18.資訊系'],
       '4': ['19.數學系', '20.物理系', '21.化學系', '22.光電科學系', '23.地科系'],
       '5': ['24.生科系', '25.生技系'],
       '6': ['26.護理系', '27.醫學系', '28.醫技系', '29.物治系', '30.職治系', '31.藥學系', '32.牙醫系'],
       '7': ['33.中文系', '34.外文系', '35.台文系', '36.歷史系'],
       '8': ['37.法律系', '38.政治系', '39.經濟系', '40.心理系'],
       '9': ['41.企管系', '42.統計系', '43.會計系', '44.工資系', '45.交管系']
    }
    print(department[department1])

Department = {
    '1':'機械系','2':'化工系','3':'材料系','4':'資源系','5':'土木系','6':'水利系','7':'系統系','8':'能源學程','9':'航太系','10':'工科系','11':'環工系',
    '12':'測量系','13':'醫工系','14':'建築系','15':'都計系','16':'工設系','17':'電機系','18':'資訊系','19':'數學系','20':'物理系','21':'化學系','22':'光電科學系',
    '23':'地科系','24':'生科系','25':'生技系','26':'護理系','27':'醫學系','28':'醫技系','29':'物治系','30':'職治系','31':'藥學系','32':'牙醫系','33':'中文系',
    '34':'外文系','35':'台文系','36':'歷史系','37':'法律系','38':'政治系','39':'經濟系','40':'心理系','41':'企管系','42':'統計系','43':'會計系','44':'工資系',
    '45':'交管系'
}
Way = {'1':'轉系','2':'雙主修','3':'輔系','4':'自由選修'}
determination = '1'
while (determination == '1'):
        print('請輸入想選之科系(數字)')
        Department1 = input()
        while (int(Department1) > 45 or int(Department1) < 1):
            print('請重新輸入')
            Department1 = input()
        print("你輸入的數字是:{}".format(Department1))
        print("是否重新輸入？(是請輸入1,不是請輸入0)")
        determination = input()
determination = '1'
while (determination == '1'):
        print('請輸入想選擇之辦法(數字)')
        print('1.轉系 2.雙主修 3.輔系 4.自由選修')
        way = input()
        while (int(way) > 4 or int(way) < 1):
            print('請重新輸入')
            way = input()
        print("你輸入的數字是:{}".format(way))
        print("是否重新輸入？(是請輸入1,不是請輸入0)")
        determination = input()

print('你輸入的科系是: {} ，辦法是: {} '.format(Department[Department1],Way[way]))
details = [0]*12
if way == '1':
    with open(filename3, encoding='utf-8') as csvfile:
        conditions = csv.DictReader(csvfile)
        for row in conditions:
            if Department[Department1] in row['院 系 別']:
                details[0] = (row['轉 系 條 件'])
                details[1] = (row['可轉入名額\n一般生'])
                details[2] = (row['可轉入名額\n僑生'])
                break
            elif Department[Department1][0] in row['院 系 別'] and Department[Department1][1] in row['院 系 別']:
                details[0] = (row['轉 系 條 件'])
                details[1] = (row['可轉入名額\n一般生'])
                details[2] = (row['可轉入名額\n僑生'])

    print(
    '''
{0}({1}):
轉系條件:
{2}
--------------------------------------------------------------------
可轉入名額(一般生): {3}
---------------------------------------------------------------------
可轉入名額(僑生): {4}
    '''.format(Department[Department1],Way[way],'。\n\n'.join(details[0].split('。')),details[1],details[2]))


    with open(filename4, encoding='utf-8') as csvfile:
        conditions = csv.DictReader(csvfile)
        for row in conditions:
            if Department[Department1] in row['科系']:
                details[3] = (row['招生'])
                details[4] = (row['年級限制'])
                details[5] = (row['成績要求'])
                details[6] = (row['繳交資料'])
                details[7] = (row['先修課程'])
                details[8] = (row['測驗'])
                details[9] = (row['面試'])
                details[10] = (row['備註'])
                details[11] = (row['心得'])
                break
            elif Department[Department1][0] in row['科系'] and Department[Department1][1] in row['科系']:
                details[3] = (row['招生'])
                details[4] = (row['年級限制'])
                details[5] = (row['成績要求'])
                details[6] = (row['繳交資料'])
                details[7] = (row['先修課程'])
                details[8] = (row['測驗'])
                details[9] = (row['面試'])
                details[10] = (row['備註'])
                details[11] = (row['心得'])
    print(
    '''
招生:{}
--------------------------------------------------------------------
年級限制: {}
---------------------------------------------------------------------
成績要求: {}
---------------------------------------------------------------------
繳交資料: {}
---------------------------------------------------------------------
先修課程: {}
---------------------------------------------------------------------
測驗: {}
---------------------------------------------------------------------
面試: {}
---------------------------------------------------------------------
備註: {}
---------------------------------------------------------------------
心得: {}
    '''.format(details[3],details[4],details[5],details[6],details[7],details[8],details[9],details[10],details[11]))


elif way == '2':
    with open(filename5, encoding='utf-8') as csvfile:
        conditions = csv.DictReader(csvfile)
        for row in conditions:
            if Department[Department1] in row['院系別']:
                details[0] = (row['雙主修條件'])
                break
            elif Department[Department1][0] in row['院系別'] and Department[Department1][1] in row['院系別']:
                details[0] = (row['雙主修條件'])

    print(
    '''
{0}({1}):
雙主修條件:
{2}
    '''.format(Department[Department1],Way[way],'。\n\n'.join(details[0].split('。'))))


    with open(filename6, encoding='utf-8') as csvfile:
        conditions = csv.DictReader(csvfile)
        for row in conditions:
            if Department[Department1] in row['科系']:
                details[1] = (row['招收'])
                details[2] = (row['年級限制'])
                details[3] = (row['成績要求'])
                details[4] = (row['繳交資料'])
                details[5] = (row['先修課程'])
                details[6] = (row['測驗'])
                details[7] = (row['聯絡電話'])
                details[8] = (row['備註'])
                details[9] = (row['課程名稱'])
                details[10] = (row['課程備註'])
                break
            elif Department[Department1][0] in row['科系'] and Department[Department1][1] in row['科系']:
                details[1] = (row['招收'])
                details[2] = (row['年級限制'])
                details[3] = (row['成績要求'])
                details[4] = (row['繳交資料'])
                details[5] = (row['先修課程'])
                details[6] = (row['測驗'])
                details[7] = (row['聯絡電話'])
                details[8] = (row['備註'])
                details[9] = (row['課程名稱'])
                details[10] = (row['課程備註'])
    count = 0
    print(
    '''
招收:{}
--------------------------------------------------------------------
年級限制: {}
---------------------------------------------------------------------
成績要求: {}
---------------------------------------------------------------------
繳交資料: {}
---------------------------------------------------------------------
先修課程: {}
---------------------------------------------------------------------
測驗: {}
---------------------------------------------------------------------
連絡電話: {}
---------------------------------------------------------------------
備註: {}
---------------------------------------------------------------------
課程名稱: 
{}
---------------------------------------------------------------------
課程備註: {}
    '''.format(details[1],details[2],details[3],details[4],details[5],details[6],details[7],details[8],"\n".join((details[9].split(","))),details[10]))

elif way == '3':
    with open(filename7, encoding='utf-8') as csvfile:
        conditions = csv.DictReader(csvfile)
        for row in conditions:
            if Department[Department1] in row['院 系 別']:
                details[0] = (row['輔 系 條 件'])
                break
            elif Department[Department1][0] in row['院 系 別'] and Department[Department1][1] in row['院 系 別']:
                details[0] = (row['輔 系 條 件'])

    print(
    '''
{0}({1}):
輔系條件:
{2}
    '''.format(Department[Department1],Way[way],'。\n\n'.join(details[0].split('。'))))


    with open(filename8, encoding='utf-8') as csvfile:
        conditions = csv.DictReader(csvfile)
        for row in conditions:
            if Department[Department1] in row['科系']:
                details[1] = (row['招收'])
                details[2] = (row['年級限制'])
                details[3] = (row['成績要求'])
                details[4] = (row['繳交資料'])
                details[5] = (row['先修課程'])
                details[6] = (row['測驗'])
                details[7] = (row['聯絡電話'])
                details[8] = (row['備註'])
                details[9] = (row['課程名稱'])
                details[10] = (row['課程備註'])
                break
            elif Department[Department1][0] in row['科系'] and Department[Department1][1] in row['科系']:
                details[1] = (row['招收'])
                details[2] = (row['年級限制'])
                details[3] = (row['成績要求'])
                details[4] = (row['繳交資料'])
                details[5] = (row['先修課程'])
                details[6] = (row['測驗'])
                details[7] = (row['聯絡電話'])
                details[8] = (row['備註'])
                details[9] = (row['課程名稱'])
                details[10] = (row['課程備註'])
    count = 0
    print(
    '''
招收:{}
--------------------------------------------------------------------
年級限制: {}
---------------------------------------------------------------------
成績要求: {}
---------------------------------------------------------------------
繳交資料: {}
---------------------------------------------------------------------
先修課程: {}
---------------------------------------------------------------------
測驗: {}
---------------------------------------------------------------------
連絡電話: {}
---------------------------------------------------------------------
備註: {}
---------------------------------------------------------------------
課程名稱: 
{}
---------------------------------------------------------------------
課程備註: {}
    '''.format(details[1],details[2],details[3],details[4],details[5],details[6],details[7],details[8],"\n".join((details[9].split(","))),details[10]))

elif way == '4':
    print(
    '''
    {0}({1}):


    '''.format(Department[Department1],Way[way]))
else:
    print("method search error")