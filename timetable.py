# new.py
# programed by 정현아
 
 
from pandas import DataFrame
import pandas as pd
import random as rd
 
### 멱집합을 생성하는 함수###
def powerset(s): 
    return [[s[j] for j in range(len(s)) if (i&(1<<j))] for i in range(1<<len(s))]
 
 
### 엑셀파일을 튜플로 전환하는 함수 ###   
def convert_excel(file_name): 
 
    full = 'data_file/'+file_name+'.xlsx' 
    data = []
 
    df = pd.read_excel(full,encoding = 'euc-kr')
    frame = DataFrame(df)
 
    index = frame.index
 
    for i in index:
        a = df.ix[i].values # 배열로 값이 생성된다.
        line = tuple(a) # 튜플로 바꾼다.
        data.append(line)
 
    return data
 
 
### 엑셀파일의 데이터를 난이도별로 분류하는 프로그램##
def slice_level(s, l): 
 
    d = list(convert_excel(s))
 
    lev = []
    for i in range(len(d)):
        st = d[i][0].split('-')
                   
        if st[1] == l:
                 
          lev.append(d[i])
 
    return lev
 
 
### 크기가 들을 전공 수와 동일한 모든 경우의 수의 전공시간표를 만드는 프로그램 ###
def major(s, num, level): 
 
 
    c = []
    delete = []
    data = []
    left = []
     
    a = slice_level(s, level)
 
    b= list(range(len(a)))  
 
    m = powerset(b)   # 난이도 별로 자른 데이터의 인덱스 값으로 멱집합 생성
 
 
   #들을 전공의 수와 동일한 멱집합의 요소로 데이터를 생성
    for w in range(len(m)):
        if len(m[w]) == num:
     
          c.append(m[w])
 
 
    #생성된 데이터를 토대로 전공시간표를 생성
    for x in range(len(c)):
 
        d=[]
        for y in range(len(c[x])):
 
                 
            k = c[x][y]
 
            d.append(a[k])
            del(k)
                   
        data.append(d)
        del(d)
         
    #수업명이 겹치거나 수업시간이 겹치는 수업의 인덱스로 제거할 데이터를 생성
    for e in range(len(data)):
 
        for t in range(len(data[e])):
             
            st = str(data[e][t][9]).split(',')
            stt = str(data[e][t][11]).split(',')
 
            for o in range(len(data[e])):
 
                if t != o:
 
                     
                    sk = str(data[e][o][9]).split(',')
                    skk = str(data[e][o][11]).split(',')
                     
 
                    for h in range(len(st)):
 
               
                        if (data[e][t][8] == data[e][o][8]) & (st[h] in sk) | (data[e][t][1] == data[e][o][1]):
 
                            if not(e in delete):
                                delete.append(e)                   
                     
                     
                    for ht in range(len(stt)):
 
                        if (data[e][t][10] == data[e][o][10]) & (stt[ht] in skk) :
 
                            if not(e in delete):
                                delete.append(e)
 
 
    # 제거할 데이터의 값을 제외한 시간표를 생성                       
    for v in range(len(data)):
        if not(v in delete):
            dt = []
            dt.extend(data[v])
            left.append(dt)
            del(dt)
         
    return left
 
 
### 생성된 전공시간표와 교양시간표를 합쳐 생성할 수 있는 모든 시간표를 생성###
def Add_liberal(s, num1, level, s1, num2, L1, L2, L3, L4):
     
    lev=[]
 
    c = []
    data_lib = []
    de = []
    data = []
    delete = []
    left = []
 
    data_major = major(s, num1, level)
     
    lev2 = [L1, L2, L3, L4]
 
    # 입력받은 난이도 선택 유무를 토대로 난이도 정보 생성
    for f in range(len(lev2)):
        if lev2[f] == 1:
          lev.append(f+1)
           
 
    # 난이도 정보를 토대로 교양시간표 정보 생성   
    lib = [slice_level(s1, str(lev[i])) for i in range(len(lev))]
 
 
    b = list(range(len(lib)))
    m = powerset(b) #교양시간표의 인덱스 값으로 멱집합 생성
 
 
    #들을 교양의 수와 동일한 멱집합의 요소로 데이터를 생성    
    for j in range(len(m)):
        if len(m[j]) == num2:
            c.append(m[j])
 
    #생성된 데이터를 토대로 교양시간표 생성 
    for q in range(len(c)):
 
        if num2 == 1:
         
            k0 = int(c[q][0])
 
            for t0 in range(len(k0)):
                dt = []
                dt.append(lib[k0][t0])
                data.append(dt)
 
                del(dt)
 
            del(k0)
 
        elif num2 == 2:
         
            k0 = int(c[q][0])
            k1 = int(c[q][1])
           
            for t0 in range(len(lib[k0])):        
             
             
                for t1 in range(len(lib[k1])):
 
                    dt = []
                    dt.append(lib[k0][t0])
                    dt.append(lib[k1][t1])
                    data_lib.append(dt)
             
                    del(dt)          
          
            del(k0)
            del(k1)
 
        elif num2 == 3:
         
            k0 = int(c[q][0])
            k1 = int(c[q][1])
            k2 = int(c[q][2])
           
            for t0 in range(len(lib[k0])):                 
             
                for t1 in range(len(lib[k1])):
 
                    for t2 in range(len(lib[k2])):
 
                        dt = []
                        dt.append(lib[k0][t0])
                        dt.append(lib[k1][t1])
                        dt.append(lib[k2][t2])
                        data_lib.append(dt)
                         
                        del(dt)          
                    
            del(k0)
            del(k1)
            del(k2)
 
        else:
         
            k0 = int(c[q][0])
            k1 = int(c[q][1])
            k2 = int(c[q][2])
            k3 = int(c[q][3])
               
            for t0 in range(len(lib[k0])):                 
                 
                for t1 in range(len(lib[k1])):
 
                    for t2 in range(len(lib[k2])):
 
                        for t3 in range(len(lib[k3])):
 
                            dt = []
                            dt.append(lib[k0][t0])
                            dt.append(lib[k1][t1])
                            dt.append(lib[k2][t2])
                            dt.append(lib[k3][t3])
                            data_lib.append(dt)
                             
                            del(dt)          
              
            del(k0)
            del(k1)
            del(k2)
            del(k3)
 
 
       
    # 생성된 전공시간표와 교양시간표를 합쳐 총 시간표 생성
    for mx in range(len(data_major)):
 
        for lx in range(len(data_lib)):
             
            dt = []
            dt.extend(data_major[mx])
            dt.extend(data_lib[lx])
 
            data.append(dt)
 
            del(dt)
 
    # 총 시간표중 시간이 겹치거나 수업명이 겹치는 수업의 인덱스를 삭제할 데이터에 추가
    for e in range(len(data)):
 
        for t in range(len(data[e])):
             
            st = str(data[e][t][9]).split(',')
            stt = str(data[e][t][11]).split(',')
             
            for o in range(len(data[e])):
 
                if t != o:
               
                    sk = str(data[e][o][9]).split(',')
                    skk = str(data[e][o][11]).split(',')
 
                    for h in range(len(st)):
 
               
                        if (data[e][t][8] == data[e][o][8]) & (st[h] in sk) | (data[e][t][1] == data[e][o][1]):
 
                            if not(e in delete):
                                delete.append(e)                   
                     
                    for ht in range(len(stt)):
 
 
                        if (data[e][t][10] == data[e][o][10]) & (stt[ht] in skk):
 
                            if not(e in delete):
                                delete.append(e)
 
    # 삭제할 데이터에 있는 수업을 제외하고 총 시간표 생성                       
    for v in range(len(data)):
        if not(v in delete):
            dt = []
            dt.extend(data[v])
            left.append(dt)
            del(dt)
                             
 
    return left
 
 
### 총 시간표에서 수업1개를 랜덤으로 추출###
def Sampling(s, num1, level, s1, num2, L1, L2, L3, L4):
     
    data = []
    d = Add_liberal(s, num1, level, s1, num2, L1, L2, L3, L4)
     
    if len(d) > 0:
        sp = rd.sample(range(len(d)), 1)
 
        data = [d[i] for i in sp]
 
    return data
 
 
### 추출된 시간표를 요일별 시간별로 분류 ###       
def Day(s, num1, level, s1, num2, L1, L2, L3, L4):
 
    d = Sampling(s, num1, level, s1, num2, L1, L2, L3, L4)
   
    monday = []
    tuesday = []
    wednesday = []
    thursday = []
    friday = []
 
 
    if len(d) > 0 :       
     
        for x in range(len(d)):
 
            mon = []
            mon0 = []
            mon1 = []
            mon2 = []
            mon3 = []
            mon4 = []
            mon5 = []
            mon6 = []
            mon7 = []
            mon8 = []
            mon9 = []
 
            tue = []
            tue0 = []
            tue1 = []
            tue2 = []
            tue3 = []
            tue4 = []
            tue5 = []
            tue6 = []
            tue7 = []
            tue8 = []
            tue9 = []
 
            wed = []
            wed0 = []
            wed1 = []
            wed2 = []
            wed3 = []
            wed4 = []
            wed5 = []
            wed6 = []
            wed7 = []
            wed8 = []
            wed9 = []
 
            thu = []
            thu0 = []            
            thu1 = []
            thu2 = []
            thu3 = []
            thu4 = []
            thu5 = []
            thu6 = []
            thu7 = []
            thu8 = []
            thu9 = []
 
            fri = []
            fri0 = []
            fri1 = []
            fri2 = []
            fri3 = []
            fri4 = []
            fri5 = []
            fri6 = []
            fri7 = []
            fri8 = []
            fri9 = []
             
            for i in range(len(d[x])):
                 
                st = str(d[x][i][9]).split(',')
                stt = str(d[x][i][11]).split(',')
                 
                cn = d[x][i][1] + "\n(" + d[x][i][7] +")"
                
                if(d[x][i][8] == '월'):
                     
                    for j in st:
                               
                         
                        if j=='0' :
                            mon0.append(cn)
                             
                        elif j =='1':
                            mon1.append(cn)
                             
                        elif j =='2':
                            mon2.append(cn)
 
                        elif j =='3':
                            mon3.append(cn)
 
                        elif j =='4':
                            mon4.append(cn)
 
                        elif j =='5':
                            mon5.append(cn)
 
                        elif j =='6':
                            mon6.append(cn)
 
                        elif j =='7':
                            mon7.append(cn)
 
                        elif j =='8':
                            mon8.append(cn)
 
                        elif j =='9':
                            mon9.append(cn)
                  
                             
                elif(d[x][i][8] == '화'):
 
                    for j in st:                   
 
                         
                        if j=='0' :
                            tue0.append(cn)
                             
                        elif j =='1':
                            tue1.append(cn)
                             
                        elif j =='2':
                            tue2.append(cn)
 
                        elif j =='3':
                            tue3.append(cn)
 
                        elif j =='4':
                            tue4.append(cn)
 
                        elif j =='5':
                            tue5.append(cn)
 
                        elif j =='6':
                            tue6.append(cn)
 
                        elif j =='7':
                            tue7.append(cn)
 
                        elif j =='8':
                            tue8.append(cn)
 
                        elif j =='9':
                            tue9.append(cn)
 
 
 
                elif(d[x][i][8] == '수'):
                    for j in st:
                         
 
                        if j=='0' :
                            wed0.append(cn)
                             
                        elif j =='1':
                            wed1.append(cn)
                             
                        elif j =='2':
                            wed2.append(cn)
 
                        elif j =='3':
                            wed3.append(cn)
 
                        elif j =='4':
                            wed4.append(cn)
 
                        elif j =='5':
                            wed5.append(cn)
 
                        elif j =='6':
                            wed6.append(cn)
 
                        elif j =='7':
                            wed7.append(cn)
 
                        elif j =='8' :
                            wed8.append(cn)
 
                        elif j =='9':
                            wed9.append(cn)
 
 
 
                elif(d[x][i][8] == '목'):
                    for j in st:
                         
                         
                        if j=='0' :
                            thu0.append(cn)
                             
                        elif j =='1':
                            thu1.append(cn)
                             
                        elif j =='2':
                            thu2.append(cn)
 
                        elif j =='3':
                            thu3.append(cn)
 
                        elif j =='4':
                            thu4.append(cn)
 
                        elif j =='5':
                            thu5.append(cn)
 
                        elif j =='6':
                            thu6.append(cn)
 
                        elif j =='7':
                            thu7.append(cn)
 
                        elif j =='8' :
                            thu8.append(cn)
 
                        elif j =='9':
                            thu9.append(cn)
 
 
 
                elif (d[x][i][8] == '금'):
                    for j in st:
                         
               
                        if j=='0' :
                            fri0.append(cn)
                             
                        elif j =='1':
                            fri1.append(cn)
                             
                        elif j =='2':
                            fri2.append(cn)
 
                        elif j =='3':
                            fri3.append(cn)
 
                        elif j =='4':
                            fri4.append(cn)
 
                        elif j =='5':
                            fri5.append(cn)
 
                        elif j =='6':
                            fri6.append(cn)
 
                        elif j =='7':
                            fri7.append(cn)
 
                        elif j == '8' :
                            fri8.append(cn)
                         
                        elif j =='9':
                            fri9.append(cn)
 
         
                 
                if(d[x][i][10] == '월'):
                     
                    for k in stt:
 
                        if (k=='0.0') | (k == '0') :
                            mon0.append(cn)
                             
                        elif (k =='1.0') | (k == '1'):
                            mon1.append(cn)
                             
                        elif (k =='2.0') | (k == '2'):
                            mon2.append(cn)
 
                        elif (k =='3.0') | (k == '3'):
                            mon3.append(cn)
 
                        elif (k =='4.0') | (k == '4'):
                            mon4.append(cn)
 
                        elif (k =='5.0') | (k == '5'):
                            mon5.append(cn)
 
                        elif (k =='6.0') | (k == '6') :
                            mon6.append(cn)
 
                        elif (k =='7.0') | (k == '7'):
                            mon7.append(cn)
 
                        elif (k =='8.0') | (k == '8')  :
                            mon8.append(cn)
 
                        elif (k =='9.0') | (k == '9'):
                            mon9.append(cn)
 
 
                  
                             
                elif(d[x][i][10] == '화'):
 
                    for k in stt:
 
                        if (k=='0.0') | (k == '0') :
                            tue0.append(cn)
                             
                        elif (k =='1.0') | (k == '1'):
                            tue1.append(cn)
                             
                        elif (k =='2.0') | (k == '2'):
                            tue2.append(cn)
 
                        elif (k =='3.0') | (k == '3'):
                            tue3.append(cn)
 
                        elif (k =='4.0') | (k == '4'):
                            tue4.append(cn)
 
                        elif (k =='5.0') | (k == '5'):
                            tue5.append(cn)
 
                        elif (k =='6.0') | (k == '6') :
                            tue6.append(cn)
 
                        elif (k =='7.0') | (k == '7'):
                            tue7.append(cn)
 
                        elif (k =='8.0') | (k == '8')  :
                            tue8.append(cn)
 
                        elif (k =='9.0') | (k == '9'):
                            tue9.append(cn)
 
 
 
 
 
                elif(d[x][i][10] == '수'):
                    for k in stt:
 
                         
                        if (k=='0.0') | (k == '0') :
                            wed0.append(cn)
                             
                        elif (k =='1.0') | (k == '1'):
                            wed1.append(cn)
                             
                        elif (k =='2.0') | (k == '2'):
                            wed2.append(cn)
 
                        elif (k =='3.0') | (k == '3'):
                            wed3.append(cn)
 
                        elif (k =='4.0') | (k == '4'):
                            wed4.append(cn)
 
                        elif (k =='5.0') | (k == '5'):
                            wed5.append(cn)
 
                        elif (k =='6.0') | (k == '6') :
                            wed6.append(cn)
 
                        elif (k =='7.0') | (k == '7'):
                            wed7.append(cn)
 
                        elif (k =='8.0') | (k == '8')  :
                            wed8.append(cn)
 
                        elif (k =='9.0') | (k == '9'):
                            wed9.append(cn)
 
 
 
                elif(d[x][i][10] == '목'):
                    for k in stt:
 
                         
                        if (k=='0.0') | (k == '0') :
                            thu0.append(cn)
                             
                        elif (k =='1.0') | (k == '1'):
                            thu1.append(cn)
                             
                        elif (k =='2.0') | (k == '2'):
                            thu2.append(cn)
 
                        elif (k =='3.0') | (k == '3'):
                            thu3.append(cn)
 
                        elif (k =='4.0') | (k == '4'):
                            thu4.append(cn)
 
                        elif (k =='5.0') | (k == '5'):
                            thu5.append(cn)
 
                        elif (k =='6.0') | (k == '6'):
                            thu6.append(cn)
 
                        elif (k =='7.0') | (k == '7'):
                            thu7.append(cn)
 
                        elif (k =='8.0') | (k == '8'):
                            thu8.append(cn)
 
                        elif (k =='9.0')| (k == '9'):
                            thu9.append(cn)
 
 
 
                elif (d[x][i][10] == '금'):
                    for k in stt:
      
               
                        if (k =='0.0') | (k == '0') :
                            fri0.append(cn)
                             
                        elif (k =='1.0') | (k == '1'):
                            fri1.append(cn)
                             
                        elif (k =='2.0') | (k == '2'):
                            fri2.append(cn)
 
                        elif (k =='3.0') | (k == '3'):
                            fri3.append(cn)
 
                        elif (k =='4.0') | (k == '4'):
                            fri4.append(cn)
 
                        elif (k =='5.0') | (k == '5'):
                            fri5.append(cn)
 
                        elif (k =='6.0') | (k == '6'):
                            fri6.append(cn)
 
                        elif (k =='7.0') | (k == '7'):
                            fri7.append(cn)
 
                        elif (k =='8.0') | (k == '8'):
                            fri8.append(cn)
 
                        elif (k =='9.0') | (k =='9'):
                            fri9.append(cn)
             
            #monday
        
            mon.append(mon0)
            mon.append(mon1)
            mon.append(mon2)
            mon.append(mon3)
            mon.append(mon4)
            mon.append(mon5)
            mon.append(mon6)
            mon.append(mon7)
            mon.append(mon8)
            mon.append(mon9)
            monday.append(mon)
             
            #tuesday
            tue.append(tue0)
            tue.append(tue1)
            tue.append(tue2)
            tue.append(tue3)
            tue.append(tue4)
            tue.append(tue5)
            tue.append(tue6)
            tue.append(tue7)
            tue.append(tue8)
            tue.append(tue9)
            tuesday.append(tue)
             
            #wednesday
            wed.append(wed0)
            wed.append(wed1)
            wed.append(wed2)
            wed.append(wed3)
            wed.append(wed4)
            wed.append(wed5)
            wed.append(wed6)
            wed.append(wed7)
            wed.append(wed8)
            wed.append(wed9)
            wednesday.append(wed)
 
            #thursday
            thu.append(thu0)
            thu.append(thu1)
            thu.append(thu2)
            thu.append(thu3)
            thu.append(thu4)
            thu.append(thu5)
            thu.append(thu6)
            thu.append(thu7)
            thu.append(thu8)
            thu.append(thu9)
            thursday.append(thu)
 
            #friday
            fri.append(fri0)
            fri.append(fri1)
            fri.append(fri2)
            fri.append(fri3)
            fri.append(fri4)
            fri.append(fri5)
            fri.append(fri6)
            fri.append(fri7)
            fri.append(fri8)
            fri.append(fri9)
            friday.append(fri)
 
            del(mon0, mon1, mon2, mon3, mon4, mon5, mon6, mon7, mon8, mon9)
            del(tue0, tue1, tue2, tue3, tue4, tue5, tue6, tue7, tue8, tue9)
            del(wed0, wed1, wed2, wed3, wed4, wed5, wed6, wed7, wed8, wed9)
            del(thu0, thu1, thu2, thu3, thu4, thu5, thu6, thu7, thu8, thu9)
            del(fri0, fri1, fri2, fri3, fri4, fri5, fri6, fri7, fri8, fri9)
             
            del(mon, tue, wed, thu, fri)
                 
    else:
        monday = [[],[],[],[],[],[],[],[],[],[]]
        tuesday = [[],[],[],[],[],[],[],[],[],[]]
        wednesday = [[],[],[],[],[],[],[],[],[],[]]
        thursday = [[],[],[],[],[],[],[],[],[],[]]
        friday = [[],[],[],[],[],[],[],[],[],[]]
 
         
    return monday, tuesday, wednesday, thursday, frida


#gui.py
#programed by 정현아
 
 
from tkinter import*
from tkinter import ttk
from tkinter import messagebox
from PIL import ImageGrab
import new
import pyautogui
import smtplib
from email6.mime.text import MIMEText
from email6.header import Header
from email6.encoders import *
from email6.mime.image import MIMEImage
from email6.mime.multipart import MIMEMultipart
from email6.mime.base import MIMEBase
import os
import socket
 
 
### 시간표를 만드는 프로그램###
def table_make(*event):
     
    lib = [int(choice_lib1.get()), int(choice_lib2.get()),
                     int(choice_lib3.get()), int(choice_lib4.get())]
    cnt = 0
    for i in range(len(lib)):
        if lib[i] == 1:
            cnt +=1
 
    # 선택된 난이도 수보다 교양수가 작아야 에러가 안나기 때문에 예외처리
    if (cnt < int(choice_lib.get())):
 
        messagebox.showinfo('warning','교양수가 선택된 난이도 수보다 작아야 합니다.')
 
    else:
        d = new.Day(str(choice_major.get()), int(choice_num.get()), str(choice_level.get()),
                         str(choice_1.get()),  int(choice_lib.get()),
                         int(choice_lib1.get()), int(choice_lib2.get()),
                         int(choice_lib3.get()), int(choice_lib4.get()))
 
 
        mon = d[0]
        tue = d[1]
        wed = d[2]
        thu = d[3]
        fri = d[4]
 
 
        #월요일
         
        label_11['text'] = mon[0][0]
        label_12['text'] = mon[0][1]
        label_13['text'] = mon[0][2]
        label_14['text'] = mon[0][3]
        label_15['text'] = mon[0][4]
        label_16['text'] = mon[0][5]
        label_17['text'] = mon[0][6]
        label_18['text'] = mon[0][7]
        label_19['text'] = mon[0][8]
        label_110['text'] = mon[0][9]
 
 
 
        #화요일
        label_21['text'] = tue[0][0]
        label_22['text'] = tue[0][1]
        label_23['text'] = tue[0][2]
        label_24['text'] = tue[0][3]
        label_25['text'] = tue[0][4]
        label_26['text'] = tue[0][5]
        label_27['text'] = tue[0][6]
        label_28['text'] = tue[0][7]
        label_29['text'] = tue[0][8]
        label_210['text'] = tue[0][9]
 
        #수요일
        label_31['text'] = wed[0][0]
        label_32['text'] = wed[0][1]
        label_33['text'] = wed[0][2]
        label_34['text'] = wed[0][3]
        label_35['text'] = wed[0][4]
        label_36['text'] = wed[0][5]
        label_37['text'] = wed[0][6]
        label_38['text'] = wed[0][7]
        label_39['text'] = wed[0][8]
        label_310['text'] = wed[0][9]
 
        #목요일
        label_41['text'] = thu[0][0]
        label_42['text'] = thu[0][1]
        label_43['text'] = thu[0][2]
        label_44['text'] = thu[0][3]
        label_45['text'] = thu[0][4]
        label_46['text'] = thu[0][5]
        label_47['text'] = thu[0][6]
        label_48['text'] = thu[0][7]
        label_49['text'] = thu[0][8]
        label_410['text'] = thu[0][9]
 
        #금요일
        label_51['text'] = fri[0][0]
        label_52['text'] = fri[0][1]
        label_53['text'] = fri[0][2]
        label_54['text'] = fri[0][3]
        label_55['text'] = fri[0][4]
        label_56['text'] = fri[0][5]
        label_57['text'] = fri[0][6]
        label_58['text'] = fri[0][7]
        label_59['text'] = fri[0][8]
        label_510['text'] = fri[0][9]
 
   
### 생성된 시간표 캡쳐###
def capture(*event):
     
    x,y = pyautogui.position() #현재 화살표의 위치
     
    icon = ImageGrab.grab(bbox = (x-985, y-810, x-120, y+140))
    #현재 화살표의 위치로 시간표 캡쳐
     
    s = 'table.png'
    icon.save(s)  #캡쳐된 시간표 저장
    messagebox.showinfo('complete!!','저장이 완료되었습니다.')
 
 
### 캡쳐된 시간표를 이메일로 전송하는 프로그램###    
def send_email(*event):
 
    you = str(email_entry.get())
    ad = you.split('@')
    me = '아이디'
    pw = '비밀번호'
            
    file = 'table.png'
     
    #######
    # 네이버의 SMTP서버를 사용하기 때문에 다른 이메일을 주소로 전송하면
    # 첨부파일이 전송안되는 경우가 있음
    # 에러도 아니고 어떤 경우는 되고 어떤 경우는 안되기 때문에
    # 메시지 창으로 알려주기는 하나 이메일로 전송은 진행
    #######
     
    try:
        if ad[1] == 'naver.com':
         
            msg=MIMEMultipart('alternative')
 
            part = MIMEBase('application', 'octet-stream')
 
             
            part.set_payload(open(file, 'rb').read())
            encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"'
                           % os.path.basename(file))
            msg.attach(part)
             
            msg['From'] = me
            msg['To'] = you
 
            mailServer = smtplib.SMTP("smtp.naver.com", 587)
 
            mailServer.ehlo()
            mailServer.starttls()
            mailServer.ehlo()
            mailServer.login(me,pw)
             
            mailServer.sendmail(me, you, msg.as_string())
            mailServer.close()
            messagebox.showinfo('complete!!','전송이 완료되었습니다.')
 
        else:
            messagebox.showinfo('checking!!','본 프로그램은 네이버의 서버를 사용하고 있습니다. \n네이버 이메일 주소를 사용하시길 권장드립니다.')
 
            msg=MIMEMultipart('alternative')
 
            part = MIMEBase('application', 'octet-stream')
 
             
            part.set_payload(open(file, 'rb').read())
            encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"'
                           % os.path.basename(file))
            msg.attach(part)
             
            msg['From'] = me
            msg['To'] = you
 
            mailServer = smtplib.SMTP("smtp.naver.com", 587)
 
            mailServer.ehlo()
            mailServer.starttls()
            mailServer.ehlo()
            mailServer.login(me,pw)
             
            mailServer.sendmail(me, you, msg.as_string())
            mailServer.close()
            messagebox.showinfo('complete!!','전송이 완료되었습니다.')
             
 
    # 캡쳐를 안한 상태에서 이메일 전송을 시도하면 에러가 생김    
    except FileNotFoundError:
        messagebox.showinfo('warning','파일의 생성유무와 경로를 확인해 주십시오.')
 
 
    # 이메일 주소가 잘못되었을 경우 생기는 에러
    except smtplib.SMTPRecipientsRefused:
        email_entry.delete(0, END)
        messagebox.showinfo('warning','이메일 주소를 잘못 입력하셨습니다.')
 
 
    # 인터넷 연결이 안되어 있을 경우 생기는 에러
    except socket.gaierror:
        messagebox.showinfo('warning','인터넷연결을 확인해 주십시오.')
         
 
     
if __name__ == '__main__':
     
    win = Tk()
    win.title("큐티쁘띠흐지쀼")
 
 
    choice_level = IntVar() # 학년 선택 값을 입력받는 정수 변수
 
    choice_num = IntVar() #들을 전공 수의 선택 값을 입력 받는 정수 변수
 
    choice_major = StringVar() #전공 선택 값을 입력받는 문자열 변수
 
    choice_lib = IntVar() #들을 교양 수의 선택 값을 입력 받는 정수 변수
     
    choice_lib1 = IntVar() #교양 난이도 1이 선택되면 1, 선택 안되면 0을 입력 받는 정수 변수
    choice_lib2 = IntVar() #교양 난이도 2이 선택되면 1, 선택 안되면 0을 입력 받는 정수 변수
    choice_lib3 = IntVar() #교양 난이도 3이 선택되면 1, 선택 안되면 0을 입력 받는 정수 변수
    choice_lib4 = IntVar() #교양 난이도 4이 선택되면 1, 선택 안되면 0을 입력 받는 정수 변수
 
    choice_1 = StringVar() #교양 선택 값을 입력받는 문자열 변수
 
 
 
 
    #########label로 표 셀 만들기###########
     
    row_title_label1 =  Label(win, text = '월', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 13)
    row_title_label2 =  Label(win, text = '화', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 13)
    row_title_label3 =  Label(win, text = '수', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 13)
    row_title_label4 =  Label(win, text = '목', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 13)
    row_title_label5 =  Label(win, text = '금', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 13)
 
    col_title_label1 = Label(win, text = '0교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 6, height = 3)
    col_title_label2 = Label(win, text = '1교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 6, height = 3)
    col_title_label3 = Label(win, text = '2교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                               relief = 'ridge', width = 6, height = 3)
    col_title_label4 = Label(win, text = '3교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                              relief = 'ridge', width = 6, height = 3)
    col_title_label5 = Label(win, text = '4교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                             relief = 'ridge', width = 6, height = 3)
    col_title_label6 = Label(win, text = '5교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                              relief = 'ridge', width = 6, height = 3)
    col_title_label7 = Label(win, text = '6교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                              relief = 'ridge', width = 6, height = 3)
    col_title_label8 = Label(win, text = '7교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                             relief = 'ridge', width = 6, height = 3)
    col_title_label9 = Label(win, text = '8교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                              relief = 'ridge', width = 6, height = 3)
    col_title_label10 = Label(win, text = '9교시', font = ('맑은 고딕',10),bg='#FFB2F5',
                             relief = 'ridge', width = 6, height = 3)
 
    #월요일
    label_11 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_12 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_13 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_14 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_15 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_16 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_17 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_18 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_19 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_110 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
 
    #화요일
    label_21 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_22 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_23 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_24 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_25 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_26 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_27 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_28 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_29 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_210 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
 
    #수요일
    label_31 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_32 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_33 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_34 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_35 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_36 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_37 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_38 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_39 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_310 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
 
    #목요일
    label_41 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_42 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_43 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_44 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_45 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_46 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_47 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_48 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_49 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_410 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
 
    #금요일
    label_51 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_52 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_53 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_54 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_55 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_56 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_57 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_58 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_59 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
    label_510 = Label(win, text = '', font = ('맑은 고딕',10), relief = 'ridge', width = 13, height = 3, wraplength = 120)
 
    ##########화살표 버튼###########
    right_arrow = Button(win, text = '▶', width = 10, command = table_make)
    capture = Button(win, text = '캡쳐',command = capture)
 
 
    #########전공##########
    major_label = Label(win, text = '학과', font = ('맑은 고딕',11), padx=2, pady=2)
 
    combo_major = ttk.Combobox(win, textvariable = choice_major)
 
    #########학년###########
    level_label = Label(win, text = '학년' , font = ('맑은 고딕', 11))
 
    level_but1 = Radiobutton(win, text = '1', variable = choice_level, value = 1)
    level_but2 = Radiobutton(win, text = '2', variable = choice_level, value = 2)
    level_but3 = Radiobutton(win, text = '3', variable = choice_level, value = 3)
    level_but4 = Radiobutton(win, text = '4', variable = choice_level, value = 4)
 
 
    ########들을 전공 수########
    numlev_label = Label(win, text = '들을 전공 수', font = ('맑은 고딕',11))
 
    num_but1 = Radiobutton(win, text = '1', variable = choice_num, value = 1)
    num_but2 = Radiobutton(win, text = '2', variable = choice_num, value = 2)
    num_but3 = Radiobutton(win, text = '3', variable = choice_num, value = 3)
    num_but4 = Radiobutton(win, text = '4', variable = choice_num, value = 4)
    num_but5 = Radiobutton(win, text = '5', variable = choice_num, value = 5)
    num_but6 = Radiobutton(win, text = '6', variable = choice_num, value = 6)
    num_but7 = Radiobutton(win, text = '7', variable = choice_num, value = 7)
 
    ############들을 교양 수 #################
    lib_la = Label(win, text = '들을 교양 수', font = ('맑은 고딕',11))
 
    lib_but1 = Radiobutton(win, text = '1', variable = choice_lib, value = 1)
    lib_but2 = Radiobutton(win, text = '2', variable = choice_lib, value = 2)
    lib_but3 = Radiobutton(win, text = '3', variable = choice_lib, value = 3)
    lib_but4 = Radiobutton(win, text = '4', variable = choice_lib, value = 4)
 
 
    ###########교양 난이도###########
    lib_label = Label(win, text = '교양 난이도', font = ('맑은 고딕',11))
 
    lib_box1 = Checkbutton(win, text = '1', variable = choice_lib1, onvalue= 1, offvalue=0)
    lib_box2 = Checkbutton(win, text = '2', variable = choice_lib2, onvalue= 1, offvalue=0)
    lib_box3 = Checkbutton(win, text = '3', variable = choice_lib3, onvalue= 1, offvalue=0)
    lib_box4 = Checkbutton(win, text = '4', variable = choice_lib4, onvalue= 1, offvalue=0)
 
 
    #############교양###############
 
    lib1_label = Label(win, text = '교양', font = ('맑은 고딕',11))
 
    lib1 = ttk.Combobox(win, textvariable = choice_1)
    lib1['value'] = ('교양_과학과기술', '교양_글로벌문화와제2외국어' , '교양_사회와경제',
                     '교양_서울권역e러닝' ,'교양_실용영어','교양_언어와표현','교양_예술과체육',
                     '교양_인간과철학')
 
 
    ##############이메일################
    email_label = Label(win, text = 'E-mail', font = ('맑은 고딕',11))
    email_entry = Entry(win)
    email_but = Button(win, text = '보내기', command= send_email)
    mes = Label(win, text = '※이 정보는 이메일을 보내는 용도로만 사용됩니다.', font = ('맑은 고딕',11))
 
    build_but= Button(win, text = '만들기', command= table_make)
 
 
    ##########전공 메뉴추가##########
    combo_major['value'] = ('건축공학과(공대)', '건축학과(공대)', '경영학부(경영대)',
                      '국어국문(인문대)', '국어국문(인문대)', '국제통상학과(동북아대)',
                      '국제통상학부(경영대)', '국제학부(동북아대)', '국제학부(정법대)',
                      '동북문화산업학부(동북아대)', '동북문화산업학부(인사대)',
                      '동북아통상학부(동북아대)', '로봇학부(전정대)', '미디어영상학부(사과대)',
                      '미디어영상학부(인사대)', '법학과(정법대)', '법학부(법과대)',
                      '부동산법무학과(법과대)', '사이버정보보안학과(자과대)', '산업심리학과(사과대)'
                      '산업심리학과(인사대)', '생활체육학과(자과대)', '소프트웨어학부(소융대)',
                      '수학과(자과대)', '영어영문(인문대)', '영어영문(인사대)', '자산관리학과(법과대)',
                      '자산관리학과(정법대)', '전기공학과(전정대)', '전자공학과(전정대)',
                      '전자바이오물리학과(자과대)', '전자재료공학과(전정대)', '전자통신공학과(전정대)',
                      '정보융합학부(소융대)','정보콘텐츠학과(자과대)','컴퓨터공학과(전정대)',
                      '컴퓨터소프트웨어학과(전정대)','컴퓨터정보공학부(소융대)','행정학과(사과대)',
                      '행정학과(정법대)','화학공학과(공대)','화학공학과(공대)','화학과(자과대)',
                      '환경공학과(공대)')
 
 
 
 
 
 
 
 
 
    ######## 시간표 셀 배치###########
    row_title_label1.grid(column = 1, row = 0)
    row_title_label2.grid(column = 2, row = 0)
    row_title_label3.grid(column = 3, row = 0)
    row_title_label4.grid(column = 4, row = 0)
    row_title_label5.grid(column = 5, row = 0)
 
    col_title_label1.grid(column = 0, row = 1)
    col_title_label2.grid(column = 0, row = 2)
    col_title_label3.grid(column = 0, row = 3)
    col_title_label4.grid(column = 0, row = 4)
    col_title_label5.grid(column = 0, row = 5)
    col_title_label6.grid(column = 0, row = 6)
    col_title_label7.grid(column = 0, row = 7)
    col_title_label8.grid(column = 0, row = 8)
    col_title_label9.grid(column = 0, row = 9)
    col_title_label10.grid(column = 0, row = 10)
 
 
    #월요일
    label_11.grid(column = 1, row = 1)
    label_12.grid(column = 1, row = 2)
    label_13.grid(column = 1, row = 3)
    label_14.grid(column = 1, row = 4)
    label_15.grid(column = 1, row = 5)
    label_16.grid(column = 1, row = 6)
    label_17.grid(column = 1, row = 7)
    label_18.grid(column = 1, row = 8)
    label_19.grid(column = 1, row = 9)
    label_110.grid(column = 1, row = 10)
 
    #화요일
    label_21.grid(column = 2, row = 1)
    label_22.grid(column = 2, row = 2)
    label_23.grid(column = 2, row = 3)
    label_24.grid(column = 2, row = 4)
    label_25.grid(column = 2, row = 5)
    label_26.grid(column = 2, row = 6)
    label_27.grid(column = 2, row = 7)
    label_28.grid(column = 2, row = 8)
    label_29.grid(column = 2, row = 9)
    label_210.grid(column = 2, row = 10)
 
    #수요일
    label_31.grid(column = 3, row = 1)
    label_32.grid(column = 3, row = 2)
    label_33.grid(column = 3, row = 3)
    label_34.grid(column = 3, row = 4)
    label_35.grid(column = 3, row = 5)
    label_36.grid(column = 3, row = 6)
    label_37.grid(column = 3, row = 7)
    label_38.grid(column = 3, row = 8)
    label_39.grid(column = 3, row = 9)
    label_310.grid(column = 3, row = 10)
 
    #목요일
    label_41.grid(column = 4, row = 1)
    label_42.grid(column = 4, row = 2)
    label_43.grid(column = 4, row = 3)
    label_44.grid(column = 4, row = 4)
    label_45.grid(column = 4, row = 5)
    label_46.grid(column = 4, row = 6)
    label_47.grid(column = 4, row = 7)
    label_48.grid(column = 4, row = 8)
    label_49.grid(column = 4, row = 9)
    label_410.grid(column = 4, row = 10)
 
    #금요일
    label_51.grid(column = 5, row = 1)
    label_52.grid(column = 5, row = 2)
    label_53.grid(column = 5, row = 3)
    label_54.grid(column = 5, row = 4)
    label_55.grid(column = 5, row = 5)
    label_56.grid(column = 5, row = 6)
    label_57.grid(column = 5, row = 7)
    label_58.grid(column = 5, row = 8)
    label_59.grid(column = 5, row = 9)
    label_510.grid(column = 5, row = 10)
 
 
    ######### 화살표 위치선정##############
    right_arrow.grid(column = 7, row = 9)
    capture.grid(column = 8, row = 9)
 
    ############전공 선택 위치###############
    major_label.grid(column = 7, row = 0, sticky = E)
 
    combo_major.grid(column = 8, row = 0, columnspan=3)
 
    ############학년 위치#############
    level_label.grid(column = 7, row =1, sticky = E)
 
    level_but1.grid(column = 8, row =1, sticky = E)
    level_but2.grid(column = 9, row =1, sticky = W)
    level_but3.grid(column = 10, row =1, sticky = W)
    level_but4.grid(column = 11, row =1, sticky = W)
 
 
    ##########들을 전공의 수 위치###########
    numlev_label.grid(column = 7, row =2, columnspan=2,sticky = SE)
 
    num_but1.grid(column = 7, row =3, sticky = NE)
    num_but2.grid(column = 8, row =3, sticky = N)
    num_but3.grid(column = 9, row =3, sticky = N)
    num_but4.grid(column = 10, row =3, sticky = NW)
    num_but5.grid(column = 11, row =3, sticky = NW)
    num_but6.grid(column = 12, row =3, sticky = NW)
    num_but7.grid(column = 13, row =3, sticky = NW)
 
    ############들을 교양 수 위치 #################
    lib_la.grid(column = 7, row =4, columnspan=2,sticky = SE) 
 
    lib_but1.grid(column = 7, row =5, sticky = NE)
    lib_but2.grid(column = 8, row =5, sticky = N)
    lib_but3.grid(column = 9, row =5, sticky = NW)
    lib_but4.grid(column = 10, row =5, sticky = NW)
 
 
    ############교양 난이도 위치#############
    lib_label.grid(column = 7, row =6, columnspan=2,sticky = SE)
 
    lib_box1.grid(column = 7, row =7, sticky = NE)
    lib_box2.grid(column = 8, row =7, sticky = N)
    lib_box3.grid(column = 9, row =7, sticky = N)
    lib_box4.grid(column = 10, row =7, sticky = NW)
 
 
    ############교양###############
 
    lib1_label.grid(column = 7, row = 8, sticky = E)
    lib1.grid(column = 8, row = 8, columnspan = 3, sticky = W)
 
 
 
    ##############이메일 위치################
 
    email_label.grid(column = 7, row = 10, sticky = E)
    email_entry.grid(column = 8, row = 10, columnspan=3, sticky = W)
    email_but.grid(column = 11, row = 10, sticky = W)
    mes.grid(column = 7, row = 11, columnspan=7, sticky = W)
 
     
    ####### 보내기 버튼 위치######
    build_but.grid(column = 9, row = 9, padx = 5, pady=5)
 
 
    win.mainloop()