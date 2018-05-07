#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import base64
from time import sleep
import pyodbc

'''
'''

script_dir=os.path.split(os.path.realpath(__file__))[0]+'/'

file_save_dir='/cygdrive/f/tempdown/wos/archives'
csv_files=['f69_life.txt','f69_pt.txt']




def dmesg(m):
    global dbg
    if dbg:
        print '[dmesg] %s'%m





def retrive(d):
    mp={}
    mp['TI'] =('title',255)
    mp['DE'] =('De',255)
    mp['AB'] =('Abstract',0)
    mp['C1'] =('address',255)
    mp['RP'] =('response',255)
    mp['PY'] =('pyear',0)
    mp['IS'] =('issue',5)
    mp['VL'] =('volumn',5)
    mp['BP'] =('beginpage',10)
    mp['EP'] =('endpage',10)
    mp['DI'] =('DOI',50)
    mp['GA'] =('GA',10)
    mp['UT'] =('UT',25)
    mp['J9'] =('abbr',30)
    mp['SN'] =('issn',10)
    mp['EM'] =('email',150)
    mp['PD'] =('PD',10)
    mp['ID'] =('id1',250)
    mp['AU'] =('Authors',250)
    mp['AF'] =('Author_full',250)
    mp['SO'] =('fullname',140)
    mp['AR'] =('AR',6)
    mp['SC'] =('SC',200)
    mp['SC'] =('typename',30)
    mp['FU'] =('fund',255)
    mp['FX'] =('fundx',1000)

    mp2={}
    #另有三列按以下几条规则，都是针对数字型
    #pages  #case when isnumeric(PG)=1 then cast(PG as int) else 2 end 
    #TC     #case when isnumeric(TC)=1 then cast(TC as int) else NULL end 
    #NR     #case when isnumeric(NR)=1 then cast(NR as int) else NULL end 
    mp2['PG'] =('pages',2)
    mp2['TC'] =('TC',1)
    mp2['NR'] =('NR',1)

    d2={}
    
    for h in mp:
        if mp[h][1]>0:
            d2[mp[h][0]]=d[h][:mp[h][1]]
        else:
            d2[mp[h][0]]=d[h]
    for h in mp2:
        try:
            d2[mp2[h][0]]=int(d[h])
        except:
            d2[mp2[h][0]]=mp2[h][1]

    #另几个需要默认值的列
    d2['r_id']=0
    return d2


def save(d):
    global cursor
    ks=d.keys()
    sql='insert into [target_table](%s) values(%s)'%(
        ','.join(ks),
        ','.join(['?']*len(ks))
        )
    try:
        cursor.execute(sql,tuple([d[k] for k in ks]))
        cursor.commit()
    except pyodbc.DatabaseError,e:
        print '[Error]**** (sql,data,e.args[0],e.args[1])'
        print sql
        print d
        print e.args[0]
        print e.args[1].decode('gbk').encode('utf-8')
        exit()
    except UnicodeDecodeError,e:
        print "UnicodeDecodeError: ",e
        pass


def run():
    global cursor
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.7.114;DATABASE=dbname;UID=sa;PWD=sa')
    cursor=conn.cursor()
    for i in range(len(csv_files)):
        f=open('%s/%s'%(file_save_dir,csv_files[i]),'r')
        line=f.readline().strip('\xef\xbb\xbf\r\n')
        heads=line.split('\t')
        print heads
        n=1
        while True:
            n+=1
            line=f.readline().strip('\xef\xbb\xbf\r\n')
            if len(line)==0:
                print 'empty line, done'
                break
            '''if n >= 500:
                print 'testing, stoped in %s lines'%(n-1)
                break
            '''
            val=line.split('\t')
            if len(val) < len(heads):
                print 'fields count less than heads in line %s, skiped'%n
                continue
            d={}
            for j in range(len(heads)):
                d[heads[j]]=val[j]
            d2=retrive(d)
            save(d2)
            if n % 1000 == 0:
                print n

        f.close()
    conn.close()


if __name__ == '__main__':
    run()
