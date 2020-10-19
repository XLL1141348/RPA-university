py begin
# -*- coding: utf-8 -*-
import os
import time

dirpath="/root/rpa_file/"
#若存在record.txt文件则删除
if (os.path.exists(dirpath+"log.txt")):
    os.remove(dirpath+"log.txt")

#创建标志文件
def text_create(name,msg):
    desktop_path = dirpath  
    full_path = desktop_path + name+'.txt'  
    file = open(full_path, 'w')
    file.write(msg) 
    file_1 = open(full_path, 'a')
    file.write("流程开始\n")
   

text_create('log',time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))

#log日志写入
def text_append(msg): 
    full_path = dirpath+"log.txt"  
    file = open(full_path, 'a')
    file.write(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))) 
    file_1 = open(full_path, 'a')
    file.write(msg)



    
py finish




