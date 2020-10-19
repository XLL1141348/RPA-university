py begin
# -*- coding: utf-8 -*-
import os

dirpath="/root/rpa_file/"
#若存在start文件则删除
if (os.path.exists(dirpath+"start")):
    os.remove(dirpath+"start")

#创建标志文件
def text_create(name):
    desktop_path = dirpath  
    full_path = desktop_path + name  
    file = open(full_path, 'w')

text_create('start')

py finish

