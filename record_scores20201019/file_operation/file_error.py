py begin
# -*- coding: utf-8 -*-
import os
#dirpath是获取创建的路径
dirpath="/root/rpa_test/"
#若存在finish文件则删除
if (os.path.exists(dirpath+"error")):
    os.remove(dirpath+"error")

#创建标志文件
def text_create(name):
    desktop_path = dirpath  
    full_path = desktop_path + name  
    file = open(full_path, 'w')

text_create('error')

py finish



