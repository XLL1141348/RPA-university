py begin
# -*- coding: utf-8 -*-
import os
import csv

dirpath="/root/rpa_file/"

if (os.path.exists(dirpath+"start")):
    os.remove(dirpath+"start")
    
if (os.path.exists(dirpath+"finish")):
    os.remove(dirpath+"finish")
    
if (os.path.exists(dirpath+"error.txt")):
    os.remove(dirpath+"error.txt")
    
if (os.path.exists(dirpath+"retry")):
    os.remove(dirpath+"retry")

if (os.path.exists(dirpath+"log.txt")):
    os.remove(dirpath+"log.txt")

if (os.path.exists("/root/record_scores/period.csv")):
    os.remove("/root/record_scores/period.csv")
    
py finish

