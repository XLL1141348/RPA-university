py begin
# -*- coding: utf-8 -*-
import os
import csv

dirpath="/root/record_scores/"

if (os.path.exists(dirpath+"info.csv")):
    os.remove(dirpath+"info.csv")
if (os.path.exists(dirpath+"wrong.csv")):
    os.remove(dirpath+"wrong.csv")
if (os.path.exists(dirpath+"record.csv")):
    os.remove(dirpath+"record.csv")
if (os.path.exists(dirpath+"web_number.csv")):
    os.remove(dirpath+"web_number.csv")

with open(dirpath+"info.csv","w") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["upload_file","course_series","grade_selection","username","password","grade_persentage"])

py finish

