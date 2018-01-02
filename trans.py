import docx
from win32com import client as wc
import os

def Trans2Docx(name):
    # 首先将doc转换成docx
    word = wc.Dispatch("Word.Application")

    doc = word.Documents.Open(r"D:\\IN\\常识\\"+name+".doc")
    #使用参数16表示将doc转换成docx
    outName = r"D:/OUT/"+name+".docx"
    print(outName)
    doc.SaveAs(outName,16)
    doc.Close()
    word.Quit()

path = "D:\\IN\\常识" #文件夹目录
files= os.listdir(path) #得到文件夹下的所有文件名称
s ={}
for file in files: #遍历文件夹
     if not os.path.isdir(file): #判断是否是文件夹，不是文件夹才打开
         name = file.split('.')[0].replace('~','').replace('$','')
         s[name]=''


for key in s:
    print(key)
    Trans2Docx(key)

