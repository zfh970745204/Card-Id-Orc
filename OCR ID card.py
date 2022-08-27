#!/usr/bin/env python
# -*- coding: utf-8 -*-

import glob
import base64
import json
import base64
import threading
import time
from tkinter import *
from tkinter import filedialog
import tkinter
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo
import pandas as pd
import requests
from tkinter import ttk
import json
import base64

import base64

ENCODING = 'utf-8'

#选择文件夹返回文件夹的路径
def selectPath():
    init()
    path_ = askdirectory() #使用askdirectory()方法返回文件夹的路径
    if path_ == "":
        path.set(path_)
        path.get() #当打开文件路径选择框后点击"取消" 输入框会清空路径，所以使用get()方法再获取一次路径
        showinfo('提示', '未选择文件夹')
    else:
        path_ = path_.replace("/", "\\")  # 实际在代码中执行的路径为“\“ 所以替换一下
        path.set(path_)
        showinfo('提示', '已选择'+str(path.get())+'文件夹！') 
    print("路径："+str(glob.glob(path.get()+"\\*")))

#识别按钮实现线程同时开始
def shibie():
    if path.get() != '':
        thread_it(create) 
        thread_it(tijiao) 
    else:
        showinfo('提示', '请选择文件夹！') 

#调用阿里云接口
def tijiao():
    init()
    global statzzx
    statzzx=0
    id =0  

    for file_abs in glob.glob(path.get()+"\\*"):
        file_ab = file_abs.replace("\\", "/")
        #如果没有configure字段，configure设为None
        #configure = None
        img_base64data = get_img_base64(file_ab)
        try:         
            name, sex, nationality, birth, address, num = predict(url, appcode, img_base64data, configure)   
        except TypeError:
            print("图片错误")
            continue
        
        if flag == 0: 
            id=id+1
            img_file.append({'姓名': name, '性别': sex, '民族': nationality, '出生': birth, '住址': address, '身份证号码': num,"图片路径":file_abs})
            img_file1.append([id ,name, sex, nationality,  birth,  address,num,file_abs])            
        else:
            print('识别错误')
            continue    
    statzzx=1    
    insert()
    
#将函数打包进线程    
def thread_it(func):
    '''将函数打包进线程'''
    # 创建
    t = threading.Thread(target=func) 
    # 守护 !!!
    t.setDaemon(True) 
    # 启动
    t.start()
#将数据导出表格
def writeExcel():
    if len(img_file):
        # 存在值即为真
        pf = pd.DataFrame(img_file)
        order = ['姓名', '性别', '民族', '出生', '住址', '身份证号码','图片路径']
        pf = pf[order]
        file_path = filedialog.asksaveasfilename(defaultextension='.py',filetypes = [("Excel files",".xlsx")])
        print("文件保存路径："+str(file_path))

        print("sadfsafasfafasf"+str(img_file[0]))
        print("sadfsafasfafasf"+str(img_file[0]['姓名']))
        pf.to_excel(file_path, encoding='utf-8', index=False, sheet_name="身份证信息")
        
        print("导出Excel成功!") 
        showinfo('成功', '导出Excel成功!')  
    else:
        print("请选择文件夹！") 
        showinfo('提示', '请先批量识别！')   
#文件识别进度框
def create():    
    top = Toplevel()
    top.title('文件识别中...')

    pb = ttk.Progressbar(top, length=280, mode="determinate", orient=HORIZONTAL)#indeterminate determinate
    w = 300
    h = 70
    x1 = int((screenwidth - w) / 2)
    y1 = int((screenheight - h) / 2)
    top.geometry('{}x{}+{}+{}'.format(w, h, x1, y1))
    pb.pack(padx=10, pady=20)
    pb["maximum"] = 100
    pb["value"] = 0
    
    print("\n"*2)
    print("执行开始".center(scale+28,'-'))
    start = time.perf_counter()
    for i in range(scale+1):
        time.sleep(0.03)
        if(statzzx != 1):
            pb["value"] = i      # 每次更新1
            root.update()            # 更新画面
            a = '*' * i
            b = '.' * (scale - i)
            c = (i/scale)*100
            t = time.perf_counter() - start
            print("\r任务进度:{:>3.0f}% [{}->{}]消耗时间:{:.2f}s".format(c,a,b,t),end="")
        else:
            print("文件获取成功！")   
            showinfo('提示', '文件识别成功！')         
            break
    print("\n"+"执行结束".center(scale+28,'-'))
    top.destroy()
    

#表格数据插入    
def insert():
    # 插入数据
    for index, data in enumerate(img_file1):
        table.insert('', END, values=data)  # 添加数据到末尾
#表格数据删除
def delete():
    obj = table.get_children()  # 获取所有对象
    for o in obj:
        table.delete(o)  # 删除对象
#图片转码
def get_img_base64(img_file):
    with open(img_file, 'rb') as infile:
        s = infile.read()
        return base64.b64encode(s).decode(ENCODING)
#接口访问
def predict(url, appcode, img_base64, kv_configure):
        param = {}
        param['image'] = img_base64
        if kv_configure is not None:
            param['configure'] = json.dumps(kv_configure)
        body = json.dumps(param)
        data1 = bytes(body, "utf-8")

        headers = {'Authorization' : 'APPCODE %s' % appcode}
        response = requests.post(url = url, headers = headers, data = data1)
        if response:
            data = response.json()
            print(data)
            name = data['name']
            sex = data['sex']
            nationality = data['nationality']
            birth = data['birth']
            address = data['address']
            num = data['num']
            return (name, sex, nationality, birth, address, num)
        else:
            flag = 1
            return flag
            
root = Tk()
root.title("身份证信息批量获取")
path = StringVar()

appcode = '' #阿里云接口APPCODE,阿里云1分钱500次调用，https://market.aliyun.com/products/57124001/cmapi010401.html?spm=5176.2020520132.101.3.4e157218wordQA#sku=yuncode440100000

url = 'http://dm-51.data.aliyun.com/rest/160601/ocr/ocr_idcard.json'
configure = {'side':'face'}
flag = 0
scale=100
statzzx = 0
img_file= []
img_file1= []

screenwidth = root.winfo_screenwidth()  # 屏幕宽度
screenheight = root.winfo_screenheight()  # 屏幕高度
width = 1000
height = 500
x = int((screenwidth - width) / 2)
y = int((screenheight - height) / 2)
root.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置
tabel_frame = tkinter.Frame(root)
xscroll = Scrollbar(tabel_frame, orient=HORIZONTAL)
yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

columns = ['id', '姓名', '性别', '民族', '出生', '住址', '身份证号码','图片路径']
table = ttk.Treeview(
        master=root,  # 父容器
        height=10,  # 表格显示的行数,height行
        columns=columns,  # 显示的列
        show='headings',  # 隐藏首列
        xscrollcommand=xscroll.set,  # x轴滚动条
        yscrollcommand=yscroll.set,  # y轴滚动条
        )

#初始化控件
def init():
    root.grid_columnconfigure(1, minsize=200)  # Here
    table.heading('id', text='序号', )  # 定义表头
    table.heading('姓名', text='姓名', )  # 定义表头
    table.heading('性别', text='性别', )  # 定义表头
    table.heading('民族', text='民族', )  # 定义表头
    table.heading('出生', text='出生', )  # 定义表头
    table.heading('住址', text='住址', )  # 定义表头
    table.heading('身份证号码', text='身份证号码', )  # 定义表头
    table.heading('图片路径', text='图片路径', )  # 定义表头

    table.column('id', width=10, minwidth=10, anchor=S, )  # 定义列
    table.column('姓名', width=30, minwidth=30, anchor=S, )  # 定义列
    table.column('性别', width=20, minwidth=10, anchor=S)  # 定义列
    table.column('民族', width=20, minwidth=10, anchor=S)  # 定义列
    table.column('出生', width=50, minwidth=50, anchor=S)  # 定义列
    table.column('住址', width=200, minwidth=100, anchor=S)  # 定义列
    table.column('身份证号码', width=150, minwidth=100, anchor=S)  # 定义列
    table.column('图片路径', width=150, minwidth=100, anchor=S)  # 定义列
    table.grid(row=3,columnspan = 4, padx = 18,ipadx = 165,ipady = 100,pady=10)
    delete()
    global img_file
    img_file=[]
    global img_file1
    img_file1=[]

def demo():    
    Button(root, text="文件夹批量选择", command=lambda :thread_it(selectPath),width=15).grid(row=0, column=0,padx=18,pady=10,sticky = 'w')
    Entry(root, textvariable=path,state="readonly",width=83).grid(row=0, column=1,pady=10,sticky = 'w')
    Button(root, text="批量识别", command=shibie,width=13).grid(row=0, padx=3,column=2,sticky = 'w',pady=10)
    Button(root, text="导出表格", command=writeExcel,width=13).grid(row=0, padx=2,column=3,sticky = 'w',pady=10)    
    init() 
        
if __name__ == '__main__':
    demo()
    root.mainloop()
