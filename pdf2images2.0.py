import re
import os
import sys, fitz
import xlwt
import openpyxl
import xlsxwriter
import datetime
import time
from os import listdir

from PIL import ImageDraw, ImageFont
import PIL.Image
from tkinter import *
from tkinter.filedialog import askdirectory

import tkinter
from tkinter import ttk # 导入ttk模块，因为下拉菜单控件在ttk中


from goto import with_goto

def pyMuPDF_fitz(pdfPath, imagePath, imageName):
    startTime_pdf2img = datetime.datetime.now()  # 开始时间


    print("图片输出路径为：" + imagePath)
    print("正在转化，请稍后...")

    file_name = os.path.basename(pdfPath)  # 获取文件名字

    name = file_name.split('.')[0]  # 去除后缀，获取名字

    pdfDoc = fitz.open(pdfPath)
    for pg in range(pdfDoc.pageCount):
        page = pdfDoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=96
        #zoom_x = 1.33333333  # (1.33333333-->1056x816)   (2-->1584x1224)
        #zoom_y = 1.33333333
        # 缩放系数都为2，分辨率提高4倍
        #zoom_x = 2  # (1.33333333-->1056x816)   (2-->1584x1224)
        #zoom_y = 2

        #zoom_x = 1  # (1.33333333-->1056x816)   (2-->1584x1224)
        #zoom_y = 1

        zoom_x = 1.111111  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 1.111111




        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)

        if not os.path.exists(imagePath):  # 判断存放图片的文件夹是否存在
            os.makedirs(imagePath)  # 若图片文件夹不存在就创建


        # 页码从0开始，水印都加一
        pix.writePNG(imagePath + '/' + imageName + '-%s.JPEG' % str(int(pg) + 1))  # 将图片写入指定的文件夹内



        # 为图片添加水印
        imageInfo = PIL.Image.open(imagePath + '/' + imageName + '-%s.JPEG' % str(int(pg) + 1))

        # 文字水印

        fontOne = ImageFont.truetype("‪C:\Windows\Fonts\simfang.ttf", 30)  # 本地字体文件

        draw = ImageDraw.Draw(imageInfo)
        #print(imageInfo.size)

        # 深灰色fill=(96, 96, 96)
        # 开始添加文字水印

        draw.text((imageInfo.size[0] // 2 - 20, imageInfo.size[1] // 2 + 406), u"第%s页"%str(int(pg) + 1), fill=(0, 0, 0), font=fontOne)

        # imageInfo.show()  #展示图片

        # 添加图片水印
        logo = PIL.Image.open("C:\\logo.png")

        w, h = logo.size   # 获取图像宽高
        logo.thumbnail((800, 800))   # 图像缩小1/2，图像缩放


        layer = PIL.Image.new('RGBA', imageInfo.size, (255, 255, 255, 0))

        # 添加图片水印

        layer.paste(logo, (imageInfo.size[0] - logo.size[0] + 70, imageInfo.size[1] - logo.size[1] - 20))

        imageInfo = PIL.Image.composite(layer, imageInfo, layer)

        #imageInfo.paste(logo, (0, 0))    # 将一张图片覆盖到另外一张图片上

        imageInfo.save(imagePath + '/' + imageName + '-%s.JPEG' % str(int(pg) + 1), quality=80, optimize=True)

    endTime_pdf2img = datetime.datetime.now()  # 结束时间
    print('pdf2img转换时间:', (endTime_pdf2img - startTime_pdf2img).seconds)

    # 拼接为长图
    '''imgs = [Image.open(imagePath + '\\' + fn) for fn in listdir(imagePath) if fn.endswith(".png")]  # 打开路径下的所有图片

    width, height = imgs[0].size  # 获取拼接图片的宽和高

    print(imgs)
    result = Image.new(imgs[0].mode, (width, height * len(imgs)))

    for j, im in enumerate(imgs):
        result.paste(im, box=(0, j * height))
        print(j)
    result.save(imagePath + '/' + name + '.png')  # 以pdf文件名命名'''





class MyFrm(Frame):
    def __init__(self, master):
        self.root=master
        self.screen_width = self.root.winfo_screenwidth()#获得屏幕宽度
        self.screen_height = self.root.winfo_screenheight()  #获得屏幕高度
        #self.root.resizable(False, False)#让高宽都固定
        self.root.update_idletasks()#刷新GUI
        self.root.withdraw() #暂时不显示窗口来移动位置
        self.root.geometry('%dx%d+%d+%d' % (self.root.winfo_width(), self.root.winfo_height() ,(self.screen_width - self.root.winfo_width()) / 2,(self.screen_height - self.root.winfo_height()) / 2))  # center window on desktop
        self.root.deiconify()

if __name__ == "__main__":
    # 需要处理的pdf文件夹路径（绝对路径）
    # pdfPath = "D:\\Pycharm\\python\\paper\\2016"
    # D:\\Pycharm\\python\\paper\\pdf\\2019     D:\\paper\\2019
    root = Tk()   # 打开一个窗口
    root.title('pdf2image')
    #root.geometry("400x300+10+20")       # 设置窗口大小
    #MyFrm(root)    # 窗口居中显示
    w = root.winfo_screenwidth()    # 电脑屏幕宽度

    h = root.winfo_screenheight()   # 电脑屏幕高度

    root.geometry("%dx%d" % (w, h))

    # root.attributes("-topmost", True)
    # root.configure(background='green') # 设置背景颜色

    photo = PhotoImage(file="C:\\logo.png")

    Label(root, image=photo).place(relx=0.5, rely=0.6, anchor=CENTER)

    # 绘制两个label, grid（）确定行列
    e1 = StringVar()
    e2 = StringVar()
    e3 = StringVar()
    Label(root, text="请输入学校名称：").place(relx=0.45, rely=0.2, anchor=CENTER)
    Label(root, text="请输入学校代码：").place(relx=0.45, rely=0.25, anchor=CENTER)
    Label(root, text="请输入真题年份：").place(relx=0.45, rely=0.3, anchor=CENTER)

    e1 = Entry(root, textvariable=e1)
    e2 = Entry(root, textvariable=e2)
    e3 = Entry(root, textvariable=e3)



    e1.place(relx=0.55, rely=0.2, anchor=CENTER)
    e2.place(relx=0.55, rely=0.25, anchor=CENTER)
    e3.place(relx=0.55, rely=0.3, anchor=CENTER)

    #sch_name = e1.get()
    #sch_code = e2.get()

    # 打开文件GUI
    def selectPath():
        path_ = askdirectory()
        path.set(path_)


    def define():
        sch_name = e1.get()
        sch_code = e2.get()
        date = e3.get()


    def clear():
        e1.delete(0, END)
        e2.delete(0, END)
        e3.delete(0, END)


    path = StringVar()
    Label(root, text="请选择真题文件：").place(relx=0.45, rely=0.35, anchor=CENTER)
    Entry(root, textvariable=path).place(relx=0.55, rely=0.35, anchor=CENTER)
    Button(root, text="点击选择文件夹", command=selectPath).place(relx=0.65, rely=0.35, anchor=CENTER)


    def Info():
        sch_name = e1.get()    # 学校名称
        sch_name = sch_name.strip()  # 去除空格
        sch_name = sch_name.replace(" ", "")  # 去除换行符
        sch_name = sch_name.replace("\n", "")  # 去除换行符

        sch_code = e2.get()     # 学校代码
        sch_code = sch_code.strip()  # 去除空格
        sch_code = sch_code.replace(" ", "")  # 去除换行符
        sch_code = sch_code.replace("\n", "")  # 去除换行符

        date = e3.get()     # 日期
        date = date.strip()  # 去除空格
        date = date.replace(" ", "")  # 去除换行符
        date = date.replace("\n", "")  # 去除换行符
        date = date[0:4]    # 去除汉字


        pdfPath = path.get()    # 文件路径

        print('学校:' + sch_name + ',' + '学校代码:' + sch_code + ',' + '真题年份:' + date + ',' + '文件路径:' + pdfPath)

        # 载入文件列表
        file_list = os.listdir(pdfPath)
        # 文件排序
        file_list.sort()
        # 指定字符串总数
        count = 0

        # 中山大学+4144010558+2019
        #info = info.split('+')  # 切割学校信息
        # 学校名称
        #sch_name = info[0]
        # 学校代码
        #sch_code = info[1]
        # 真题年份
        #date = info[2]
        # 定义excel表名
        workbook = xlsxwriter.Workbook(sch_name + date + '真题信息表' + '.xlsx')
        # 创建一个空表
        worksheet = workbook.add_worksheet()

        # 遍历所有文件，对pdf文件进行重命名，并获取真题信息
        for file in file_list:
            count = count + 1
            pdfpath = pdfPath + '\\' + file  # pdf文件
            file_name = os.path.basename(pdfpath)  # 获取文件名字

            name = file_name.split('.')[0]  # 去除后缀，获取名字

            pdfDoc = fitz.open(pdfpath)  # 读取页码数
            page_num = pdfDoc.pageCount  # 页码

            # 将数据逐行写入excel表中
            # 表头----第一行
            worksheet.write(0, 0, '编号')
            worksheet.write(0, 1, '学校')
            worksheet.write(0, 2, '科目')
            worksheet.write(0, 3, '年份')
            worksheet.write(0, 4, '真题标签')
            worksheet.write(0, 5, '页码数')

            # 写入数据
            worksheet.write(count, 0, count)
            worksheet.write(count, 1, sch_name)
            worksheet.write(count, 2, name)
            worksheet.write(count, 3, date)
            worksheet.write(count, 4, 'pdf;纸质版;内容完整')
            worksheet.write(count, 5, page_num)

            # 获取pdf页数后关闭文件
            pdfDoc.close()

            # 对pdf文件进行重命名
            # used_name = pdfPath + '\\' + file  # 文件现有名
            # 文件新名字  学校代码-年份-学科代码.pdf
            # new_name =  pdfPath + '\\' + info[1] + '-' + info[2] + '-' + file_name[0:3] + '.pdf'
            # 重命名
            # os.rename(used_name, new_name)

        progress = Toplevel(root)
        progress.title('转换进度')
        progress.geometry('600x300')
        MyFrm(progress)  # 窗口居中显示

        # 更新进度条函数
        '''def change_schedule(now_schedule, all_schedule):
            canvas.coords(fill_line, (0, 0, (now_schedule / all_schedule) * 100, 60))
            progress.update()
            x.set(str(round(now_schedule / all_schedule * 100, 2)) + '%')
            if round(now_schedule / all_schedule * 100, 2) == 100.00:
                x.set("完成")'''

        # 设置转换进度
        #Label(progress, text='转换进度').place(relx=0.2, rely=0.5, anchor=CENTER)
        canvas = Canvas(progress, width=450, height=22, bg="white")
        canvas.place(relx=0.5, rely=0.5, anchor=CENTER)

        '''x = StringVar()
        # 进度条以及完成程度
        #out_rec = canvas.create_rectangle(5, 5, 105, 25, outline="green", width=1)
        fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill="green")
        Label(progress, textvariable=x).place(relx=0.6, rely=0.5, anchor=CENTER)'''




        # 输出查找总数
        print('真题pdf文件数：' + str(count))
        print('获取pdf信息表已完成，接下来将进入pdf2images...')
        # 关闭excel表格才会生成新表  重命名文件-获取真题pdf文件信息完成
        workbook.close()

        # 开始批量转换图片

        # 载入文件列表
        file_list = os.listdir(pdfPath)
        # 文件排序
        file_list.sort()
        # 转换的图片文件夹个数
        num = 0
        for file in file_list:
            pdf = pdfPath + '\\' + file  # pdf文件

            pdf_name = os.path.basename(pdf)  # 获取文件名字

            paper = pdf_name.split('.')[0]  # 去除pdf后缀，获取名字

            sub_code = paper[0:3]

            # 以学校代码/年份/学科代码/图片   保存图片路径(文件夹)
            imagePath = pdfPath + '\\' + sch_code + '\\' + date + '\\' + sub_code

            # 以学校代码-年份-学科代码.png命名图片
            imageName = sch_code + '-' + date + '-' + sub_code

            pyMuPDF_fitz(pdf, imagePath, imageName)  # 开始批量转换图片并添加水印


            # 填充进度条
            fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill="green")
            #n = 465 / x  # 465是矩形填充满的次数

            pdf_num = StringVar()
            Label(progress, textvariable=pdf_num).place(relx=0.5, rely=0.2, anchor=CENTER)
            pdf_num.set('当前文件总数:' + str(count))

            percent = StringVar()
            Label(progress, textvariable=percent).place(relx=0.5, rely=0.3, anchor=CENTER)



            n = 0
            num = num + 1

            if num <= count:
                percent.set('正在转化中:' + str(round(num / count * 100, 2)) + '%')

                n = n + round(num/count, 2) * 450
                canvas.coords(fill_line, (0, 0, n, 100))
                progress.update()
                time.sleep(0.02)  # 控制进度条流动的速度


                if round(num/count * 100, 2) == 100.00:
                    percent.set("太棒啦，全部转化完毕了哦~")

                    theButton = Button(progress, text="再来一次", width=10, command=progress.destroy)
                    theButton.place(relx=0.5, rely=0.7, anchor=CENTER)



            # 清空进度条
            '''fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill="white")
            x = count  # 未知变量，可更改
            n = 465 / x  # 465是矩形填充满的次数

            if num < x:
                n = n + 465 / x
                # 以矩形的长度作为变量值更新
                canvas.coords(fill_line, (0, 0, n, 60))
                progress.update()
                time.sleep(0)  # 时间为0，即飞速清空进度条'''




        print('亲，已全部转化成功了哦！！')
        print('本次转换的pdf个数为：' + str(num))





    theButton1 = Button(root, text="确认", width=10, command=Info)
    theButton2 = Button(root, text="清空", width=10, command=clear)
    theButton0 = Button(root, text="退出", width=10, command=root.destroy)  # app.quit是退出IDLE里冲突不能执行
    theButton1.place(relx=0.45, rely=0.5, anchor=CENTER)
    theButton2.place(relx=0.55, rely=0.5, anchor=CENTER)
    theButton0.place(relx=0.65, rely=0.5, anchor=CENTER)





    root.mainloop()

