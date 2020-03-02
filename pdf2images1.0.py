import re
import os
import sys, fitz
import xlwt
import openpyxl
import xlsxwriter
import datetime
from os import listdir

from PIL import Image, ImageDraw, ImageFont

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
        zoom_x = 1.33333333  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 1.33333333
        # 缩放系数都为2，分辨率提高4倍
        #zoom_x = 2  # (1.33333333-->1056x816)   (2-->1584x1224)
        #zoom_y = 2

        #zoom_x = 1  # (1.33333333-->1056x816)   (2-->1584x1224)
        #zoom_y = 1
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)

        if not os.path.exists(imagePath):  # 判断存放图片的文件夹是否存在
            os.makedirs(imagePath)  # 若图片文件夹不存在就创建


        # 页码从0开始，水印都加一
        pix.writePNG(imagePath + '/' + imageName + '-%s.png' % str(int(pg) + 1))  # 将图片写入指定的文件夹内



        # 为图片添加水印
        imageInfo = Image.open(imagePath + '/' + imageName + '-%s.png' % str(int(pg) + 1))

        # 文字水印

        fontOne = ImageFont.truetype("‪C:\Windows\Fonts\simfang.ttf", 30)  # 本地字体文件

        draw = ImageDraw.Draw(imageInfo)
        #print(imageInfo.size)

        # 深灰色fill=(96, 96, 96)
        # 开始添加文字水印

        draw.text((imageInfo.size[0] / 2 - 20, imageInfo.size[1] / 2 + 500), u"第%s页"%str(int(pg) + 1), fill=(0, 0, 0), font=fontOne)

        # imageInfo.show()  #展示图片

        # 添加图片水印
        logo = Image.open("C:\\logo.png")

        layer = Image.new('RGBA', imageInfo.size, (255, 255, 255, 0))

        # 添加图片水印

        layer.paste(logo, (imageInfo.size[0] - logo.size[0] + 120, imageInfo.size[1] - logo.size[1] - 20))

        imageInfo = Image.composite(layer, imageInfo, layer)

        imageInfo.save(imagePath + '/' + imageName + '-%s.png' % str(int(pg) + 1))

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


if __name__ == "__main__":
    # 需要处理的pdf文件夹路径（绝对路径）
    # pdfPath = "D:\\Pycharm\\python\\paper\\2016"
    # D:\\Pycharm\\python\\paper\\pdf\\2019     D:\\paper\\2019
    info = input("请输入学校+学校代码+年份(如[中山大学+4144010558+2019],并以Enter键结束):\n")
    pdfPath = input("请输入真题pdf文件所在路径(以Enter键结束):\n")
    # 存放结果信息的文件路径（绝对路径）
    # fpath = "D:\\Pycharm\\python\\paper.txt"
    # 载入文件列表
    file_list = os.listdir(pdfPath)
    # 文件排序
    file_list.sort()
    # 指定字符串总数
    count = 0

    # 中山大学+4144010558+2019
    info = info.split('+')   # 切割学校信息
    # 学校名称
    sch_name = info[0]
    # 学校代码
    sch_code = info[1]
    # 真题年份
    date = info[2]
    # 定义excel表名
    workbook = xlsxwriter.Workbook(sch_name + date + '真题信息表' + '.xlsx')
    # 创建一个空表
    worksheet = workbook.add_worksheet()

    # 遍历所有文件，对pdf文件进行重命名
    for file in file_list:
        count = count + 1
        pdfpath = pdfPath + '\\' + file  # pdf文件
        file_name = os.path.basename(pdfpath)  # 获取文件名字

        name = file_name.split('.')[0]  # 去除后缀，获取名字


        pdfDoc = fitz.open(pdfpath)   # 读取页码数
        page_num = pdfDoc.pageCount   # 页码



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
        #used_name = pdfPath + '\\' + file  # 文件现有名
        # 文件新名字  学校代码-年份-学科代码.pdf
        #new_name =  pdfPath + '\\' + info[1] + '-' + info[2] + '-' + file_name[0:3] + '.pdf'
        # 重命名
        #os.rename(used_name, new_name)


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

        num = num + 1

    print('亲，已全部转化成功了哦！！')
    print('本次转换的pdf个数为：' + str(num))



