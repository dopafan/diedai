from sympy import *
import math
import xlsxwriter

workbook = xlsxwriter.Workbook('牛顿迭代法结果.xlsx') # 建立文件
worksheet = workbook.add_worksheet()#建立sheet
worksheet.write(0,0,'牛顿迭代法计算实现线性方程组求根')

i = 3#从第三行开始输出
x = symbols("x")
y1_ = x ** 3 - 2 * x - 5
y2_ = x**3-x-1
y3_ = x**3+4*(x**2)-10
y4_ = x*(exp(x))-1
y5_ = (exp(x))#定义题目给出的函数表达式

def diedai(result):#定义迭代功能函数，实现代码复用
    y, value, cishu, panduan =result#分别表示函数表达式，试探值，函数运行次数，置信区间
    global i
    a = float(y.subs(x, value))#带入得f（x）的函数值
    n = (value - (a) / (diff(y, x).subs(x, value)))#x（n+1）=（x（n）-f（x）/f‘（x）
    worksheet.write(i, 0, str(y))
    worksheet.write(i, 1, float(n))
    worksheet.write(i, 2, cishu)
    i += 1
    if (float(n) >= float(value)-panduan) and ((float(n) <= float(value)+panduan)):#判断是否在置信区间内
        worksheet.write(i, 0, '迭代次数为：{0}，x的值为{1}'.format(i, float(n)))
        i += 1
        return 0
    else:
        cishu += 1
        result = [y, n, cishu, panduan]#将表达式不变，试探值为新值，次数加一，置信区间不变，传回函数进行迭代
        return diedai(result)

def tianchong(ju):#定义打印表头函数
    global i
    worksheet.write(i, 0, '第{}题'.format(ju))#输出题号
    i += 1
    worksheet.write(i, 0, '迭代公式')#表头
    worksheet.write(i, 1, '值')#表头
    worksheet.write(i, 2, '迭代次数')#表头
    i += 1

if __name__ == '__main__':
        tianchong(1)
        diedai([y1_, 1, 1, 0])
        tianchong(2)
        diedai([y2_, 1.5, 1, 0])
        tianchong(3)
        diedai([y3_, 1.5, 1, 10**(-4)])
        tianchong(4)
        diedai([y4_, 0, 1, 10**(-5)])
        tianchong(5)
        diedai([y5_, 0.5, 1, 10**(-5)])
        workbook.close()