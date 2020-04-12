from sympy import *
import xlsxwriter

workbook = xlsxwriter.Workbook('简单迭代法结果.xlsx') # 建立文件
worksheet = workbook.add_worksheet()#建立sheet
worksheet.write(0,0,'简单迭代法计算非线性方程的根')

i = 3#从第三行开始输出
x = symbols("x")
y1_ = (3 * (x ** 3)) / (3 * (x ** 5) + 2 * (x ** 4) - (x ** 3) + 5) - 1 / (4 * x) + 20
y2_1, y2_2 = pow((1 / (x + 5) + (4 * (x ** 2))) + (4 * x) - 1, 1 / 3), (1/4)-((x**3)/4)-(1/(4*(x+5)))-(x**2)
y3_1, y3_2, y3_3, y3_4 = x-x**3-4*(x**2)+10, pow((10/x)-(4*x),0.5), 0.5*pow(10-x**3,0.5), pow(10/(4+x),0.5)
y4_1, y4_2, y4_3 = pow(2*x+5,1/3), (x**3-5)/2, pow((5/x)+2,0.5)
#每题可能或给定的公式

def diedai(result):#定义迭代功能函数，实现代码复用
    y, value, cishu, panduan =result#分别表示函数表达式，试探值，函数运行次数，置信区间
    global i
    n = y.subs(x, value)#带入得新的函数值
    worksheet.write(i, 0, str(y))
    worksheet.write(i, 2, float(n))
    worksheet.write(i, 1, cishu)
    i += 1
    if abs(float(n)-float(value)) > 1000:#判断是否所得值与上一值是否相差过大，即发散
        worksheet.write(i, 0, '该公式发散')
        i += 1
        return 0
    if (float(n) >= float(value)-panduan) and ((float(n) <= float(value)+panduan)):#判断是否在置信区间中
        kkk = abs(y.subs(x, n)-float(n))
        worksheet.write(i, 0, '迭代次数为：{0}，x的值为{1}，误差大小为{2}'.format(cishu, float(n),float(kkk)))
        i += 1#若是则计算次数，得出答案
        return 0
    else:
        cishu += 1#若否则进行迭代
        result = [y, n, cishu, panduan]#将表达式不变，试探值为新值，次数加一，置信区间不变，传回函数进行迭代
        return diedai(result)

def mds(abc,ccc):
    global i
    worksheet.write(i, 0, '公式{}'.format(ccc))
    i += 1
    try:
        diedai(abc)#代入初始值
    except OverflowError:
        worksheet.write(i, 0, '该公式发散')
        i += 1
    except TypeError:#判断是否出现虚数
        worksheet.write(i, 0, '出现虚数')
        i += 1

if __name__ == '__main__':#主函数入口
    worksheet.write(1, 0, '迭代公式')
    worksheet.write(1, 1, '迭代次数')
    worksheet.write(1, 2, '值')#打印表头
    i += 1
    worksheet.write(i, 0, '第一题')#输出题号
    i += 2
    mds([y1_, 10, 1, 0],1)
    i += 1
    worksheet.write(i, 0, '第二题')#输出题号
    i += 2
    mds([y2_1, 10, 1, 0],2)
    mds([y2_2, 10, 1, 0],3)
    i += 1
    worksheet.write(i, 0, '第三题')#输出题号
    i += 2
    mds([y3_1, 1.5, 1, 10**(-8)],1)
    mds([y3_2, 1.5, 1, 10**(-8)],2)
    mds([y3_3, 1.5, 1, 10 ** (-8)],3)
    mds([y3_4, 1.5, 1, 10**(-8)],4)
    i += 1
    worksheet.write(i, 0, '第四题')#输出题号
    i += 2
    mds([y4_1, -1.5, 1, 10**(-8)],1)
    mds([y4_2, 1.5, 1, 10**(-8)],2)
    mds([y4_3, -1.5, 1, 10**(-8)],3)
    workbook.close()