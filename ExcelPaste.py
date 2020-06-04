# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
import xlwt
import pyperclip
from pynput import mouse, keyboard
import threading
import sys
import re
from openpyxl import Workbook, load_workbook


def copy():
    k.press_key(k.control_l_key)
    k.tap_key("c")  # 改小写！！！！ 大写的话由于单进程会触发shift键 ctrl键就失效了
    k.release_key(k.control_l_key)


def getCopy(maxTime=1.2):
    # maxTime = 3  # 3秒复制 调用copy() 不管结果对错
    while (maxTime > 0):
        maxTime = maxTime - 0.1
        time.sleep(0.1)
        # print('doing')
        copy()

    result = pyperclip.paste()
    return result


def tapkey(key, count=1, waitTime=0.2):
    for i in range(0, count):
        k.tap_key(key)
        time.sleep(waitTime)
#暂定excel使用


def Quit():
    global end
    end = True


def Do():
    global lastCo
    global zhucais
    global start
    if start:
        #dodoododododododoodododododoodododododoododododododododo

        #print(zhucais)
        datas = []
        excel = xlrd.open_workbook(excelUrl)
        table = excel.sheets()[0]
        rowCount = table.nrows
        colCount = table.ncols

        for i in range(rowCount):
            value= str(table.cell_value(i, 0)).split('$')  # 不锈钢$地漏
            if('正确答案是' in value[0]):
                continue
            if(value!=['']):
                print(value)

                wait=0.02
                targetStr=str(value[0])

                if(len(str(targetStr))>15):
                    wait=0.5
                while(len(targetStr)>150):
                    # 粘贴中文
                    pyperclip.copy(targetStr[:150])
                    time.sleep(0.05+wait)
                    k.press_key(k.control_key)
                    k.tap_key('v')
                    k.release_key(k.control_key)
                    # 粘贴中文
                    time.sleep(0.05+wait)
                    tapkey(k.enter_key,1,0)
                    tapkey(k.down_key,1,0)
                    tapkey(k.escape_key,1,0)
                    tapkey(k.left_key,2,0)
                    tapkey(k.enter_key,1,0)
                    time.sleep(0.05+wait)

                    targetStr= targetStr.replace(targetStr[:150],'')
                else:
                    pyperclip.copy(targetStr)
                    time.sleep(0.05 + wait)
                    k.press_key(k.control_key)
                    k.tap_key('v')
                    k.release_key(k.control_key)
                    # 粘贴中文
                    time.sleep(0.05 + wait)
                    tapkey(k.enter_key, 1, 0)
                    tapkey(k.down_key, 1, 0)
                    tapkey(k.escape_key, 1, 0)
                    tapkey(k.left_key, 2, 0)
                    tapkey(k.enter_key, 1, 0)
                    time.sleep(0.05 + wait)


        start=False
        end=True






# 我的代码
def onpressed(Key):
    while True:
        # print(Key)
        if (Key == keyboard.Key.caps_lock):  # 开始
            global start
            if (start == True):
                start = False
                print('stop')
            else:
                start = True
                print('go')
        if (Key == keyboard.Key.f3):  # 结束
            sys.exit()

        global end
        if (end):
            sys.exit()
        return True


def main():
    while True:
        # 主程序在这
        Do()


if __name__ == '__main__':
    k = PyKeyboard()
    m = PyMouse()
    end = False
    start = False
    excelUrl = r"C:\Users\Administrator\Desktop\TempA.xlsx"#to do-------------

    excel = xlrd.open_workbook(excelUrl)
    table = excel.sheets()[0]
    rowCount = table.nrows
    print('需要这么多行:')
    print(rowCount)
    print('记得手动关！！！！')

    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    print('press Capital to start')

    with keyboard.Listener(on_press=onpressed) as listener:
        listener.join()

