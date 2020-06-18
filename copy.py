# -*- coding: utf-8 -*-

from ctypes import *
from ctypes import wintypes
import win32gui
import win32con
import win32api
from time import sleep

sleep(5)
#调用Windows系统动态链接库user32.dll

user32 = windll.user32

p = wintypes.POINT()

buffer = create_string_buffer(255)

#获取鼠标位置

user32.GetCursorPos(byref(p))

#获取鼠标所处位置的窗口句柄

hwnd = user32.WindowFromPoint(p)
print(hwnd)
text = win32gui.GetWindowText(hwnd)
# print(text)
#注释掉的代码本来是可以实现星号密码查看的，在Win7以后的系统中失效了

#dwStyle = user32.GetWindowLongA(hwnd, -16) #-16是GWL_STYLE消息的值

#user32.SetWindowWord(hwnd, -16, 0)


#获取窗口文本

user32.SendMessageA(hwnd, win32con.WM_GETTEXT, 255, byref(buffer)) #13是WM_GETTEXT消息的值

#user32.SetWindowLongA(hwnd, -16, dwStyle)

print(buffer.value.decode('gbk'))
sleep(100)
