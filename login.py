# -*- coding: utf-8 -*-
# @Time : 2019/8/1 16:24
# @Author : zjl
# @Site :
# @File : qq登录.py
import os
import time
import win32gui
import win32api
import win32con
import pyautogui
from ctypes import *
import redis
import json
import aircv as ac
from PIL import Image, ImageGrab
import requests


class QQ_login(object):
    def __init__(self, username=None, password=None):
        # 运行tim
        # os.system('"C:\qq\Bin\QQScLauncher.exe"')
        self.conn = redis.StrictRedis.from_url('redis://192.168.17.109/15')
        self.lock_user = 'tencent_selenium_accounts_lock_2'
        os.system('"C:\qq\Bin\QQScLauncher.exe"')
        time.sleep(5)
        a = win32gui.FindWindow(None, "TIM")  # 获取窗口的句柄，参数1: 类名，参数2： 标题QQ
        self.loginid = win32gui.GetWindowPlacement(a)  # 获取的是窗口的坐标
        self.username = username
        self.password = password
        self.Tim_img = './img/Tim.png'
        self.yanzheng_img = './img/yanzheng.png'
        self.full_img = './img/full_img.png'
        self.flag_img = './img/flag.png'
        self.flag_login = './img/flag_login.jpg'

    def Tim_login(self):
        print u'窗口的坐标是{},{},{},{}'.format(self.loginid[4][0], self.loginid[4][1], self.loginid[4][2],
                                          self.loginid[4][3])
        print u'点击的坐标是{},{}'.format(self.loginid[4][0] + 300, self.loginid[4][1] + 273)
        windll.user32.SetCursorPos(self.loginid[4][0] + 300, self.loginid[4][1] + 273)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)  # 按下鼠标
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)  # 放开鼠标
        time.sleep(0.2)
        ###输入账号
        # SendKeys.SendKeys(qq)
        print u'输入账号：{}'.format(self.username)
        for _ in range(10):
            win32api.keybd_event(8, 0, 0, 0)
            win32api.keybd_event(8, 0, win32con.KEYEVENTF_KEYUP, 0)
        pyautogui.typewrite(self.username, interval=0.3)
        time.sleep(0.2)
        ##tab切换
        print u'输入密码:{}'.format(self.password)

        win32api.keybd_event(9, 0, 0, 0)
        win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
        for _ in range(10):
            win32api.keybd_event(8, 0, 0, 0)
            win32api.keybd_event(8, 0, win32con.KEYEVENTF_KEYUP, 0)
        pyautogui.typewrite(self.password, interval=0.3)
        # 点击回车键登录
        time.sleep(0.5)
        print u'点击登录'
        win32api.keybd_event(13, 0, 0, 0)
        win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)

    def jietu(self, type=None):
        win32api.keybd_event(win32con.VK_SNAPSHOT, 0)
        time.sleep(0.5)
        full_img = ImageGrab.grabclipboard()
        full_img.save(self.full_img)
        if type == 1:
            print u'截图查看是否需要输入验证码'
            bbox = (self.loginid[4][0], self.loginid[4][1], self.loginid[4][2], self.loginid[4][3])
            img = Image.open(self.full_img)
            # print(img.size)
            cropped = img.crop(bbox)  # (left, upper, right, lower)
            cropped.save(self.Tim_img)
        elif type == 2:
            print u'需要输入验证码,获取验证码进行截图'
            bbox = (
                self.loginid[4][0] + 121, self.loginid[4][1] + 201, self.loginid[4][2] - 242, self.loginid[4][3] - 213)
            img = Image.open(self.full_img)
            # print(img.size)
            cropped = img.crop(bbox)  # (left, upper, right, lower)
            cropped.save(self.yanzheng_img)

    def get_door(self):
        resp = requests.post("http://127.0.0.1:7788", data=open(self.yanzheng_img, "rb"))
        code = json.loads(resp.text)["code"]
        return code

    def door_login(self, door):
        # win32api.keybd_event(16, 0, 0, 0)
        # win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)
        windll.user32.SetCursorPos(self.loginid[4][0] + 276, self.loginid[4][1] + 182)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)  # 按下鼠标
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)  # 放开鼠标
        time.sleep(0.2)
        for _ in range(10):
            win32api.keybd_event(8, 0, 0, 0)
            win32api.keybd_event(8, 0, win32con.KEYEVENTF_KEYUP, 0)
        pyautogui.typewrite(door, interval=0.3)
        time.sleep(0.2)
        # 点击回车键登录
        win32api.keybd_event(13, 0, 0, 0)
        win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(3)

    def find_img(self):

        imsrc = ac.imread(self.Tim_img)
        imobj = ac.imread(self.flag_img)
        # find the match position
        pos = ac.find_template(imsrc, imobj)
        if pos:
            return {'code': 2, 'desc': '需要输入验证码'}
        # -------------
        imobj = ac.imread('./img/login_lock.png')
        pos = ac.find_template(imsrc, imobj)
        if pos:
            return {'code': 3, 'desc': '账号被锁定'}
        # -------------
        imobj = ac.imread('./img/login_wait.png')
        pos = ac.find_template(imsrc, imobj)
        if pos:
            return {'code': 1, 'desc': '正在加载页面，请稍后'}
        # -------------
        imobj = ac.imread(self.flag_login)
        pos = ac.find_template(imsrc, imobj)
        if pos:
            return {'code': 1, 'desc': '正在登录中'}
        return {'code': 0, 'desc': '账号登录成功'}

    def run(self):
        # 1.输入账号，密码，登录
        self.Tim_login()
        # 2.点击登录后，看页面跳转是否需要验证
        flag = 0
        while flag <= 5:
            self.jietu(type=1)
            resp = self.find_img()
            code = resp.get('code')
            if code == 0:
                print '登录成功'
                return
            elif code == 1:
                print '正在登录中，请等待'
                time.sleep(2)
                flag += 1
            elif code == 3:
                print '账号被锁定'
                info = {
                    "username": self.username,
                    "password": self.password
                }
                self.conn.lpush(self.lock_user, json.dumps(info))
                return
            else:
                # 3.需要输入验证码，验证码截图，获取验证码
                self.jietu(type=2)
                code = self.get_door()
                self.door_login(code)
                flag += 1


if __name__ == '__main__':
     
        username = 'username'
        password = 'password'
        QQ_login(username, password).run()
        time.sleep(5)
        print '-----------------------------'
        
