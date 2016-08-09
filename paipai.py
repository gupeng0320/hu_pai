# -*- coding: utf-8 -*-
import os, sys
from selenium import webdriver
import time, datetime
from PIL import Image
#import pytesser
import pytesseract
from PIL import ImageGrab
from PIL import ImageFilter
from PIL import ImageEnhance
import pyautogui
import pyHook
import pythoncom
import win32gui
import time,httplib
import ctypes
import subprocess
import re

global is_finish_yanzhengma
global is_ie_already_run
is_ie_already_run = False
global j

class PaipaiMgr():
    def __init__(self):
        self.iedriver = "d:\softwares\Python27\IEDriverServer.exe"
        self.whole_pic = r'e:\projects\paipai\all.png'
        self.time_pic = r'e:\projects\paipai\time.bmp'
        self.price_pic = r'Ce:\projects\paipai\price.bmp'
        self.input_price_time = "11:29:10"
        self.time_rangle = (135, 525, 210, 540)

        self.position_delta_price = [780, 440]
        self.position_max_price = [333, 556]
        self.position_my_price = [727, 380]
        self.position_add_price = [850, 430]
        self.position_submit_price = [850, 540]
        #self.position_ok_after_submit_price = [580, 610]
        self.position_ok_after_submit_price = [600, 620]
        self.position_refresh_yanzhengma= [580, 550]
        #self.position_cancel_after_submit_price= [790, 610]
        self.position_cancel_after_submit_price = [790, 620]
        self.position_ok_after_accept_price = [700, 600]

        self.current_max_price_rangle = (self.position_max_price[0], self.position_max_price[1], self.position_max_price[0] + 42, self.position_max_price[1] + 15)
        self.submit_price_rangle = (self.position_my_price[0], self.position_my_price[1], self.position_my_price[0]+51, self.position_my_price[1] + 25)

        self.ie_title = "51沪牌模拟拍牌系统 - Windows Internet Explorer"
        self.driver = ''
        #record which turn to submit the price, cus different strategies
        self.price_submit_shoots = 0
        self.regex_str = r"\w+\s+(\d+)\s+"
        self.regex_obj = re.compile(self.regex_str)

        '''
        Ocr.set_up() # one time setup
        self.ocrEngine = Ocr()
        #self.ocrEngine.input_license("123456", "123456789123456789123456789")
        self.ocrEngine.start_engine("eng")
        '''
        self.start()

    def test_mouse_position(self):
        while 1:
            self.add_price("700")
            print "delta price"
            pyautogui.moveTo(self.position_delta_price[0], self.position_delta_price[1])
            time.sleep(1)
            print "max price"
            pyautogui.moveTo(self.position_max_price[0], self.position_max_price[1])
            time.sleep(1)
            print "add price"
            pyautogui.moveTo(self.position_add_price[0], self.position_add_price[1])
            pyautogui.click()
            time.sleep(1)
            print "submit price"
            pyautogui.moveTo(self.position_submit_price[0], self.position_submit_price[1])
            pyautogui.click()
            time.sleep(1)
            print "position_ok_after_submit_price"
            pyautogui.moveTo(self.position_ok_after_submit_price[0], self.position_ok_after_submit_price[1])
            time.sleep(1)
            print "position_my_price"
            pyautogui.moveTo(self.position_my_price[0], self.position_my_price[1])
            time.sleep(1)
            print "position_refresh_yanzhengma"
            pyautogui.moveTo(self.position_refresh_yanzhengma[0], self.position_refresh_yanzhengma[1])
            pyautogui.click()
            time.sleep(1)
            print "position_cancel_after_submit_price"
            pyautogui.moveTo(self.position_cancel_after_submit_price[0], self.position_cancel_after_submit_price[1])
            pyautogui.click()

            time.sleep(2)

    def start(self):
        '''
        window_name = u'test_price.bmp'
        hwnd =win32gui.FindWindow("IEFrame", None)
        #ret =win32gui.FindWindow(None, None)
        print hwnd
        #print "--- title is %s, name is %s, hwnd is %s" % (win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd), hwnd)
        print "--- name is %s, hwnd is %s" % (win32gui.GetClassName(hwnd), hwnd)
        print "--- text  is %s, hwnd is %s" % ((win32gui.GetWindowText(hwnd).decode('gbk').encode('utf-8')), hwnd)
        time.sleep(2)
        '''
        tasks = subprocess.Popen("tasklist", stdin = subprocess.PIPE, stdout=subprocess.PIPE, shell=True)
        #print tasks.stdout.readlines()

        #IEDriverServer.exe
        print "my pid is :", os.getpid()
        for line in tasks.stdout.readlines():
            if "python.exe" in line:
                print line
                if str(os.getpid()) not in line:
                    result = self.regex_obj.search(line)
                    print "clean environment, kill existed python process:%s" % result.group(1)
                    os.system("taskkill /pid %s" % result.group(1))
                    os.system("taskkill /F /IM IEDriverServer.exe")

        win32gui.EnumWindows(self.validate_ie, 0)

        if is_ie_already_run != True:
            print "restart web driver IE to paipai"
            os.environ["webdriver.ie.driver"] = self.iedriver
            self.driver = webdriver.Ie(self.iedriver)
            self.driver.get("http://moni.51hupai.org/?new=13")
            self.driver.set_window_size(1000,930)
            time.sleep(2)

        #hwnd = win32gui.WindowFromPoint((300, 300))
        #win32gui.EnumWindows(show_windows, 0)

    def add_price(self, price):
        pyautogui.moveTo(self.position_delta_price[0], self.position_delta_price[1])
        pyautogui.click()
        pyautogui.typewrite(chr(8) * 10 + price)
        

    def get_current_price(self, image, rangle):
        #price_part = image.crop(rangle).convert('1')
        price_part = image.crop(rangle).convert('L')
        price_part = price_part.resize((240, 120))
        price_part = price_part.convert("RGBA")
        #print price_part.format, price_part.size, price_part.mode, price_part.info

        for i in xrange(0,1):
            price_part = price_part.filter(ImageFilter.EDGE_ENHANCE)
            #price_part = price_part.filter(ImageFilter.SMOOTH)
            #price_part = price_part.filter(ImageFilter.SMOOTH_MORE)
            price_part = price_part.filter(ImageFilter.SMOOTH_MORE)

            #price_part = price_part.filter(ImageFilter.DETAIL)
            #price_part = price_part.filter(ImageFilter.SHARPEN)
            price_part = ImageEnhance.Contrast(price_part)
            #price_part = ImageEnhance.Brightness(price_part)
            #price_part = ImageEnhance.Sharpness(price_part)

            price_part =price_part.enhance(4)
            #price_part.save(path, 'bmp')
        #price_part.show()

        #print("tessact price start %s" % datetime.datetime.now())
        price = pytesseract.image_to_string(price_part, lang="num", config="-psm 7").strip().replace(' ', '')
        #print("tessact price end %s" % datetime.datetime.now())
        #print("current price is: %s\n" % price)
        #self.price_pic = r'e:\projects\paipai\price_%s.png' % i
        #price_part.save(r'e:\projects\code\hu_pai\price.bmp', 'bmp')
        '''
        price_part.save(self.price_pic, 'bmp')
        #price_part.show()
        #print("aprise price start %s" % datetime.now())
        price = self.ocrEngine.recognize(self.price_pic, -1, -1, -1, -1, -1, OCR_RECOGNIZE_TYPE_ALL, OCR_OUTPUT_FORMAT_PLAINTEXT)
        #print("aprise price end %s" % datetime.now())
        #print("aprise price recognize Result: %s" % price)
        '''
        return price.strip().replace(' ', '')

    def get_current_time(self, image, i):
        #print("aprise time start %s" % datetime.now())
        time_part = image.crop(self.time_rangle).convert('L')
        #time_part = image.crop(self.time_rangle).convert('1')
        #time_part = image.crop(self.time_rangle).convert('RGB')
        time_part = time_part.resize((360, 150))
        time_part = time_part.convert("RGBA")
        time_part = time_part.filter(ImageFilter.EDGE_ENHANCE)
        time_part = time_part.filter(ImageFilter.SMOOTH_MORE)
        time_part = ImageEnhance.Contrast(time_part)
        time_part = time_part.enhance(4)
        #self.time_pic = r'e:\projects\paipai\time_%s.png' % i

        #time_part.save(self.time_pic,'bmp')
        #print time_part.format, time_part.size, time_part.mode
        #time_part=time_part.convert("RGB")
        #print time_part.format, time_part.size, time_part.mode

        #s = self.ocrEngine.recognize(r"e:\projects\paipai\time_99.png", -1, -1, -1, -1, -1, OCR_RECOGNIZE_TYPE_ALL, OCR_OUTPUT_FORMAT_PLAINTEXT)
        time_now = self.ocrEngine.recognize(self.time_pic, -1, -1, -1, -1, -1, OCR_RECOGNIZE_TYPE_ALL, OCR_OUTPUT_FORMAT_PLAINTEXT)
        #print("aprise end %s" % datetime.now())
        #print("aprise recognize Result: %s" % time_now)
        '''

        print("tessact time start %s" % datetime.now())
        time_now = pytesseract.image_to_string(time_part).strip()
        print("tessact time end %s" % datetime.now())
        #print("tessact recognize Result: %s" % time_now)
        '''
        #time_part.show()
        #print("current time is: ", time_now)
        return time_now.strip().replace(' ', '')

    def execute(self):
        self.add_price("800")
        i = 0
        while 1:
            i = i + 1
            im = ImageGrab.grab()
            #im = Image.open(r"e:\projects\paipai\test_price.bmp")
            #print("111 %s" % datetime.datetime.now())
            #print(pytesseract.image_to_string(im, lang="num"))
            #print("222 %s" % datetime.datetime.now())
            current_time = "11:29:50"
            #current_time = self.get_current_time(im, i)
            #print("current time is: %s, %s" %(current_time, type(current_time)))
            #current_time = getBeijinTime()
            #print "current_time : %s" % current_time

            time_start = datetime.datetime.now()
            print "time now: %s" % (time_start)
            time_now = time_start.strftime("%H:%M:%S")
            [hour_now, minute_now, second_now] = time_now.split(":")

            price = self.get_current_price(im, self.current_max_price_rangle)
            print"current price is:     [====    %s    ====]\n" % price

            self.input_price_time = "11:29:47"
            #if time_now == self.input_price_time:
            if 0 == self.price_submit_shoots and second_now[1] == "8":
                print(self.input_price_time, ": input price!!")

                #add price and submit
                pyautogui.moveTo(self.position_add_price[0], self.position_add_price[1])
                #print "test 111111111"
                pyautogui.click()
                #print "test"
                pyautogui.moveTo(self.position_submit_price[0], self.position_submit_price[1])
                pyautogui.click()

                #time.sleep(5)
                #yanzhengma = raw_input("input yanzhengma:")
                #print "yzm: %s" % yanzhengma
                #enter after input yanzhengma
                hm = pyHook.HookManager()
                hm.KeyDown = self.onKeyboardEvent
                hm.HookKeyboard()
                #pythoncom.PumpMessages()
                global is_finish_yanzhengma
                is_finish_yanzhengma = False
                while is_finish_yanzhengma == False:
                    pythoncom.PumpWaitingMessages()
                    #pythoncom.PumpMessages()

                print "user has finish inputing yanzhengma!!"

                im = ImageGrab.grab()
                submit_price = self.get_current_price(im, self.submit_price_rangle).replace(' ', '')
                print "submit_price:" , submit_price
                #pyautogui.moveTo(self.position_submit_price[0], self.position_submit_price[1])

                max_price = self.get_current_price(im, self.current_max_price_rangle).replace(' ', '')
                print "max_price:" , max_price

                while(int(submit_price) > int(max_price)):
                    time_now = datetime.datetime.now()
                    time_simple = time_now.strftime("%H:%M:%S")
                    print "(%s) submit price %s is bigger than current max price %s, please wait!\n" % (time_now, submit_price, max_price)
                    im = ImageGrab.grab()
                    max_price = self.get_current_price(im, self.current_max_price_rangle)
                    #if time_simple > "11:29:55":
                    if time_now > time_start + datetime.timedelta(seconds=8):
                        print "********************* time out (%s)!!! we need to submit anyway!!!\n" % datetime.datetime.now()
                        break

                print "#########################please submit!!!!!!!!!!!!!!!!!!"
                print "#########################please submit!!!!!!!!!!!!!!!!!!"

                pyautogui.moveTo(self.position_ok_after_submit_price[0], self.position_ok_after_submit_price[1])
                pyautogui.click()

                is_finish_yanzhengma = False
                while is_finish_yanzhengma == False:
                    pythoncom.PumpWaitingMessages()

                print "server accept price submit"
                #time.sleep(3)

                #enter after server admit my request
                pyautogui.moveTo(self.position_ok_after_accept_price[0], self.position_ok_after_accept_price[1])
                pyautogui.click()
                self.price_submit_shoots += 1
            elif 1 == self.price_submit_shoots:
                print "try to submit price the second time"
                self.add_price("")
                #add price and submit
                pyautogui.moveTo(self.position_add_price[0], self.position_add_price[1])
                pyautogui.click()
                pyautogui.moveTo(self.position_submit_price[0], self.position_submit_price[1])
                pyautogui.click()

                hm = pyHook.HookManager()
                hm.KeyDown = self.onKeyboardEvent
                hm.HookKeyboard()
                #pythoncom.PumpMessages()
                global is_finish_yanzhengma
                is_finish_yanzhengma = False
                while is_finish_yanzhengma == False:
                    pythoncom.PumpWaitingMessages()
                    #pythoncom.PumpMessages()

                print "user has finish inputing yanzhengma!!"

                im = ImageGrab.grab()
                submit_price = self.get_current_price(im, self.submit_price_rangle).replace(' ', '')
                print "submit:" , submit_price
                #pyautogui.moveTo(self.position_submit_price[0], self.position_submit_price[1])

                pyautogui.moveTo(self.position_ok_after_submit_price[0], self.position_ok_after_submit_price[1])
                pyautogui.click()

                is_finish_yanzhengma = False
                while is_finish_yanzhengma == False:
                    pythoncom.PumpWaitingMessages()

                print "server accept price submit"
                #time.sleep(3)
                #enter after server admit my request
                pyautogui.moveTo(self.position_ok_after_accept_price[0], self.position_ok_after_accept_price[1])
                pyautogui.click()


    def exit(self):
        if(self.driver):
            self.driver.quit()
        self.ocrEngine.stop_engine()

    def validate_ie(self, hwnd, mouse):
        #if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            #print "--- title is %s, name is %s, hwnd is %s" % (win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd), hwnd)
            #print "--- name is %s, hwnd is %s" % (win32gui.GetClassName(hwnd), hwnd)
            #print "--- text  is %s, hwnd is %s" % ((win32gui.GetWindowText(hwnd).decode('gbk').encode('utf-8')), hwnd)
            #if self.ie_title == win32gui.GetWindowText(hwnd).decode('gbk').encode('utf-8'):
            if win32gui.GetClassName(hwnd) == "Internet Explorer_Hidden" or "51沪牌模拟拍牌系统 - Internet Explorer" == (win32gui.GetWindowText(hwnd).decode('gbk').encode('utf-8')):
                global is_ie_already_run
                is_ie_already_run = True
                print "[INFO] name is %s, IE already run!" % ((win32gui.GetClassName(hwnd)))

    def onKeyboardEvent(self, event):
        '''
        print "MessageName:", event.MessageName
        print "Message:", event.Message
        print "Time:", event700
        print "Window:", event.W700
        print "WindowName:", event.WindowName
        print "Ascii:", event.Ascii, chr(event.Ascii)
        print "Key:", event.Key
        print "KeyID:", event.KeyID
        print "ScanCode:", event.ScanCode
        print "Extended:", event.Extended
        print "Injected:", event.Injected
        print "Alt", event.Alt
        print "Transition", event.Transition
        print "---"
        '''

        print "input Key:", event.Key
        global is_finish_yanzhengma
        if 13 == event.Ascii:
            is_finish_yanzhengma = True
            print "finish input yanzhengma!! %s" % is_finish_yanzhengma
        elif 27 == event.Ascii:
            print "refresh the identifying code"
            pyautogui.moveTo(self.position_refresh_yanzhengma[0], self.position_refresh_yanzhengma[1])
            pyautogui.click()
        #pythoncom.EnableQuitMessage()
        return True

def validate_mouse_position(event):

    print "MessageName:", event.MessageName
    print "Message:", event.Message
    print "Time:", event700
    print "Window:", event.W700
    print "WindowName:", event.WindowName
    print "Ascii:", event.Ascii, chr(event.Ascii)
    print "Key:", event.Key
    print "KeyID:", event.KeyID
    print "ScanCode:", event.ScanCode
    print "Extended:", event.Extended
    print "Injected:", event.Injected
    print "Alt", event.Alt
    print "Transition", event.Transition
    print "---"

    print "input Key:", event.Key
    global is_test_position_finish
    if 13 == event.Ascii:
        is_test_position_finish = True
        print "mouse position is right!! %s" % is_test_position_finish
        return True
    elif 14 == event.Ascii:
        is_test_position_finish = False
        print "mouse position is wrong!! need to be corrected %s" % is_test_position_finish
        return False

if __name__ == '__main__':
    print "test python here"

    #print("%s" % datetime.datetime.now())
    #os.system("tesseract.exe e:\projects\paipai\price_47.png result -l num")
    #print("%s" % datetime.datetime.n700
    #AspriseOCR = ctypes.windll.LoadLibrary(r":\gupeng\software\Python27\Lib\site-packages\asprise_ocr_api\AspriseOcr.dll")

    paipai = PaipaiMgr()
    #paipai.test_mouse_position()
    paipai.execute()
    paipai.exit()
