from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from openpyxl import workbook ,load_workbook ,Workbook
import os ,time ,random ,ytFuntion
#0816,periodDetail sleep 1s
testdayFile = time.strftime("%y_%m_%d")
testdayTime  = time.strftime("%y_%m_%d_%H_%M_%S")

funtionError = []
funtionCountPng = 1

class accountSetting():
    def __init__(self ,username = "" ,password = "" ,safePassword = ""):
        self.username = str(username).strip()
        self.password = str(password).strip()
        self.safePassword = str(safePassword).strip()

class LocalStorage():
    def __init__(self ,webDriver):
        self.webDriver = webDriver    

    def __len__(self):
        return self.webDriver.execute_script("return window.localStorage.length;")

    def items(self) :
        return self.webDriver.execute_script( \
            "var ls = window.localStorage, items = {}; " \
            "for (var i = 0, k; i < ls.length; ++i) " \
            "  items[k = ls.key(i)] = ls.getItem(k); " \
            "return items; ")

    def keys(self) :
        return self.webDriver.execute_script( \
            "var ls = window.localStorage, keys = []; " \
            "for (var i = 0; i < ls.length; ++i) " \
            "  keys[i] = ls.key(i); " \
            "return keys; ")

    def get(self, key):
        return self.webDriver.execute_script("return window.localStorage.getItem(arguments[0]);", key)

    def set(self, key, value):
        self.webDriver.execute_script("window.localStorage.setItem(arguments[0], arguments[1]);", key, value)

    def has(self, key):
        return key in self.keys()

    def remove(self, key):
        self.webDriver.execute_script("window.localStorage.removeItem(arguments[0]);", key)

    def clear(self):
        self.webDriver.execute_script("window.localStorage.clear();")

    def __getitem__(self, key) :
        value = self.get(key)
        if value is None :
          raise KeyError(key)
        return value

    def __setitem__(self, key, value):
        self.set(key, value)

    def __contains__(self, key):
        return key in self.keys()

    def __iter__(self):
        return self.items().__iter__()

    def __repr__(self):
        return self.items().__str__()
    
class test_web(LocalStorage):
    def __init__(self ,webDriver):
        super().__init__(webDriver)

    def webUrl(self):
        return self.webDriver.current_url

    def webPageSelect(self ,i=" "):
        i = str(i).strip()
        if i == "" or i == "None" or i.lower() == "all":
            return len(self.webPage())
        else:
            return 1

    def showMoneyClick(self):
        self.periodConfirm()
        self.elementClick("span[class='ShowMoney showMoney'] i" ,6) 

    def getMoney(self):
        sleep(1)
        self.periodConfirm()
        return float(self.element("span[class='GetMoney getMoney'] em" ,6).text)
    
    def reflashMoney(self):
        self.periodConfirm()
        self.elementClick("icon" ,1)
    
    def periodConfirm(self): #溫馨提示
        try:
            #self.webDriver.find_element_by_xpath("//span[.='确定']").click()  #固定寫法
            if self.webDriver.find_element_by_id("layermcont").text != "·投注成功·":
                self.webDriver.find_element_by_css_selector("div[class='section'] div[class='layermchild layerConfirm layermanim'] div[class='layermbtn'] span").click() #0808更新
        except:
            pass

    def morePlayClick(self ,i): #更多玩法
        self.periodConfirm()
        actions = webdriver.ActionChains(self.webDriver)
        val1 = self.element("moreMethod" ,1) #固定寫法
        actions.move_to_element(val1).perform()
        clickOk = "NG"
        while clickOk == "NG":
            self.periodConfirm()
            try:
                self.elementsClickOne("ul[class='morePlay'] li" ,6 ,i)
                #self.webDriver.find_element_by_class_name("morePlay").find_elements_by_tag_name("li")[i].click() #固定寫法
                clickOk = "OK"
            except:
                clickOk = "NG"

    def radioWord(self): #秒秒彩判斷
        try:
            self.webDriver.find_element_by_css_selector("span[class='radioWord']") #固定寫法
            return "YES"
        except:
            return "NO"
    
    def timeTitle(self): #期號
        return self.element("div[class='timeTitle'] b" ,6).text 

    def submitCheckOK(self):
        try:
            #if self.element("layermcont" ,1).text == "·投注成功·":
            return "OK"
        except:
            return "OK"
    
    def webPlay(self):
        return self.elements("ul[class='betFilter'] li" ,6) #取得玩法

    def webPlayClick(self ,i): #玩法分頁點擊
        clickOk = "NG"
        while clickOk == "NG":
            self.periodConfirm()
            try:
                self.webPlay()[i].click()
                clickOk = "OK"
            except:
                clickOk = "NG"              

    def webPage(self): #該彩種所有遊戲
        return self.elements("ul[class='betNav fix'] li" ,6) #取得分頁

    def webPageClick(self ,i ,elementText = "" ,link_type = None): #分頁點擊
        clickOK = "NG"
        while clickOK == "NG":
            self.periodConfirm()
            try:
                self.webPage()[i].click()
                clickOK = "OK"
            except:
                clickOK = "NG"
        if i >= 5 and i < len(self.webPage())-1 and len(self.webPage()) > 6:
            clickOk = "NG"
            while clickOk == "NG":
                try:
                    self.elementClick(elementText ,link_type)
                    clickOk = "OK"
                except:
                    clickOk = "NG"

    def webPlayBranchTitle(self): #所有玩法分支
        if self.elements("ul[class='betFilterAnd'] span" ,6) != []:
            return self.elements("ul[class='betFilterAnd'] span" ,6)

    def webPlayBranch(self): #所有玩法分支
        if self.elements("ul[class='betFilterAnd'] a" ,6) != []:
            return self.elements("ul[class='betFilterAnd'] a" ,6)
        else:
            return self.elements("ul[class='betFilterAnd modeZM'] a" ,6)

    def webPlayBranchClick(self ,i ,elementText = "" ,link_type = None): #所有玩法分支點擊
        clickOk = "NG"
        while clickOk == "NG":
            self.periodConfirm()
            try:
                self.webPlayBranch()[i].click()
                clickOk = "OK"
            except:
                clickOk = "NG"

    def webPlayBranchLHC(self): #所有玩法分支(LHC正码)
        if self.elements("ul[class='betFilterAnd modeZM'] a" ,6) != []:
            return self.elements("ul[class='betFilterAnd modeZM'] a" ,6)
        else:
            return self.elements("ul[class='betFilterAnd'] a" ,6)

    def webBall(self ,i = 0):#彩球 
        #return self.elements("div[class='buyNumber fix'] a" ,6)
        return self.webDriver.find_elements_by_css_selector("div[class='buyNumber fix']")[i].find_elements_by_tag_name("a")

    def webBallClick(self ,i ,j = 0):
        clickOk = "NG"
        while clickOk == "NG":
            self.periodConfirm()
            try:
                self.webBall(j)[i].click()
                clickOk = "OK"
            except:
                clickOk = "NG"
                #print("webBall" + "_" + str(j) + "_" + str(i) + "球點擊失敗。")
                
    def webBallDsds(self ,i = 0):#大小單雙
        #return self.elements("div[class='buyNumber fix dsds'] a" ,6)
        return self.webDriver.find_elements_by_css_selector("div[class='buyNumber fix dsds']")[i].find_elements_by_tag_name("a")  #固定寫法

    def webBallDsdsClick(self ,i ,j = 0):
        clickOk = "NG"
        while clickOk == "NG":
            self.periodConfirm()
            try:
                self.webBallDsds(j)[i].click()
                clickOk = "OK"
            except:
                clickOk = "NG"
                #print("webBallDsds" + "_" + str(j) + "_" + str(i) + "球點擊失敗。")
                
    def savePng(self ,save_Text = None ,drop_Down_count = "" ,donot_Save = ""): #存圖
        if save_Text == None or str(donot_Save) != "":
            return
        global funtionError ,funtionCountPng  #全域變數被當成區域變數的解法
        web_Height = self.webDriver.execute_script("return document.body.scrollHeight")
        webPosition_y ,i ,drop_Down = 0 ,1 ,1
        while i <= drop_Down:
            try:
                self.webDriver.execute_script("window.scroll(0, "+ str(webPosition_y) +");")
                sleep(1)
                self.periodConfirm()
                self.webDriver.save_screenshot(testdayFile + "/" + str(testdayTime) + "_" + str(funtionCountPng) + "_" + str(save_Text) + ".png")
                funtionCountPng += 1
                webPosition_y += 600
                i += 1
                if drop_Down_count != "":
                    drop_Down = int(drop_Down_count)
                else:
                    drop_Down = (self.webDriver.execute_script("return document.body.scrollHeight") / 600) + 1
            except:
                funtionError.append(save_Text + str(funtionCountPng) + "_NG")
                return funtionError

    def elementClick(self ,elementText = "" ,link_type = None ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            delayTime = int(str(delayTime).strip())
            if link_type == 1:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.ID,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.ID,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_id(elementText).click()
            elif link_type == 2:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.CLASS_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_class_name(elementText).click()
            elif link_type == 3:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_link_text(elementText).click()
            elif link_type == 4:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_partial_link_text(elementText).click()
            elif link_type == 5:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.NAME,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_name(elementText).click()
            elif link_type == 6:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_css_selector(elementText).click()
            elif link_type == 7:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.TAG_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_tag_name(elementText).click()
            elif link_type == 8:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.XPATH,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.XPATH,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_xpath(elementText).click()
            else:
                return funtionError.append(self.webUrl() + ":" + elementText + "或" + str(link_type) + "點擊錯誤,")
        except:
            funtionError.append(self.webUrl() + ":" + elementText + "_" + str(link_type) + "點擊失敗,")
            #return funtionError
            return "NG"

    def elementsClickOne(self ,elementText = "",link_type = None ,elements_num = 0 ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        self.periodConfirm()
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            elements_num = int(str(elements_num).strip()) 
            if link_type == 1:
                self.webDriver.find_elements_by_id(elementText)[elements_num].click()
            elif link_type == 2:
                self.webDriver.find_elements_by_class_name(elementText)[elements_num].click()
            elif link_type == 3:
                self.webDriver.find_elements_by_link_text(elementText)[elements_num].click()
            elif link_type == 4:
                self.webDriver.find_elements_by_partial_link_text(elementText)[elements_num].click()
            elif link_type == 5:
                self.webDriver.find_elements_by_name(elementText)[elements_num].click()
            elif link_type == 6:
                self.webDriver.find_elements_by_css_selector(elementText)[elements_num].click()
            elif link_type == 7:
                self.webDriver.find_elements_by_tag_name(elementText)[elements_num].click()
            elif link_type == 8:
                self.webDriver.find_elements_by_xpath(elementText)[elements_num].click()
            else:
                return funtionError.append(self.webUrl() + ":" + elementText + "或" + str(link_type) + "(點擊其中一個)錯誤,")
        except:
            funtionError.append(self.webUrl() + ":" + elementText + "_" + str(link_type) + "(點擊其中一個)失敗,")
            return funtionError

    def elementsClickAll(self ,elementText = "",link_type = None ,elements_num = 0 ,delayTime = 0 ,start_num = 0):
        global funtionError #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            elements_num = int(str(elements_num).strip())
            sleep(int(delayTime))
            start_num = int(str(start_num).strip())
            if link_type == 1:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_id(elementText)[i].click()
            elif link_type == 2:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_class_name(elementText)[i].click()
            elif link_type == 3:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_link_text(elementText)[i].click()
            elif link_type == 4:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_partial_link_text(elementText)[i].click()
            elif link_type == 5:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_name(elementText)[i].click()
            elif link_type == 6:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_css_selector(elementText)[i].click()
            elif link_type == 7:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_tag_name(elementText)[i].click()
            elif link_type == 8:
                for i in range(start_num ,elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_xpath(elementText)[i].click()
            else:
                return funtionError.append(self.webUrl() + ":" + elementText + "或" + str(link_type) + "(點擊全部)錯誤,")
        except:
            funtionError.append(self.webUrl() + ":" + elementText + "_" + str(link_type) + "(點擊全部)失敗,")
            return funtionError

    def elements(self ,elementText = "" ,link_type = None ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        self.periodConfirm()
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            sleep(int(delayTime))
            if link_type == 1:
                return self.webDriver.find_elements_by_id(elementText)
            elif link_type == 2:
                return self.webDriver.find_elements_by_class_name(elementText)
            elif link_type == 3:
                return self.webDriver.find_elements_by_link_text(elementText)
            elif link_type == 4:
                return self.webDriver.find_elements_by_partial_link_text(elementText)
            elif link_type == 5:
                return self.webDriver.find_elements_by_name(elementText)
            elif link_type == 6:
                return self.webDriver.find_elements_by_css_selector(elementText)
            elif link_type == 7:
                return self.webDriver.find_elements_by_tag_name(elementText)
            elif link_type == 8:
                return self.webDriver.find_elements_by_xpath(elementText)
            else:
                funtionError.append(self.webUrl() + ":" + elementText + "或" + str(link_type) + "取得(所有)元素錯誤,")
                return funtionError
        except:
            funtionError.append(self.webUrl() + ":" +elementText + "_" + str(link_type) + "取得(所有)元素失敗,")
            return funtionError

    def element(self ,elementText = "",link_type = None):
        global funtionError #全域變數被當成區域變數的解法
        self.periodConfirm()
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            if link_type == 1:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.ID,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.ID,elementText)))
                return self.webDriver.find_element_by_id(elementText)
            elif link_type == 2:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.CLASS_NAME,elementText)))
                return self.webDriver.find_element_by_class_name(elementText)
            elif link_type == 3:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.LINK_TEXT,elementText)))
                return self.webDriver.find_element_by_link_text(elementText)
            elif link_type == 4:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                return self.webDriver.find_element_by_partial_link_text(elementText)
            elif link_type == 5:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.NAME,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.NAME,elementText)))
                return self.webDriver.find_element_by_name(elementText)
            elif link_type == 6:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR,elementText)))
                return self.webDriver.find_element_by_css_selector(elementText)
            elif link_type == 7:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.TAG_NAME,elementText)))
                return self.webDriver.find_element_by_tag_name(elementText)
            elif link_type == 8:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.XPATH,elementText)))
                #WebDriverWait(self.webDriver, 10).until(EC.presence_of_element_located((By.XPATH,elementText)))
                return self.webDriver.find_element_by_xpath(elementText)
            else:
                funtionError.append(self.webUrl() + ":" + elementText + "或" + str(link_type) + "取得元素錯誤,")
                return funtionError
        except:
            funtionError.append(self.webUrl() + ":" + elementText + "_" + str(link_type) + "取得元素失敗,")
            return funtionError

    def elementSendKeys(self ,elementText = "" ,link_type = None ,delayTime = 0 ,text = ""):
        global funtionError #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            delayTime = int(str(delayTime).strip())
            text = str(text).strip()
            if link_type == 1:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.ID,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_id(elementText).send_keys(text)
            elif link_type == 2:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_class_name(elementText).send_keys(text)
            elif link_type == 3:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_link_text(elementText).send_keys(text)
            elif link_type == 4:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_partial_link_text(elementText).send_keys(text)
            elif link_type == 5:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_name(elementText).send_keys(text)
            elif link_type == 6:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_css_selector(elementText).send_keys(text)
            elif link_type == 7:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.TAG_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_tag_name(elementText).send_keys(text)
            elif link_type == 8:
                WebDriverWait(self.webDriver, 30).until(EC.visibility_of_element_located((By.XPATH,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_xpath(elementText).send_keys(text)
            else:
                return funtionError.append(self.webUrl() + ":" + elementText + "或" + str(link_type) + "sendKey錯誤,")
        except:
            funtionError.append(self.webUrl() + ":" + elementText + "_" + str(link_type) + "sendKey失敗,")
            return funtionError

    def rebate(self ,elementText = "" ,link_type = None ,element2_Text = "" ,link2_type = None ,rebate_Number = ""): #返點
        link_type = int(str(link_type).strip())
        elementText = str(elementText).strip()
        link2_type = int(str(link2_type).strip())
        element2_Text = str(element2_Text).strip()
        rebate_Number = str(self.element(elementText ,link_type).text)
        self.elementSendKeys(element2_Text ,link2_type ,text = rebate_Number[0:-2])
        return rebate_Number[0:-2]
    
    def periodDetail(self ,delayTime = 0 ,dateSelect = 0 ,allOrOne = 3): #投注明細
        periodDetail = []
        dateSelect = int(str(dateSelect).strip())
        allOrOne = int(str(allOrOne).strip())

        if dateSelect == 0:
            self.elementClick("li[class='_time'] a:nth-child(1)" ,6) #今天
        elif dateSelect == 1:
            self.elementClick("li[class='_time'] a:nth-child(2)" ,6) #昨天
        elif dateSelect == 2:
            self.elementClick("li[class='_time'] a:nth-child(3)" ,6) #七天
        else:
            print("日期選擇錯誤。")
        sleep(1)
        if allOrOne == 3:
            self.elementClick("li[class='_state'] a:nth-child(1)" ,6) #全部
        elif allOrOne == 4:
            self.elementClick("li[class='_state'] a:nth-child(1)" ,6) #已中奖
        elif allOrOne == 5:
            self.elementClick("li[class='_state'] a:nth-child(1)" ,6) #未中奖
        elif allOrOne == 6:
            self.elementClick("li[class='_state'] a:nth-child(1)" ,6) #等待开奖
        else:
            print("開獎與否選擇錯誤。")
        sleep(1)
        while(True):
            waitTd = 0
            while(waitTd == 0):
                try:
                    self.webDriver.find_element_by_css_selector("div[class='noContent'][style='display: none;']")
                    waitTd = 1
                except:
                    pass
            sleep(int(delayTime)) 
            for i in range(len(self.elements("td" ,7))):
                try:
                    periodDetail.append(self.elements("td" ,7)[i].text)
                except:
                    pass
            try:
                #self.elementClick("a[class='laypage_next']" ,6)
                self.webDriver.find_element_by_xpath("//a[.='下一页']").click() #固定寫法
                sleep(1)
            except:
                break
        return periodDetail

    def CTK3_r(self ,elementText = "" ,link_type = None ,max_Td = "0" ,max_Money = "0"): #傳統快3
        money = [" "]
        link_type = str(link_type).strip()
        elementText = str(elementText).strip()
        money_box = self.elements(elementText ,link_type)
        max_Td = str(max_Td).strip()
        max_Money = str(max_Money).strip()
        if int(max_Td) == 0:
            money_box = money_box[0:-1]
        else:
            money_box = money_box[0:int(max_Td)]
        for i in range(1 ,len(money_box)):
            self.periodConfirm()
            money_box[i].clear()
            if int(max_Money) == 0:
                money.append(random.randint(0 ,99))
            else:
                money.append(int(max_Money))
            money_box[i].send_keys(str(money[i]))
        return money

    def K3_r(self ,elementText = "" ,link_type = None ,max_Td = "0" ,max_Money = "0"): #快3,六合彩
        money = [" "]
        link_type = str(link_type).strip()
        elementText = str(elementText).strip()
        money_box = self.elements(elementText ,link_type)
        max_Td = str(max_Td).strip()
        max_Money = str(max_Money).strip()
        if int(max_Td) == 0:
            money_box = money_box
        else:
            money_box = money_box[0:int(max_Td)]
            
        for i in range(len(money_box)):
            self.periodConfirm()
            money_box[i].clear()
            if int(max_Money) == 0:
                money.append(self.elements("order_type" ,2)[i].text)
                money.append(self.elements("order_zhushu" ,2)[i].text)
                money.append(random.randint(0 ,99))
            else:
                money.append(self.elements("order_type" ,2)[i].text)
                money.append(self.elements("order_zhushu" ,2)[i].text)
                money.append(int(max_Money))
            money_box[i].send_keys(str(money[3 + 3*i]))
        return money

    def KL8(self ,elementText = "" ,link_type = None):
        money = [" "]
        link_type = str(link_type).strip()
        elementText = str(elementText).strip()
        money_box = self.elements(elementText ,link_type)
        for i in range(len(money_box)):
            self.periodConfirm()
            money.append(self.elements("order_type" ,2)[i].text)
            money.append(self.webDriver.find_element_by_css_selector("div[class='checkedListCon']").find_elements_by_tag_name("td")[1 + 7*i].text)
            money.append(self.webDriver.find_element_by_css_selector("div[class='checkedListCon']").find_elements_by_tag_name("td")[2 + 7*i].text)
        return money

class sheet_work():
    def __init__(self ,sheet_work):
        self.sheet_work = sheet_work

    def sheet_value(self ,col = "" ,colNumber = "" ,value = ""):
        col = str(col).strip()
        colNumber = str(colNumber).strip()
        value = str(value).strip()
        self.sheet_work[col + str(len(self.sheet_work[colNumber]))].value = value

    def periodDetail(self ,periodDetail):
        for i in range(int(len(periodDetail)/8)):
            self.sheet_work["B"+str(len(self.sheet_work["B"]) + 1)].value = periodDetail[8*i]
            self.sheet_work["C"+str(len(self.sheet_work["B"]))].value = periodDetail[1 + 8*i]
            
            if periodDetail[2 + 8*i][-4:] == "[详情]":
                periodDetail[2 + 8*i] = periodDetail[2 + 8*i][:-4]
                
            try:
                self.sheet_work["D"+str(len(self.sheet_work["B"]))].value = int(periodDetail[2 + 8*i])
            except:
                self.sheet_work["D"+str(len(self.sheet_work["B"]))].value = periodDetail[2 + 8*i]

            try:
                self.sheet_work["E"+str(len(self.sheet_work["B"]))].value = float(periodDetail[3 + 8*i])
            except:
                self.sheet_work["E"+str(len(self.sheet_work["B"]))].value = periodDetail[3 + 8*i]
            
            if periodDetail[4 + 8*i][-4:] == "[详情]":
                periodDetail[4 + 8*i] = periodDetail[4 + 8*i][:-4]
                
            self.sheet_work["F"+str(len(self.sheet_work["B"]))].value = periodDetail[4 + 8*i]
            
            try:
                self.sheet_work["G"+str(len(self.sheet_work["B"]))].value = float(periodDetail[5 + 8*i])
            except:
                self.sheet_work["G"+str(len(self.sheet_work["B"]))].value = periodDetail[5 + 8*i]
                
            self.sheet_work["H"+str(len(self.sheet_work["B"]))].value = periodDetail[6 + 8*i]
            self.sheet_work["I"+str(len(self.sheet_work["B"]))].value = periodDetail[7 + 8*i]
