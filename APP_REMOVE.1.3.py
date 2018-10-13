# -*- coding: utf-8 -*-

from appium import webdriver
from time import sleep
from openpyxl import workbook , load_workbook
import os ,sys ,re ,time ,ytFuntion ,threading

#1.3多台,啟動chrome並貼上url
def deleteApp(mobileNum ,appNumber):
      desired_caps = {}
      desired_caps = mobileNumber[str(mobileNum)]

      try:
            driver = webdriver.Remote(mobileNumber["{}_{}".format(mobileNum ,"Server")], desired_caps)
      except:
            print("{} 連線有問題,或是UAT APP不存在,請檢查。".format(mobileNumber["{}_{}".format(mobileNum ,"Name")]))
            return
      
      sheetApp = wb["APP"] # 獲取一張表
      for i in appNumber:
            desired_caps["appPackage"] = ""
            desired_caps["appActivity"] = ""
            for j in range(1 ,len(sheetApp["B"]) + 1):
                  if str(sheetApp["B" + str(j)].value) == i:
                        desired_caps["appPackage"] = sheetApp["D" + str(j)].value   #appPackage ：它是App進程包名。
                        desired_caps["appActivity"] = sheetApp["G" + str(j)].value  #appActivity ：它是入口Activity名，這裏通常需要以 . 開頭
                        print("{} {} {}".format(i ,sheetApp["C" + str(j)].value ,sheetApp["F" + str(j)].value))
            if desired_caps["appPackage"] != "" and desired_caps["appActivity"] != "":
                  if driver.is_app_installed(desired_caps["appPackage"]):
                        try:
                              driver.remove_app(desired_caps["appPackage"])#app移除
                              print("{} {}移除成功".format(i ,mobileNumber["{}_{}".format(mobileNum ,"Name")]))
                        except:
                              print("{} {}移除失敗".format(i ,mobileNumber["{}_{}".format(mobileNum ,"Name")]))
                  else:
                        print("{} 不存在{}".format(i ,mobileNumber["{}_{}".format(mobileNum ,"Name")]))
            else:
                  print("請先更新設定中的APP欄位")
                  driver.quit()

def openUrl(mobileNum ,appNumber):
      desired_caps = {}
      desired_caps = mobileNumber[str(mobileNum)]
      
      del desired_caps["appPackage"] #必須刪除appPackage資訊
      del desired_caps["appActivity"]#appActivity
      
      desired_caps["browserName"] = "Chrome"
      desired_caps["noReset"] = "true"
      try:
            driver = webdriver.Remote(mobileNumber["{}_{}".format(mobileNum ,"Server")], desired_caps)
      except:
            print("{} 連線有問題,或是Chrome不存在,請檢查。".format(mobileNumber["{}_{}".format(mobileNum ,"Name")]))
            return
      
      sheetApp = wb["APP"] # 獲取一張表
      for i in appNumber:            
            for j in range(1 ,len(sheetApp["B"])+1):
                  if str(sheetApp["B" + str(j)].value) == i:
                        driver.get(sheetApp["H{}".format(j)].value)
                        break
            while True:
                  try:
                        if "点击下载安卓版" in testWeb.webDriver.find_elements_by_css_selector("div[class='downLink'] a[style='']")[1].text:
                              break
                  except:
                        pass
                  
            testWeb.elementsClickOne("div[class='downLink'] a[style='']" ,6 ,1) #安卓下載
            '''sleep(3)
            chromePermission = driver.find_element_by_id("android:id/button1") #給儲存空間權限
            chromePermission.click()
            sleep(3)
            chromeAllow = driver.find_element_by_id("com.android.chrome:id/button_primary") #允許儲存
            chromeAllow.click()
            openApk = driver.find_element_by_id("com.android.chrome:id/snackbar_button") #開啟
            openApk.click()'''
            
            input("等{} {} 下載完畢,請按ENTER鍵。".format(mobileNumber["{}_{}".format(mobileNum ,"Name")] ,i))
      print()

while(True):
      #mobileNumber = input("請輸入手機編號:").strip()
      appNumber = list(input("請輸入APP序列號(用空白隔開):").split())

      mobileNumber = {}
      
      wb = load_workbook("設定.xlsx")
      sheet = wb["手機"] # 獲取一張表

      for i in range(3 ,len(sheet["B"]) + 1):
            j = i - 3
            desired_caps = {}
            desired_caps["platformName"] = sheet["D" + str(i)].value
            desired_caps["deviceName"] = sheet["E" + str(i)].value
            desired_caps["udid"] = sheet["E" + str(i)].value
            desired_caps['newCommandTimeout'] = '900'
            desired_caps["appPackage"] = "bwt.yfbhj"
            desired_caps["appActivity"] = "bwt.yfbhj.MainActivity"
            devicesName = sheet["C" + str(i)].value
            server = 'http://localhost:{}/wd/hub'.format(sheet["F" + str(i)].value)
            mobileNumber[str(j)] = desired_caps
            mobileNumber["{}_{}".format(j ,"Name")] = devicesName
            mobileNumber["{}_{}".format(j ,"Server")] = server
      
      threads = []
      
      '''for i in range(len(mobileNumber)//3):
            threads.append(threading.Thread(target = deleteApp, args = (i ,appNumber ,)))
            threads[i].start()
            sleep(1)
            
      for i in range(len(mobileNumber)//3):
            threads[i].join()

      threadsChrome = []
      
      for i in range(len(mobileNumber)//3):
            threadsChrome.append(threading.Thread(target = openUrl, args = (i ,appNumber ,)))
            threadsChrome[i].start()
            sleep(1)

      for i in range(len(mobileNumber)//3):
            threadsChrome[i].join()'''

      #print("456")      
      #break
      '''for i in range(1 ,len(sheet["B"])+1):
            if str(sheet["B" + str(i)].value) == mobileNumber:
                  desired_caps["platformName"] = sheet["D" + str(i)].value    #platformName ：它是平台名稱，需要區分Android或iOS，此處填寫Android。
                  desired_caps["deviceName"] = sheet["E" + str(i)].value      #deviceName ：它是設備名稱，此處是手機的具體類型。JFMNT48DYSIB6SYL紅米
                  desired_caps["udid"] = sheet["E" + str(i)].value      #多台同時使用
                  devicesName = sheet["C" + str(i)].value
                  server = 'http://localhost:{}/wd/hub'.format(sheet["F" + str(i)].value) 
      desired_caps['newCommandTimeout'] = '900'

      desired_caps["appPackage"] = "bwt.yfbhj"
      desired_caps["appActivity"] = "bwt.yfbhj.MainActivity"
      try:
            driver = webdriver.Remote(server, desired_caps)
      except:
            input("{} 連線有問題,或是UAT APP不存在,請檢查。".format(devicesName))

      sheetApp = wb["APP"] # 獲取一張表
      for i in appNumber:
            desired_caps["appPackage"] = ""
            desired_caps["appActivity"] = ""
            for j in range(1 ,len(sheetApp["B"])+1):
                  if str(sheetApp["B" + str(j)].value) == i:
                        desired_caps["appPackage"] = sheetApp["D" + str(j)].value   #appPackage ：它是App進程包名。
                        desired_caps["appActivity"] = sheetApp["G" + str(j)].value  #appActivity ：它是入口Activity名，這裏通常需要以 . 開頭
                        print("{} {} {}".format(i ,sheetApp["C" + str(j)].value ,sheetApp["F" + str(j)].value))
            if "appPackage" in desired_caps and "appActivity" in desired_caps:
                  if driver.is_app_installed(desired_caps["appPackage"]):
                        try:
                              #driver = webdriver.Remote(server, desired_caps)
                              driver.remove_app(desired_caps["appPackage"])#app移除
                              print("{} 移除成功".format(i))
                        except:
                              print("{} 移除失敗".format(i))
                  else:
                        print("{} 不存在此手機".format(i))
            else:
                  input("請先更新設定中的APP欄位")
                  driver.quit()
                  
      del desired_caps["appPackage"] #必須刪除appPackage資訊
      del desired_caps["appActivity"]#appActivity
      desired_caps["browserName"] = "Chrome"
      driver = webdriver.Remote(server, desired_caps)
      for i in appNumber:            
            for j in range(1 ,len(sheetApp["B"])+1):
                  if str(sheetApp["B" + str(j)].value) == i:
                        driver.get(sheetApp["H{}".format(j)].value)
                        break
            input("等 {} 下載安裝完畢,請按ENTER鍵。".format(i))'''
      print()
