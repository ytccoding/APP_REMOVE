# -*- coding: utf-8 -*-

from appium import webdriver
from time import sleep
from openpyxl import workbook , load_workbook
import os ,sys ,re ,time ,ytFuntion

#1.1新增先啟動UAT APP,確保有連線。增加確認APP是否存在,改善移除速度
while(True):
      mobileNumber = input("請輸入手機編號:").strip()
      appNumber = list(input("請輸入APP序列號(用空白隔開):").split())

      desired_caps = {}

      wb = load_workbook("設定.xlsx")
      sheet = wb["手機"] # 獲取一張表

      for i in range(1 ,len(sheet["B"])+1):
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
            input("{} 連線有問題,請檢查".format(devicesName))

      sheetApp = wb["APP"] # 獲取一張表
      for i in range(len(appNumber)):
            desired_caps["appPackage"] = ""
            desired_caps["appActivity"] = ""
            for j in range(1 ,len(sheetApp["B"])+1):
                  if str(sheetApp["B" + str(j)].value) == appNumber[i]:
                        desired_caps["appPackage"] = sheetApp["D" + str(j)].value   #appPackage ：它是App進程包名。
                        desired_caps["appActivity"] = sheetApp["G" + str(j)].value  #appActivity ：它是入口Activity名，這裏通常需要以 . 開頭
                        print("{} {} {}".format(appNumber[i] ,sheetApp["C" + str(j)].value ,sheetApp["F" + str(j)].value))
            if "appPackage" in desired_caps and "appActivity" in desired_caps:
                  if driver.is_app_installed(desired_caps["appPackage"]):
                        try:
                              driver = webdriver.Remote(server, desired_caps)
                              driver.remove_app(desired_caps["appPackage"])#app移除
                        except:
                              print("{}移除失敗".format(appNumber[i]))
                  else:
                        print("此手機無{}".format(appNumber[i]))
            else:
                  input("請先更新設定中的APP欄位")
                  driver.quit()
      #break
