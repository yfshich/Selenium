#!/user/bin/env python
#-*- coding:utf-8 -*-
from selenium import webdriver
import xlrd
import xlwt
from xlutils.copy import copy
import os
import time

driver = webdriver.Chrome()
driver.fullscreen_window()
driver.get("http://eduweixin.fycms.com/admin.php?s=/Public/login.html ")
driver.find_element_by_name("username").send_keys("admin")
driver.find_element_by_name("password").send_keys("admin")
driver.find_element_by_name("Submit").click()
time.sleep(2)
driver.switch_to.frame("ifrmenu")
driver.find_element_by_link_text("会议管理").click()
time.sleep(2)
driver.switch_to.default_content()
driver.switch_to.frame("ifrcontent")
driver.find_element_by_xpath("/html[1]/body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/a[1]").click()
time.sleep(2)


def load_Table(page):
    #创建工作簿
    wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
    #创建工作表
    sheet = wbk.add_sheet('sheet 1', cell_overwrite_ok=True)
    excel = r"C:\test\test.xls"
    driver.switch_to.default_content()
    driver.switch_to.frame("ifrcontent")
    table_rows = driver.find_element_by_id("HFCMS").find_elements_by_tag_name("tr")
    row = 20
    print(row)
    print("开始保存第%s页" % str(page+1))
    for i, tr in enumerate(table_rows):
        if i==0 and page==0:
            print("开始保存表头")
            table_cols1 = tr.find_elements_by_tag_name('th')
            for j, tc in enumerate(table_cols1):
                sheet.write(i, j, tc.text)
                wbk.save(excel)
            print("表头保存成功")
        else:
            table_cols2 = tr.find_elements_by_tag_name('td')
            for j, tc in enumerate(table_cols2):
                 #老的工作簿，打开excel
                 oldWb = xlrd.open_workbook(excel, formatting_info=True)
                 #新的工作簿,复制老的工作簿
                 newWb = copy(oldWb)
                 #新的工作表
                 newWs = newWb.get_sheet(0)
                 newWs.write(i + page * row, j, tc.text)
                 os.remove(excel)
                 newWb.save(excel)
    print("保存成功")

def switch_page():
    pages = driver.find_element_by_class_name("pagelist").find_elements_by_tag_name("a")
    t = len(pages)
    print(t)
    for i in range(t):
        if i>=1:
            driver.find_element_by_class_name("pagelist").find_element_by_link_text(str(i+1)).click()
            load_Table(i)
        else:
            load_Table(i)
switch_page()
print("数据获取完毕")