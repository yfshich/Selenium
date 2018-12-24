#!/user/bin/env python
#-*- coding:utf-8 -*-
from selenium import webdriver
import xlwt
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
#创建工作薄
wbk = xlwt.Workbook(encoding="utf-8", style_compression=0)
#创建工作表
sheet = wbk.add_sheet("Sheet1", cell_overwrite_ok=True)


excel = "C:\\test\\test.xls"

def saveData():
    #table_list=[]
    #row_list=[]
    #表头
    driver.switch_to.default_content()
    driver.switch_to.frame("ifrcontent")
    table_top_list = driver.find_element_by_id("HFCMS").find_element_by_tag_name("thead").find_elements_by_tag_name("th")
    for c,top in enumerate(table_top_list):
        #row_list.apend(top.text)
        sheet.write(0, c, top.text)
        print(top.text)
    #table_list.apend(row_list)

    #表的内容
    #将表的每一行存在table_tr_list中
    table_tr_list = driver.find_element_by_id("HFCMS").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr")
    # 每行输出到row_list中,将所有的row_list输入到table_list中
    for r, tr in enumerate(table_tr_list, 1):
        # 将表的每一行的每一列内容存在table_td_list中
        table_td_list = tr.find_elements_by_tag_name('td')
        # 将行列的内容加入到table_list中
        for c, td in enumerate(table_td_list):
            # row_list.append(td.text)
            sheet.write(r, c, td.text)
            print(td.text)
        # table_list.append(row_list)
    # 最后返回table_list
    # return table_list
saveData()

#保存文件
wbk.save(excel)
#先回到原来桌面
driver.switch_to.default_content()
print('done')
