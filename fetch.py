#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import base64
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile

import config


def dmesg(m):
    global dbg
    if dbg:
        print '[dmesg] %s'%m



script_dir=os.path.split(os.path.realpath(__file__))[0]+'/'

profile_dir=config.browser_profile_dir
dbg=config.debug
command_executor=config.browser_driver_command_executor


def init():
    global driver
    dmesg('loading browser profile')
    profile=FirefoxProfile(profile_directory=profile_dir)
    #profile=base64.b64encode(profile)
    dmesg('starting browser')
    driver = webdriver.Remote(command_executor=command_executor,desired_capabilities=DesiredCapabilities.FIREFOX,browser_profile=profile)
    dmesg('打开 wos 页面，apps子域首页')
    driver.get('http://apps.webofknowledge.com/')
    if 'login.webofknowledge.com/error/Error' in driver.current_url:
        dmesg('login status Error, trying re-login')
        ele=driver.find_element_by_name('username')
        ele.clear()
        ele.send_keys(config.wos_username)
        ele=driver.find_element_by_name('password')
        ele.clear()
        ele.send_keys(config.wos_password)
        ele=driver.find_element_by_name('rememberme')
        if ele.get_property('checked')==False:
            ele.click()
        ele=driver.find_element_by_tag_name('button')
        ele.click()
        sleep(1)
    if 'login.webofknowledge.com/error/Error' in driver.current_url:
        exit('login failed, Maybe VPN required.')
    elif 'apps.webofknowledge.com/' in driver.current_url:
        dmesg('login passed')
    else:
        exit('登录状态未知，为可靠起见，请检查')
        pass

# 进入指定search_text的高级检索结果页。首选从检索历史中相应历史结果直接进入
def adv_search_and_go(search_text):
    global driver
    wos_his_label=check_searched(search_text)
    if wos_his_label=='':
        # wos搜索历史中没有，去高级搜索
        adv_search(search_text)
    dmesg('再次检查检索历史查询')
    wos_his_label=check_searched(search_text)
    goto_wos_search_history()
    assert '.com/WOS_CombineSearches_input.do?' in driver.current_url

    ele=driver.find_elements_by_xpath("//form[@name='WOS_CombineSearches_input_form']//tbody/tr[@id]")
    idx=-1
    for i in range(len(ele)):
        e=ele[i].find_elements_by_class_name("historySetNum")
        if len(e)==1 and e[0].text==wos_his_label:
            idx=i
            dmesg('搜索结果中找到该条目，排第 %s 条' %i)
    if idx == -1:
        exit('ERROR *** wos搜索结果中没找到该项 %s' %wos_his_label)
    e=ele[idx].find_element_by_class_name("historyResults")
    e=e.find_element_by_tag_name("a")
    dmesg('点击结果条目，进入结果页')
    e.click()



def goto_wos_search_history():
    global driver
    ele=driver.find_elements_by_xpath("//div[@id='skip-to-navigation']//a[contains(@href,'_CombineSearches_input.do')]")
    if len(ele)==0:
        exit('无法在主导航栏中找到“检索历史”: href contains(_CombineSearches_input.do)')
    elif len(ele)>1:
        exit('在主导航栏中找到多个“检索历史”项: href contains(_CombineSearches_input.do)')
    else:
        pass
    ele[0].click()
    sleep(1)
    if '.com/WOS_CombineSearches_input.do?' not in driver.current_url:
        url=driver.current_url.encode('utf-8')
        #u[u.find('/',10):u.find('?')]
        dmesg('当前非检索历史页面，尝试切换数据库')
        dmesg(url)
        dmesg('展开数据库下拉列表')
        ele=driver.find_element_by_xpath("//div[@class='dbselectdiv']//span[@class='select2-selection__arrow']")
        ele.click()
        sleep(0.3)
        dmesg('点击wos核心合集切换')
        ele=driver.find_element_by_xpath("//ul[@id='select2-databases-results'][@class='select2-results__options']/li[contains(text(),'Web of Science')]")
        ele.click()
        sleep(0.3)
        dmesg('进入wos核心合集检索历史页')


#检查搜索历史中是否有search_text，返回wos检索式序列，形式为 # 2
def check_searched(search_text):
    global driver
    goto_wos_search_history()
    assert '.com/WOS_CombineSearches_input.do?' in driver.current_url
    # 遍历历史条目，寻找与search_text一致项,索引号存储于 idx
    idx = -1
    wos_his_label=''
    ele=driver.find_elements_by_xpath("//form[@name='WOS_CombineSearches_input_form']//tbody/tr[@id]")
    for i in range(len(ele)):
        # 下一行如果用 .find_elements_by_xpath("//div[@class='historyQuery']") 似乎会把其所有 ele的结果都选择出来，存疑中
        e=ele[i].find_element_by_class_name('historyQuery')
        if e.text==search_text:
            wos_his_label=ele[i].find_element_by_class_name('historySetNum').text.encode('utf-8')
            wos_his_count=ele[i].find_element_by_class_name('historyResults').text.encode('utf-8')
            idx=i
            dmesg('检索历史中找到本项检索式 [%s]，结果  %s 条' %(wos_his_label,wos_his_count))
            break
    if idx == -1:
        dmesg('检索历史中没有本项检索式，需要转去高级搜索一次再回来')
        # TODO
    else:
        #点击结果链接进入结果页
        #ele[idx].find_element_by_class_name('historyResults')
        pass
    return wos_his_label





def adv_search(search_text):
    global driver
    dmesg('返回首页')
    ele=driver.find_element_by_xpath("//div[@class='logoBar']//a[contains(@href,'/home.do?')]")
    ele.click()
    sleep(1)
    # 点选“Web of Scienct核心合集”进入高级检索页面
    dmesg('展开数据库下拉列表')
    ele=driver.find_element_by_xpath("//div[@class='dbselectdiv']//span[@class='select2-selection__arrow']")
    ele.click()
    sleep(0.3)
    dmesg('点击wos核心合集切换')
    ele=driver.find_element_by_xpath("//ul[@id='select2-databases-results'][@class='select2-results__options']/li[contains(text(),'Web of Science')]")
    ele.click()
    sleep(0.3)
    dmesg('点tab上高级检索')
    ele=driver.find_element_by_xpath("//ul[@class='searchtype-nav']//a[contains(@href,'WOS_AdvancedSearch_input.do?')]")
    ele.click()
    sleep(1)
    dmesg('高级检索页面输入条件:\n%s'%search_text)
    ele=driver.find_element_by_xpath("//form[@id='WOS_AdvancedSearch_input_form']//div[@class='AdvSearchBox']//textarea[@id='value(input1)']")
    ele.clear()
    ele.send_keys(search_text)
    ele=driver.find_element_by_xpath("//form[@id='WOS_AdvancedSearch_input_form']//div[@class='AdvSearchBox']//span[@id='searchButton']//button[@id='search-button']")
    ele.click()
    dmesg('高级搜索完成，搜索历史中应该已经有该项')



# 普通搜索，目前已经没用
def search(search_text):
    #点输入框尾的 x 号图标
    ele=driver.find_element_by_id('clearIcon1')
    ele.click()
    sleep(0.3)
    ele=driver.find_element_by_id('value(input1)')
    ele.click()
    ele.send_keys('2017')
    sleep(0.3)


    #展开检索框右的下拉选单，并寻找其中“出版年”一项点选 u'\u51fa\u7248\u5e74'
    #似乎登录与不登录的表单是不同的
    ele=driver.find_element_by_id('searchrow1')
    e=ele.find_element_by_class_name('select2-selection__arrow')
    e.click()
    sleep(1)
    ele=driver.find_element_by_id('select2-select1-results')
    e=ele.find_elements_by_tag_name('li')
    et=[t.text for t in e]
    if u'\u51fa\u7248\u5e74' not in et:
        exit('找不到“出版年"下拉项')
    k=et.index(u'\u51fa\u7248\u5e74')
    e[k].click()
    sleep(0.3)

    #根据登录状态不同，会有两个不同的检索页，分支处理两种检索按钮的点击
    if 'WOS_ClearGeneralSearch.do?' in driver.current_url:
        ele=driver.find_element_by_id('searchCell3')
        e=ele.find_element_by_tag_name('button')
    elif 'WOS_GeneralSearch_input.do?' in driver.current_url:
        #ele=driver.find_element_by_class_name('searchButton')
        ele=driver.find_element_by_id('searchCell1')
        e=ele.find_element_by_tag_name('button')
    else:
        exit('查找搜索按钮时遇到未知页面 %s'%driver.current_url.encode('utf-8'))
    #e.send_keys(Keys.RETURN)
    e.click()
    sleep(1)



# 按 start, bat_size 执行一次保存；需要保证当前为检索结果页面
def cral(start,batch_size):
    global driver
    #点击按钮，打开发送文件html层；防止页面不完整而多次重试
    times=0
    while True:
        times+=1
        ele=driver.find_elements_by_xpath("//div[@class='page-options-inner']//span[@class='select2-selection__arrow']")
        if len(ele)==1:
            break
        if times > config.cral_page_reload_retry:
            exit('summary.do页面已超过重试次数限制，请手工检查是否可以正常访问')
        e=driver.find_elements_by_xpath("//form[@id='summary_navigation']//a[not(contains(@class,'Disabled'))]")
        if len(e)>=1:
            dmesg('summary.do页面加载不完整，第 %s 次，尝试翻页'%times)
            e[0].click()
            sleep(1)
        else:
            dmesg('summary.do页面加载不完整，第 %s 次，尝试重新加载'%times)
            driver.refresh()
            sleep(1)
    ele[0].click()
    sleep(0.3)

    #点选 “保存为其他文件格式” u'\u4fdd\u5b58\u4e3a\u5176\u4ed6\u6587\u4ef6\u683c\u5f0f'
    ele=driver.find_elements_by_xpath("//span[contains(@class,'select2-container')]//ul[@id='select2-saveToMenu-results']//li")
    et=[it.text for it in ele]
    if u'\u4fdd\u5b58\u4e3a\u5176\u4ed6\u6587\u4ef6\u683c\u5f0f' not in et:
        exit('找不到“保存为其他文件格式”')
    k=et.index(u'\u4fdd\u5b58\u4e3a\u5176\u4ed6\u6587\u4ef6\u683c\u5f0f')
    ele[k].click()


    # 点选记录范围并输入
    ele=driver.find_element_by_id("numberOfRecordsRange")
    ele.click()
    sleep(0.3)

    ele=driver.find_element_by_id("markFrom")
    ele.clear()
    ele.send_keys('%s'%start)
    sleep(0.3)

    ele=driver.find_element_by_id("markTo")
    ele.clear()
    ele.send_keys('%s'%(start+batch_size-1))
    sleep(0.3)

    # 更改记录内容下拉选项
    ele=driver.find_element_by_class_name("quickoutput-content")
    e=ele.find_element_by_class_name("select2-selection__arrow")
    e.click()
    sleep(0.3)

    ele=driver.find_elements_by_id("select2-bib_fields-results")
    e=ele[0].find_elements_by_tag_name('li')
    e[3].click()
    sleep(0.3)

    # 更改文件格式下拉选项
    ## 找三角图标，点击展开下拉项
    ele=driver.find_elements_by_class_name("quick-output-detail")
    e=ele[0].find_elements_by_class_name('select2-selection__arrow')
    e[0].click()
    sleep(0.3)

    ele=driver.find_elements_by_id("select2-saveOptions-results")
    e=ele[0].find_elements_by_tag_name('li')
    e[6].click()
    sleep(0.3)


    # 点下载按钮、然后点关闭html层
    ele=driver.find_elements_by_class_name("quickoutput-action")
    e=ele[0].find_element_by_tag_name('button')
    e.click()
    sleep(0.3)

    ele=driver.find_elements_by_class_name("quickoutput-cancel-action")
    ele[0].click()
    sleep(0.3)


    #翻页
    ele=driver.find_elements_by_class_name("paginationNext")
    ele[0].click()
    sleep(0.3)



def wait_for_file():
    pass
    file_name='download_file_name.txt'
    return file_name



def write_to_log(task_label,start,batch_size,file_name):
    pass





if __name__ == '__main__':
    init()
    search_text='PY=2017 AND (WC=ONCOLOGY OR WC=CHEMISTRY MULTIDISCIPLINARY OR WC=BIOCHEMISTRY)'
    adv_search_and_go(search_text)

    cral(101,50)
    sleep(10)
    file_name=wait_for_file()
    #write_to_log(task_label,start,batch_size,file_name)
    cral(151,50)
    sleep(10)
    file_name=wait_for_file()
    cral(201,50)
    sleep(10)




