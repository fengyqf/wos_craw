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

'''
1. 运行环境：
  - a. 远程selenium 服务器，其地址为 browser_driver_command_executor 变量；其profiles 为 browser_profile_dir变量
  - b. selenium 服务器 只支持firefox，因为要使用 browser_profile；如不使用该项，应该也可以支持其他浏览器
        （注意要使用中文版，wos自动适配语言的，而本脚本通过一些中文字符串定位页面元素）
  - c. 客户端（脚本运行机器）需要监测下载进度，所以把客户端的本地目录，映射为服务器上firefox下载目录，
        （即配置文件中的 file_save_dir；windows共享、虚拟机共享等都可以）
  - d. 手工将服务器端的firefox下载地址设置成上述地址、text文件不经询问直接保存到该目录，
  - e. 然后将其 profiles 目录复制到客户端上的browser_profile_dir目录，供firefox启动时加载。
        （selemium服务器端浏览器默认不加载任何配置项，相当于新装firefox，所以必须通过profiles使其生效）
  - f. 经过cygwin下通过测试，可正常运行. 服务端 windows 7, firefox 59, jdk8.161
        （windows xp下seleium启动报错；客户端脚本或许能跑，鉴于当前cygwin已不再支持xp，所以不建议使用）
  - g. 脚本运行过程中，通常不要点击服务端的firefox，以避免DOM变更而导致脚本出错

2. 一些特性
  - a. config.py 中设置 sch 变量，逐一设置每个wos检索项；
        每项要设置一个惟一的label，及wos高级检索条件search_text；
        每个检索条件，其结果应该控制 在10万条以内（wos限制），超过则直接忽略下载
        可以通过配置项中增加 limit 参数，只导出前N条，见config.py.sample 中示例
  - b. 默认每批下载 file_save_batch_size 条
  - c. 下载后，txt文件将被自动改名，下载文件自动改名为 {label}_{起始条数}.txt
  - d. 下载时会检查是否有对应的.txt文件，如果已有，则自动跳过不再下载。因此label要惟一，不能重复
  - e. 一次执行可能会有漏掉下载的部分（遇异常而跳过），只需要简单的再次执行即可，直到全部下载完成为止
        可以将全部完成的sch项注释掉（脚本每次执行都要检查所有sch项的抓取进度）
  - f. 如果怀疑某些.txt文件不完整，将其删除，再跑一遍脚本即可。
  - g. 脚本跑完到终点时，会报告本次总共抓取批数，如果是0，而且前面也没有报告异常，应该是抓完了。
'''

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
    dmesg('starting browser')
    driver = webdriver.Remote(command_executor=command_executor,desired_capabilities=DesiredCapabilities.FIREFOX,browser_profile=profile)
    dmesg('打开 wos 页面，apps子域首页')
    try:
        driver.get('http://apps.webofknowledge.com/')
    except:
        dmesg('打开失败，10s 后重试...')
        sleep(10)
        try:
            driver.get('http://apps.webofknowledge.com/')
        except:
            dmesg('打开失败，10s 后重试...')
            sleep(10)
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
#  返回值为wos检索结果条数，数值型
def adv_search_and_go(search_text):
    global driver
    rs_count=0
    wos_his_label=check_searched(search_text)
    if wos_his_label=='':
        # wos搜索历史中没有，去高级搜索
        adv_search(search_text)
    dmesg('再次检查检索历史查询')
    wos_his_label=check_searched(search_text)
    if goto_wos_search_history() ==False:
        return False
    assert '.com/WOS_CombineSearches_input.do?' in driver.current_url

    ele=driver.find_elements_by_xpath("//form[@name='WOS_CombineSearches_input_form']//tbody/tr[@id]")
    idx=-1
    for i in range(len(ele)):
        e=ele[i].find_elements_by_class_name("historySetNum")
        if len(e)==1 and e[0].text==wos_his_label:
            idx=i
            dmesg('搜索结果中找到该条目，排第 %s 条' %(i+1))
    if idx == -1:
        dmesg('ERROR *** wos搜索结果中没找到该项 %s' %wos_his_label)
        return idx

    e=ele[idx].find_element_by_class_name("historyResults")
    rs_text=e.text
    rs_count=int(rs_text.replace(',',''))
    e=e.find_element_by_tag_name("a")
    dmesg('wos检索历史显示有 %s 条结果'%rs_count)
    dmesg('点击结果条目，进入结果页')
    e.click()
    return rs_count



def goto_wos_search_history():
    global driver
    ele=driver.find_elements_by_xpath("//div[@id='skip-to-navigation']//a[contains(@href,'_CombineSearches_input.do')]")
    if len(ele)==0:
        dmesg('无法在主导航栏中找到“检索历史”: href contains(_CombineSearches_input.do)')
        return False
    elif len(ele)>1:
        dmesg('在主导航栏中找到多个“检索历史”项: href contains(_CombineSearches_input.do)')
        return False
    else:
        pass
    ele[0].click()
    sleep(1)
    if '.com/WOS_CombineSearches_input.do?' not in driver.current_url:
        url=driver.current_url.encode('utf-8')
        dmesg('当前非检索历史页面，尝试切换数据库')
        print(url)
        dmesg('展开数据库下拉列表')
        ele=driver.find_element_by_xpath("//div[@class='dbselectdiv']//span[@class='select2-selection__arrow']")
        ele.click()
        sleep(0.3)
        dmesg('点击wos核心合集切换')
        ele=driver.find_element_by_xpath("//ul[@id='select2-databases-results'][@class='select2-results__options']/li[contains(text(),'Web of Science')]")
        ele.click()
        sleep(0.3)
        dmesg('进入wos核心合集检索历史页')
    return True


#检查搜索历史中是否有search_text，返回wos检索式序列，形式为 # 2
def check_searched(search_text):
    global driver
    if goto_wos_search_history()==False:
        return False
    sleep(1)
    if not '.com/WOS_CombineSearches_input.do?' in driver.current_url:
        dmesg('当前不是搜索结果页 WOS_CombineSearches_input.do ， retry again')
        if goto_wos_search_history()==False:
            dmesg('retry failed')
            return False
        sleep(1)

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
    else:
        pass
    return wos_his_label





def adv_search(search_text):
    global driver
    sleep(1)
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

    if '/WOS_AdvancedSearch_input.do?product=WOS&' in driver.current_url:
        dmesg('当前即是wos核心合集库的高级检索')
    else:
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



# 按 start, end 执行一次保存；需要保证当前为检索结果页面
def crawl(start,end,label):
    global driver
    #点击按钮，打开发送文件html层；防止页面不完整而多次重试
    times=0
    while True:
        times+=1
        ele=driver.find_elements_by_xpath("//div[@class='page-options-inner']//span[@class='select2-selection__arrow']")
        if len(ele)==1:
            break
        if times > config.crawl_page_reload_retry:
            dmesg('summary.do页面已超过重试次数限制，请手工检查是否可以正常访问')
            return False
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
    sleep(1)

    #点选 “保存为其他文件格式” u'\u4fdd\u5b58\u4e3a\u5176\u4ed6\u6587\u4ef6\u683c\u5f0f'
    dmesg('点选 “保存为其他文件格式”')
    ele=driver.find_elements_by_xpath("//span[contains(@class,'select2-container')]//ul[@id='select2-saveToMenu-results']//li")
    et=[it.text for it in ele]
    if u'\u4fdd\u5b58\u4e3a\u5176\u4ed6\u6587\u4ef6\u683c\u5f0f' not in et:
        dmesg('找不到“保存为其他文件格式”')
        return False
    k=et.index(u'\u4fdd\u5b58\u4e3a\u5176\u4ed6\u6587\u4ef6\u683c\u5f0f')
    ele[k].click()
    sleep(1)

    dmesg('点选记录范围并输入')
    ele=driver.find_element_by_id("numberOfRecordsRange")
    ele.click()
    sleep(1)

    ele=driver.find_element_by_id("markFrom")
    ele.clear()
    ele.send_keys('%s'%start)
    sleep(1)

    ele=driver.find_element_by_id("markTo")
    ele.clear()
    ele.send_keys('%s'%end)
    sleep(1)

    dmesg('更改记录内容下拉选项')
    ele=driver.find_element_by_class_name("quickoutput-content")
    e=ele.find_element_by_class_name("select2-selection__arrow")
    e.click()
    sleep(1)

    ele=driver.find_elements_by_id("select2-bib_fields-results")
    e=ele[0].find_elements_by_tag_name('li')
    e[3].click()
    sleep(1)

    dmesg('更改文件格式下拉选项')
    # 更改文件格式下拉选项
    ## 找三角图标，点击展开下拉项
    ele=driver.find_elements_by_class_name("quick-output-detail")
    e=ele[0].find_elements_by_class_name('select2-selection__arrow')
    e[0].click()
    sleep(1)

    ele=driver.find_elements_by_id("select2-saveOptions-results")
    e=ele[0].find_elements_by_tag_name('li')
    e[6].click()
    sleep(1)


    snap=os.listdir(config.file_save_dir)
    dmesg('扫描下载目录，缓存文件列表，计 %s 个'%len(snap))

    # 点下载按钮、然后点关闭html层
    ele=driver.find_elements_by_class_name("quickoutput-action")
    e=ele[0].find_element_by_tag_name('button')
    e.click()
    dmesg('点下载按钮、等待文件下载完成')
    #sleep(0.3)

    filename='%s_%s.txt'%(label,start)
    rtn=wait_new_file(config.file_save_dir,snap,filename)
    dmesg('等候下载wait_new_file()完成状态： %s '%rtn)

    dmesg('点“取消”按钮关闭弹出层，并翻页')
    ele=driver.find_elements_by_class_name("quickoutput-cancel-action")
    ele[0].click()
    sleep(1)
    ele=driver.find_elements_by_class_name("paginationNext")
    ele[0].click()
    sleep(1)

    #返回下载成功与否的状态，以wait_new_file()返回值为准
    return rtn



# 扫描目录中，发现新增文件，下载完成后，改名为storage_name
def wait_new_file(dir,snap,storage_name):
    times=0
    download_max_wait=10
    f1_last_size=0
    f2_last_size=0
    download_zombie=0
    while True:
        files=os.listdir(dir)
        new=[file for file in files if file not in snap]
        #firefox 正在下载的文件，文件名加 .part，据此识别新文件正在下载中
        if len(new)==1 and new[0][:9]=='savedrecs' and new[0][-4:]=='.txt':
            dmesg('发现预期新文件并改名归档：%s -> %s'%(new[0],storage_name))
            os.rename('%s/%s'%(config.file_save_dir,new[0]) , '%s/%s'%(config.file_save_dir,storage_name) )
            return True
        elif len(new)==2 and (new[0][-5:]=='.part' or new[1][-5:]=='.part') and (new[0] in new[1] or new[1] in new[0]):
            new.sort()
            f1=new[0]
            f2=new[1]
            f1_path='%s/%s'%(config.file_save_dir,new[0])
            f2_path='%s/%s'%(config.file_save_dir,new[1])
            try:
                f1_size=os.path.getsize(f1_path)
                f2_size=os.path.getsize(f2_path)
            except OSError,e:
                dmesg('临时文件读取失败，可能已下载完成')
                sleep(1)
                continue
            if download_zombie > config.download_zombie_retry:
                dmesg('**WARNING** 下载进度僵死过多次数，失败。 %s 次'%download_zombie)
                return False
            elif f1_size==f1_last_size and f2_size==f2_last_size:
                download_zombie+=1
                dmesg('下载进度僵死中，第 %s 次'%download_zombie)
                sleep(1)
            elif f1_size>f1_last_size or f2_size>f2_last_size:
                download_zombie=0
                if f1_size >0:
                    dmesg('下载中...   %s: %s     %s: %10d'%(f1,f1_size,f2,f2_size))
                else:
                    dmesg('下载中...   %s: %10d'%(f2,f2_size))
                sleep(1)
            else:
                dmesg('发现正在下载的文件：')
                dmesg('        %s:  %s Bytes'%(f1,f1_size))
                dmesg('        %s:  %s Bytes'%(f2,f2_size))
                sleep(1)
            f1_last_size=f1_size
            f2_last_size=f2_size
        elif len(new)==2 or len(new)==1:
            dmesg('**WARNING** 发现新文件，但文件名非预期，未做任何处理 %s'%new)
            return False
        elif times >= config.crawl_page_reload_retry:
            dmesg('第 %s 次扫描未发现新文件，已达最大重试次数'%times)
            return False
        else:
            times+=1
            dmesg('第 %s 次扫描未发现新文件，3s后重试'%times)
            sleep(3)


# 监测文件直到其大小不再增长
#   看着挺好，但没用了，因为firefox下载中的文件带.part后续，完成后再改名
def wait_for_finish_down(file_full_path,warning_size=1024*1024):
    last_size=0
    stable_times=0
    sys.stdout.write('        ')
    while True:
        sleep(1)
        size=os.path.getsize(file_full_path)
        if size > last_size:
            stable_times=0
            last_size=size
            sys.stdout.write('+')
        else:
            stable_times+=1
            if stable_times >= 5:
                sys.stdout.write('  finished')
                break
            else:
                pass
                sys.stdout.write('_')
    print ''
    return size



def write_to_log(task_label,start,batch_size,file_name):
    pass





def run():
    total_crawled_batch_count=0
    sch_count=len(config.sch)
    for i in range(sch_count):
        sch_crawled_batch_count=0
        label=config.sch[i]['label']
        search_text=config.sch[i]['search_text']
        batch_size=config.file_save_batch_size
        print('\n==== 开始处理第 %s/%s 批检索 [%s] ==== '%(i+1,sch_count,label))
        print(search_text)
        rs_count=adv_search_and_go(search_text)
        if rs_count > 100000 and not config.sch[i].has_key('limit'):
            print '\n\n*********************************************'
            print '*** [注意] 检索结果超过100000条，超出部分将无法下载 ***'
            print '*** 本项检索条件已被忽略，请修改规则后再运行脚本 ***'
            print '*** 您可以在sci[i]中指定 limit 参数，以只导出前面部分 ***'
            print '第 %s 条检索条件，其label为 %s '%(i+1,config.sch[i]['label'])
            print '*********************************************\n\n'
            print ''
            continue
        elif rs_count == -1:
            print '\n\n*********************************************'
            print '*** 未找到该检索项，跳过处理 ***'
            print '*** ，请检查该检索项配置，是否有前后空格 ***'
            print '*********************************************\n\n'
            continue
        elif rs_count == False:
            print '\n\n*********************************************'
            print '*** 检索结果失败，adv_search_and_go(...) 返回 False ***'
            print '*** 将跳过处理本项检索条件：第  %s 条，其label为 %s ***'%(i+1,config.sch[i]['label'])
            print '*********************************************\n\n'
            continue
        if config.sch[i].has_key('limit'):
            rs_count=min(rs_count,config.sch[i]['limit'])
        crawl_fail_batchs=0     #连续抓取失败的批次数
        for pos in range(1,rs_count, batch_size):
            if crawl_fail_batchs > config.crawl_fail_retry_limit:
                dmesg('连续抓取失败的批次数达到限制，将跳过，并进入下一组检索条件')
                break
            filename='%s_%s.txt'%(label,pos)
            if os.path.isfile('%s/%s'%(config.file_save_dir,filename)):
                dmesg('文件 %s 已存在，不再重复下载'%filename)
                continue
            pos_to=min(pos+batch_size-1,rs_count)
            dmesg('处理到 %s 的 %.2f%%， 导出范围 [%s,%s] '%(label,100.0*pos/rs_count,pos,pos_to))
            for i in range(1,config.crawl_page_reload_retry+1):
                try:
                    status=crawl(pos,pos_to,label)
                except:
                    dmesg('crawl(..)处理失败')
                    status=False
                if status:
                    crawl_fail_batchs=0
                    sch_crawled_batch_count += 1
                    break
                else:
                    dmesg('下载失败，重试，第 %s 次'%i)
            crawl_fail_batchs += int(status)
        print '*** %s 抓取完成，计成功抓取 %s 批.txt文件 ***\n\n'%(label,sch_crawled_batch_count)
        total_crawled_batch_count += sch_crawled_batch_count
    print '\n***** run() 执行完成。 *****\n共计成功抓取 %s 批.txt文件'%total_crawled_batch_count



if __name__ == '__main__':
    init()
    run()
