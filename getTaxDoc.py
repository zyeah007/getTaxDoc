#!/usr/bin/env python
# coding=utf-8

import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup as bs
import openpyxl
import random


class taxDocSpider(object):
    def __init__(self):
        # 从国税总局税收政策首页进入
        self.start_url = 'http://www.chinatax.gov.cn/n810341/index.html'

    def getToDocPage(self, driver):
        # 进入到法规库页面
        start_url = self.start_url
        driver.get(start_url)
        time.sleep(random.uniform(15,20))  #等待10秒，让页面完全加载
        # 下一步，获取"检索"按钮
        xpath='//html/body/table[1]/tbody/tr/td[3]/table[1]/tbody/tr[1]/td/table[2]/tbody/tr/td/form/table[7]/tbody/tr/td[1]/table/tbody/tr/td/input'
        search_btn=driver.find_element_by_xpath(xpath)
        search_btn.click()
        print('等待页面加载，60秒...')
        time.sleep(40+random.randrange(1,10))  #这次等待时间较长
        # 此时会在原标签页后面产生新的标签页，需要switch windows到新生成到标签页
        driver.switch_to.window(driver.window_handles[1])
        # 此处可加上断言，确认页面跳转正确！
        return driver

    def getData(self, driver,tax):
        '''

        :param driver: 由getToDocPage函数返回的driver
        :param tax: 要收集的税种名称
        :return: 一个列表，列表中每个元素是一组字典，包括法规的标题、发文日期、文号以及链接、
        '''
        # 休息30秒，等待页面完全加载
        try:
            print('等待30秒...')
            time.sleep(random.randrange(10,21))
            # 转到左侧名叫'leftList'的iframe
            driver.switch_to.frame('leftList')
            tax_btn=driver.find_element_by_link_text(tax)
            time.sleep(random.randrange(10,20))
            tax_btn.click()
            print('再次等待60秒...')
            time.sleep(random.randrange(30,45)) #等待60秒，待该税种下的法规目录完全加载
            # 此处可以再次加入断言，保证没有被网站发现是机器人访问
            driver.switch_to.parent_frame()
            time.sleep(random.randrange(10,20))
            driver.switch_to.frame('rightList') #切入到主数据所在到frame中
            tbody=driver.find_element_by_xpath('//html/body/form/table/tbody/tr/td/table[4]/tbody')
            trs=tbody.find_elements_by_xpath('.//tr')
            out=[]  # 设定输出结果为一个列表，初始值为空
            for tr in trs[1:]: #从第2行开始。因为第一行为标题
                link=tr.find_element_by_xpath('.//a[1]').get_attribute('href')
                title=tr.find_element_by_xpath('.//a[1]').text.strip()
                date=tr.find_element_by_xpath('.//td[2]').text.strip()
                docCode=tr.find_element_by_xpath('.//td[3]').text.strip()
                out.append({'title':title,'date':date,'docCode':docCode,'link':link})
            return out
        except:
            print('访问失败！')
            # driver.quit()

    def crawl(self,tax):
        driver=webdriver.Firefox(executable_path='/Users/zhengye/软件/geckodriver')
        driver2=self.getToDocPage(driver)
        out=self.getData(driver2,tax=tax)
        for line in out:
            print(out)
        print('done.')

if __name__=='__main__':
    spider=taxDocSpider()
    spider.crawl('增值税')