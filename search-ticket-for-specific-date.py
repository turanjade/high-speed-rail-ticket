import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas
from datetime import date

data = pandas.read_excel('城市对.xlsx') #need to prepare ahead
d = []
chrome_option = webdriver.ChromeOptions()
# p=r'C:\Users\lms\AppData\Local\Google\Chrome\User Data'
# chrome_option.add_argument('--user-data-dir='+p)
chrome_option.add_argument("--disable-blink-features=AutomationControlled")
chrome_option.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
browser = webdriver.Chrome(options=chrome_option)
#browser.maximize_window()
url = 'https://kyfw.12306.cn/otn/leftTicket/init?linktypeid=dc'
date_today = str(date.today()) ###get the date of today
date_tmr = str(date.today() + 1) ###get the date of today
start_time_list = ['00:00--06:00', '06:00--12:00', '12:00--18:00', '18:00--24:00'] #search based on the starting time.

browser.get(url)
while True:
    if len(browser.find_elements(by=By.ID, value='fromStationText')) != 0:
        break
ActionChains(browser).move_to_element(browser.find_element(by=By.ID, value='fromStationText')).perform()
try:
    for n in data.index[:]:
        p = 0
        city = data.loc[n, :].values
        print(city)
        print(n)
        city1 = city[0][:-1] + city[0][-1].replace('市', '') #origin
        city2 = city[1][:-1] + city[1][-1].replace('市', '') #destination
        for date in date_today and time in start_time_list: #or in date_tmr
            try:
                city1_input = browser.find_element(by=By.ID, value='fromStationText')
                city1_input.clear()
                city1_input.send_keys(city1)
                while True:
                    if len(browser.find_elements(by=By.ID, value='citem_0')) != 0:
                        break
                p = 0
                for i in browser.find_elements(by=By.CLASS_NAME, value='cityline'):
                    try:
                        if city1 == i.text.split()[0]:
                            i.click()
                            p = 1
                            break
                    except:
                        pass
                if p == 0:
                    city1_input.send_keys(Keys.ENTER)
                city2_input = browser.find_element(by=By.ID, value='toStationText')
                # ActionChains(browser).move_to_element(city2_input).perform()
                city2_input.clear()
                city2_input.send_keys(city2)
                while True:
                    if len(browser.find_elements(by=By.ID, value='citem_0')) != 0:
                        break
                p = 0
                for i in browser.find_elements(by=By.CLASS_NAME, value='cityline'):
                    try:
                        if city2 == i.text.split()[0]:
                            i.click()
                            p = 1
                            break
                    except:
                        pass
                if p == 0:
                    city2_input.send_keys(Keys.ENTER)
                date_input = browser.find_element(by=By.ID, value='train_date')
                date_input.clear()
                date_input.send_keys(date)
                date_input.send_keys(Keys.ENTER)
                browser.execute_script('arguments[0].click();', browser.find_element(by=By.ID, value='query_ticket'))
                time.sleep(0.3)
                p = 0
                for i in range(50):
                    if len(browser.find_element(by=By.ID, value='no_filter_ticket_2').get_attribute('style')) == 0:
                        p = 1
                        break
                    if len(browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr')) != 0:
                        break
                if p == 1:
                    item = {}
                    item['搜索日期'] = date
                    item['车次'] = '无'
                    item['出发地'] = city[0]
                    item['目的地'] = city[1]
                    print(item)
                    d.append(item)
                    continue
                if len(browser.find_elements(by=By.ID, value='cc_seat_type_O')) == 0:
                    item = {}
                    item['搜索日期'] = date
                    item['车次'] = '无'
                    item['出发地'] = city[0]
                    item['目的地'] = city[1]
                    print(item)
                    d.append(item)
                    continue
                if len(browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr')) == 0:
                    item = {}
                    item['搜索日期'] = date
                    item['车次'] = '无'
                    item['出发地'] = city[0]
                    item['目的地'] = city[1]
                    print(item)
                    d.append(item)
                    continue
                # browser.execute_script('arguments[0].click();', browser.find_element(by=By.ID, value='cc_seat_type_O'))
                browser.find_element(by=By.XPATH, value=f'//*[@id="cc_start_time"]/option[{t}]').click()
                time.sleep(0.3)
                tr_list = browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr')
                number_list = browser.find_elements(by=By.CLASS_NAME, value='number')
                s_list = browser.find_elements(by=By.CLASS_NAME, value='cdz')
                t_list = browser.find_elements(by=By.CLASS_NAME, value='cds')
                c = 0
                new_tr_list = []
                for i in tr_list:
                    try:
                        for j in i.find_elements(by=By.TAG_NAME, value='td')[1:]:
                            if j.text != '候补':
                                j.click()
                                break
                        c += 1
                    except:
                        pass
                    try:
                        if len(i.find_element(by=By.TAG_NAME, value='td').text.strip()) != 0:
                            new_tr_list.append(i)
                    except:
                        pass
                time.sleep(0.3)
                price_list = []
                for i in browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr'):
                    try:
                        for j in i.find_elements(by=By.TAG_NAME, value='td'):
                            if j.get_attribute('class') == 'p-num':
                                price_list.append([i.find_elements(by=By.TAG_NAME, value='td')[1],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[2],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[3],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[4],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[5],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[6],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[7],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[8],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[9],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[10]
                                                   ])
                                break
                    except:
                        pass
                for i in range(len(number_list)):
                    item = {}
                    item['搜索日期'] = date
                    item['车次'] = number_list[i].text
                    item['出发地'] = s_list[i].find_elements(by=By.TAG_NAME, value='strong')[0].text
                    item['目的地'] = s_list[i].find_elements(by=By.TAG_NAME, value='strong')[1].text
                    item['出发时间'] = t_list[i].find_elements(by=By.TAG_NAME, value='strong')[0].text
                    item['到达时间'] = t_list[i].find_elements(by=By.TAG_NAME, value='strong')[1].text
                    item['商务座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[1].text
                    item['商务座票价'] = price_list[i][0].text
                    item['一等座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[2].text
                    item['一等座票价'] = price_list[i][1].text
                    item['二等座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[3].text
                    item['二等座票价'] = price_list[i][2].text
                    item['高级软卧'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[4].text
                    item['高级软卧票价'] = price_list[i][3].text
                    item['软卧'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[5].text
                    item['软卧票价'] = price_list[i][4].text
                    item['硬卧'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[6].text
                    item['硬卧票价'] = price_list[i][5].text
                    item['软座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[7].text
                    item['软座票价'] = price_list[i][6].text
                    item['硬座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[8].text
                    item['硬座票价'] = price_list[i][7].text
                    item['无座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[9].text
                    item['无座票价'] = price_list[i][8].text
                    item['其他'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[10].text
                    item['其他票价'] = price_list[i][9].text
                    print(item)
                    d.append(item)
            except:
                item = {}
                item['搜索日期'] = date
                item['车次'] = '无'
                item['出发地'] = city[0]
                item['目的地'] = city[1]
                print(item)
                d.append(item)
except Exception as e:
    print(e)
    print('出错')
pandas.DataFrame(d).to_excel(str(date_today, '.xlsx'), index=False)
chromeDriver.close()
