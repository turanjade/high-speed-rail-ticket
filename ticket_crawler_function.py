import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas
from datetime import date, timedelta, datetime

## 打开chrome，开始搜索。调用的browser driver应当满足调用要求。详情可在谷歌chrome官方网页上下载。
chrome_option = webdriver.ChromeOptions()
# p=r'C:\Users\lms\AppData\Local\Google\Chrome\User Data'
# chrome_option.add_argument('--user-data-dir='+p)

chrome_option.add_argument("--disable-blink-features=AutomationControlled")
chrome_option.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
chrome_option.add_argument('--ignore-certificate-errors')
chrome_option.add_argument('--ignore-ssl-errors')
browser = webdriver.Chrome(options=chrome_option)

#若默认版本chromedriver不支持，则需要在chrome管网下载支持的版本，放到指定文件夹中使用。此处仅为范例
# browser = webdriver.Chrome(executable_path='E:\\HSR-RTU\\chromedriver\chromedriver.exe') 
# browser.maximize_window() #放大窗口观察票务数据爬取实施情况
browser.minimize_window() #最小化窗口
url = 'https://kyfw.12306.cn/otn/leftTicket/init?linktypeid=dc' #url是12306爬取数据的地址
browser.get(url)


### 定义O-D城市ID赋值function. 将市、县去掉
def cityrename(city):
    if "县" in city:
        city_rename = city[:-1] + city[-1].replace('县', '')
    elif "市" in city:
        city_rename = city[:-1] + city[-1].replace('市', '')
    return city

### 定义爬取数据的时间。如果不定义具体时间，则默认爬取全天数据。 
### 时间对应：t_list = [1,2,3,4,5], 想获取数据的时间段（五个时间段分别对应）全天，0-6，6-12，12-18，18-24，四个时间段
def currenttime(time, date): 
    #分爬取数据的两种情况
    if date == str(date.today()):
        ##如果是今天，即即时数据爬取，提前15分钟结束每个session.否则本session可能由于时间太短爬不出数据.后面可以定义提前多久
        time = datetime.strptime(time, '%H:%M')
        if time <= datetime.strptime("06:00", '%H:%M'):
            timeperiod = 1
        elif time > datetime.strptime("06:00", '%H:%M') and time <= datetime.strptime("11:45", '%H:%M'): 
            timeperiod = 2
        elif time > datetime.strptime("11:45", '%H:%M') and time <= datetime.strptime("17:45", '%H:%M'):
            timeperiod = 3
        else:
            timeperiod = 4
    else:
        ##如果不是今天，即未来数据爬取，按正常时间结束session
        time = datetime.strptime(time, '%H:%M')
        if time <= datetime.strptime("06:00", '%H:%M'):
            timeperiod = 1
        elif time > datetime.strptime("06:00", '%H:%M') and time <= datetime.strptime("12:00", '%H:%M'): 
            timeperiod = 2
        elif time > datetime.strptime("12:00", '%H:%M') and time <= datetime.strptime("18:00", '%H:%M'):
            timeperiod = 3
        else:
            timeperiod = 4
    return timeperiod

### 定义爬取数据日期。可以为“today”，“tomorrow”，或者具体日期，格式为YYYY-MM-DD
def currentdate(date):
    if date == "tomorrow":
        date_return = str(date.today() + timedelta(1))
    elif date == "today":
        date_return = str(date.today())
    else:
        date_return = date
    return date_return


### 定义票务爬取程序
##### 定义起终城市，日期和时间(timeperiod, 0-5)
def citypairticket(cityO, cityD, date, t): 
    while True:
        if len(browser.find_elements(by=By.ID, value='fromStationText')) != 0: #从stationID可搜寻数据为0，则跳出
            break

    ActionChains(browser).move_to_element(browser.find_element(by=By.ID, value='fromStationText')).perform()

    ### enter original/departing city。定义出发城市输入chrome
    cityO = cityrename(cityO)
    cityD = cityrename(cityD)

    cityO_input = browser.find_element(by = By.ID, value = "fromStationText")
    cityO_input.clear()
    cityO_input.send_keys(cityO)

    while True:
        if len(browser.find_elements(by = By.ID, value = 'citem_0')) != 0:
                break
    p = 0
    for i in browser.find_elements(by=By.CLASS_NAME, value='cityline'):
        if cityO == i.text.split()[0]:
            i.click()
            p = 1
            break
    
    if p == 0:
        cityO_input.send_keys(Keys.ENTER)
    
    
    ### enter destination/arriving city。定义到达城市输入chrome
    cityD_input = browser.find_element(by = By.ID, value = "toStationText")
    cityD_input.clear()
    cityD_input.send_keys(cityD)

    while True:
        if len(browser.find_elements(by = By.ID, value = 'citem_0')) != 0:
                break
    p = 0
    for i in browser.find_elements(by=By.CLASS_NAME, value='cityline'):
        if cityD == i.text.split()[0]:
            i.click()
            p = 1
            break
    
    if p == 0:
        cityD_input.send_keys(Keys.ENTER)
    
    ### 输入查询日期
    date_input = browser.find_element(by = By.ID, value = "train_date")
    date_input.clear()
    date_input.send_keys(date)
    date_input.send_keys(Keys.ENTER)
    browser.execute_script("arguments[0].click();", browser.find_element(by=By.ID, value='query_ticket'))
    time.sleep(0.3) ## 避免ip访问过于频繁

    d = [] ###d 用于存储所有爬取的数据

    ### 以下情况都没有票务返回
    p = 0
    for i in range(50):
        if len(browser.find_element(by = By.ID, value = "no_filter_ticket_2").get_attribute("style")) == 0:
            p = 1
            break
        if len(browser.find_elements(by = By.XPATH, value = '//tbody[@id="queryleftTable"]/tr')) != 0:
            break

    if p == 1:
        item = {}
        item['搜索日期'] = date
        item['车次'] = '无'
        item['出发地'] = city[0]
        item['目的地'] = city[1]
        print(item)
        d.append(item)
        #continue

    if len(browser.find_elements(by=By.ID, value='cc_seat_type_O')) == 0:
        item = {}
        item['搜索日期'] = date
        item['车次'] = '无'
        item['出发地'] = city[0]
        item['目的地'] = city[1]
        print(item)
        d.append(item)
        #continue

    if len(browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr')) == 0:
        item = {}
        item['搜索日期'] = date
        item['车次'] = '无'
        item['出发地'] = city[0]
        item['目的地'] = city[1]
        print(item)
        d.append(item)
        #continue

    ### 
    browser.find_element(by=By.XPATH, value=f'//*[@id="cc_start_time"]/option[{t}]').click() ### t, timeperiod 在这里
    time.sleep(0.3)
    tr_list = browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr')
    number_list = browser.find_elements(by=By.CLASS_NAME, value='number')
    s_list = browser.find_elements(by=By.CLASS_NAME, value='cdz')
    t_list = browser.find_elements(by=By.CLASS_NAME, value='cds')
    c = 0

    try:
        #### 更新tr_list列表
        new_tr_list = []
        for i in tr_list:
            try:
                for j in i.find_elements(by=By.TAG_NAME, value='td')[1:]:
                    if j.text != '候补': ## 查询出候补票
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

        #### 票价列表
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
                            i.find_elements(by=By.TAG_NAME, value='td')[10],
                            i.find_elements(by=By.TAG_NAME, value='td')[11]
                            ])
                break
            except:
                pass   
        
        ### 更新票务数据列表
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
            item['优选一等座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[2].text
            item['优选一等座票价'] = price_list[i][1].text
            item['一等座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[3].text
            item['一等座票价'] = price_list[i][2].text
            item['二等座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[4].text
            item['二等座票价'] = price_list[i][3].text
            item['高级软卧'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[5].text
            item['高级软卧票价'] = price_list[i][4].text
            item['软卧'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[6].text
            item['软卧票价'] = price_list[i][5].text
            item['硬卧'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[7].text
            item['硬卧票价'] = price_list[i][6].text
            item['软座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[8].text
            item['软座票价'] = price_list[i][7].text
            item['硬座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[9].text
            item['硬座票价'] = price_list[i][8].text
            item['无座'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[10].text
            item['无座票价'] = price_list[i][9].text
            # print(item)
            item['其他'] = new_tr_list[i].find_elements(by=By.TAG_NAME, value='td')[11].text
            item['其他票价'] = price_list[i][10].text
            print(item)
            #print(i, item['出发地'], item['目的地'])
            d.append(item)

    except Exception as e:
        item = {}
        item['搜索日期'] = date
        item['车次'] = '无'
        item['出发地'] = city[0]
        item['目的地'] = city[1]
        print(item)
        d.append(item)
    
    return d


### 尝试使用上述function挖掘一整个citypair list，明天一整天的票务数据
## 首先定义运行位置
date = currentdate("tomorrow")
time = 0 ###全天
# time = currenttime("8:15") ###如果想用某一刻时间，需要参考左边的表达
## 针对样本数据，可采取已有列举城市对进行数据挖掘（即citypair_dir_1.xlsx. 对于任意城市对，可输入任意城市对为city1和city2，代入方程中进行票务检索。
cities = pandas.read_excel('citypair_dir_1.xlsx')
ticket = []

for n in cities.index[:]:
    city = cities.loc[n, :].values
    print(city)
    print(n)
    cityO = cityrename(city[0])
    cityD = cityrename(city[1])
    try:
        ticket_od = citypairticket(cityO, cityD, date, time)
    except Exception as e:
        print(e)
        print('出错')
    ticket.append(ticket_od)

pandas.DataFrame(ticket).to_excel(date+'_allday_citypair1'+'.xlsx', index=False) # 保存数据

browser.close()

### 尝试使用上述function挖掘明天6：00-12：00上海市-苏州市的票务数据
date = currentdate("tomorrow")
time = 1 #明天上午6：00-12：00，不需要输入实时时间
ticket = []

try:
    ticketsh_sz = citypairticket("上海市", "苏州市", date, time)
except Exception as e:
    print(e)
    print("出错")

print(ticketsh_sz)