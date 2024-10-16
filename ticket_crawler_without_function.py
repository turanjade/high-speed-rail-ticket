import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas
from datetime import date, timedelta, datetime
import pytz

beijing_timezone = pytz.timezone('Asia/Shanghai')


### 定义爬取数据日期。可以为“today”，“tomorrow”，或者具体日期，格式为YYYY-MM-DD
def currentdate(date_input):
    if date_input == "tomorrow": #1 represents tomorrow
        date_return = str(datetime.now(beijing_timezone).date() + timedelta(1))
    elif date_input == "today": #0 represents today
        date_return = str(datetime.now(beijing_timezone).date())
    else:
        date_return = date_input
    return date_return

### 定义爬取数据的时间。如果不定义具体时间，则默认爬取全天数据。 
### 时间对应：t_list = [1,2,3,4,5], 想获取数据的时间段（五个时间段分别对应）全天，0-6，6-12，12-18，18-24，四个时间段
def currenttime(time_input, date_input): 

    #分爬取数据的两种情况
    if date_input == str(datetime.now(beijing_timezone).date()):
        ##如果是今天，即即时数据爬取，提前15分钟结束每个session.否则本session可能由于时间太短爬不出数据.后面可以定义提前多久
        time_input = datetime.strptime(time_input, '%H-%M')

        if time_input <= datetime.strptime("06-00", '%H-%M'):
            timeperiod = 1
        elif time_input > datetime.strptime("06-00", '%H-%M') and time_input <= datetime.strptime("11-45", '%H-%M'): 
            timeperiod = 2
        elif time_input > datetime.strptime("11-45", '%H-%M') and time_input <= datetime.strptime("17:45", '%H-%M'):
            timeperiod = 3
        else:
            timeperiod = 4

    else:
        ##如果不是今天，即未来数据爬取，按正常时间结束session
        time_input = datetime.strptime(time_input, '%H-%M')

        if time_input <= datetime.strptime("06-00", '%H-%M'):
            timeperiod = 1
        elif time_input > datetime.strptime("06-00", '%H-%M') and time_input <= datetime.strptime("12-00", '%H-%M'): 
            timeperiod = 2
        elif time_input > datetime.strptime("12-00", '%H-%M') and time_input <= datetime.strptime("18-00", '%H-%M'):
            timeperiod = 3
        else:
            timeperiod = 4
    return timeperiod



### 输入抓取时间
time_list = datetime.now(beijing_timezone).strftime("%H-%M")
date_list = [currentdate("today")]

#print(time_list)
# print(date_list)

time_i = currenttime(time_list, date_list)
# print(time_i)

# t对应12306官网上对应发车时间区间（1-5）
t_list = [1,2,3,4,5] 
#time_i = 2 ### 改变想要抓取的数据时间段，0-4分别对应：全天，0-6，6-12，12-18，18-24
t = t_list[time_i]


### 输入城市对 - samples 
inputdir = r'D:\\微云同步助手\\332667113\\2023-交科院综合交通开放课题\\研究进度\\软著\\' ### 定义你的输入/输出文件路径
data = pandas.read_excel(str(inputdir + "citypair_sample2.xlsx"),skiprows = 1, header = None)


### 手动输入城市对
cityO = "上海市"
cityD = "南京市"
data = pandas.DataFrame([[cityO, cityD]])
#print(str(data.loc[0,:][0]))


#### 打开Chrome
chrome_option = webdriver.ChromeOptions()
# p=r'C:\Users\lms\AppData\Local\Google\Chrome\User Data'
# chrome_option.add_argument('--user-data-dir='+p)
chrome_option.add_argument("--disable-blink-features=AutomationControlled")
chrome_option.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
chrome_option.add_argument('--ignore-certificate-errors')
chrome_option.add_argument('--ignore-ssl-errors')
browser = webdriver.Chrome(options=chrome_option)
# browser = webdriver.Chrome(executable_path='E:\\HSR-RTU\\chromedriver\chromedriver.exe')
#browser.maximize_window()
browser.minimize_window()
url = 'https://kyfw.12306.cn/otn/leftTicket/init?linktypeid=dc'
browser.get(url)

### d 存储所有票务数据
d = []

while True:
    if len(browser.find_elements(by=By.ID, value='fromStationText')) != 0:
        break
ActionChains(browser).move_to_element(browser.find_element(by=By.ID, value='fromStationText')).perform()
try:
    for n in data.index[:]:
        p = 0
        city = data.loc[n, :].values
        # print(city, n)
        city1 = city[0][:-1] + city[0][-1].replace('市', '')
        city2 = city[1][:-1] + city[1][-1].replace('市', '')
        # city1 = '太原'
        # city2 = '张家口'
        for date in date_list:
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
                    #print(item)
                    d.append(item)
                    continue
                if len(browser.find_elements(by=By.ID, value='cc_seat_type_O')) == 0:
                    item = {}
                    item['搜索日期'] = date
                    item['车次'] = '无'
                    item['出发地'] = city[0]
                    item['目的地'] = city[1]
                    #print(item)
                    d.append(item)
                    continue
                if len(browser.find_elements(by=By.XPATH, value='//tbody[@id="queryLeftTable"]/tr')) == 0:
                    item = {}
                    item['搜索日期'] = date
                    item['车次'] = '无'
                    item['出发地'] = city[0]
                    item['目的地'] = city[1]
                    #print(item)
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
                                                   i.find_elements(by=By.TAG_NAME, value='td')[10],
                                                   i.find_elements(by=By.TAG_NAME, value='td')[11]
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
                    #print(item)
                    #print(i, item['出发地'], item['目的地'])
                    d.append(item)
            except Exception as e:
                item = {}
                item['搜索日期'] = date
                item['车次'] = '无'
                item['出发地'] = city[0]
                item['目的地'] = city[1]
                # print(item)
                d.append(item)
except Exception as e:
    print(e)
    print('出错')
print('citypair finished')

## 导出数据
pandas.DataFrame(d).to_excel(inputdir+date_list[0]+'-'+str(time_list)+'-'+cityO+"-"+cityD+'.xlsx', index=False)

browser.close()

