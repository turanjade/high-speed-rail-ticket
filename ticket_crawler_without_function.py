import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas
from datetime import date, timedelta, datetime
import pytz
import requests


beijing_timezone = pytz.timezone('Asia/Shanghai')


## 给定两个地名，在高德地图API中获取两地距离 -- 用于计算网络指标
# Function to get latitude and longitude of a city
def get_city_coordinates(city_name, api_key):
    url = f'https://restapi.amap.com/v3/geocode/geo?address={city_name}&key={api_key}'
    response = requests.get(url)
    result = response.json()
    
    if result['geocodes']:
        location = result['geocodes'][0]['location']
        return location.split(",")  # Returns [longitude, latitude]
    else:
        return None

# Function to calculate the distance between two coordinates
def get_distance(origin_coords, dest_coords, api_key):
    origins = ','.join(origin_coords)
    destination = ','.join(dest_coords)
    
    url = f'https://restapi.amap.com/v3/distance?origins={origins}&destination={destination}&type=1&key={api_key}'
    response = requests.get(url)
    result = response.json()
    
    if result['results']:
        distance = result['results'][0]['distance']  # Distance in meters
        return distance
    else:
        return None

# Replace with your actual API key
api_key = 'a3088a73b4bedfcf95e01d08933aa701'


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
cityO = ["上海市",'宜兴市','苏州市']
cityD = ["宜兴市",'杭州市','合肥市']
data = pandas.DataFrame({'cityO':cityO, 'cityD':cityD})
print(data)


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
outputpath = str(inputdir+date_list[0]+'-'+str(time_list)+'-'+cityO[0]+"-"+cityD[0]+'.xlsx')
pandas.DataFrame(d).to_excel(outputpath, index=False)

browser.close()

## 针对给定城市对组合，计算复杂网络指标

# 获取城市间距离
def dist_twonames(add_1, add_2):
    # Get coordinates
    origin_coords = get_city_coordinates(add_1, api_key)
    dest_coords = get_city_coordinates(add_2, api_key)

    if origin_coords and dest_coords:
        # Get the distance
        distance = float(get_distance(origin_coords, dest_coords, api_key))/1000
        print(f"{add_1} 和 {add_2} 两地相距: {str(distance)} KM")
    else:
        print("地名坐标获取出错.")
    return(distance)


### 读取票务爬取数据
d = pandas.read_excel(outputpath)

#### 获取出发和到达城市
origin = d['出发城市'].unique()
destination = d['到达城市'].unique()

print(origin)
print(destination)
#print(d)

## Unique 出发站点和到达站点
originS = d['出发地'].unique()
destinationS = d['目的地'].unique()    
print(originS)
print(destinationS)    


### count the ticket number remaining. If it contains ‘有', it means there are sufficient tickets. Otherwise, sum up all the numbers. If no number exists, it is zero
def remainingseats(ticketdb, seatclass):
    abundance = ticketdb[ticketdb[seatclass].str.startswith('有')].shape[0]
    if abundance == 0:
        remaining = pandas.to_numeric(ticketdb[seatclass], errors='coerce').sum()
    else:
        remaining = 10000 # a max value 
    return(remaining)

# average price
def avgprice(ticketdb, seatclass):
    price = pandas.to_numeric(ticketdb[seatclass].str[1:], errors = 'coerce').mean()
    return(price)

# min price
def minprice(ticketdb, seatclass):
    price = pandas.to_numeric(ticketdb[seatclass].str[1:], errors = 'coerce').min()
    return(price)

# HSR where there is not seats
def noseats(ticketdb, seatclass):
    count = ticketdb[ticketdb[seatclass].str.startswith('-')].shape[0] + ticketdb[ticketdb[seatclass].str.startswith('候补')].shape[0]
    return(count)

#print(networkmetric)
networkmetric = []

dij_effi_price = 0
for i in origin:
    #print(i)
    for j in destination:
        #print(j)
        pairs = d[(d['出发城市'] == i) & (d['到达城市'] == j)]
        if len(pairs) > 0:

            #print(pairs)
        ## hsr connection includes G and D
            hsr_connection = pairs[pairs['车次'].str.startswith('G')].shape[0] + pairs[pairs['车次'].str.startswith('D')].shape[0]
            ## if 2nd remaining has the string ’有', it means there are abundant of seats. Otherwise, sum all the numbers
            
            item = {}
            item['CityO'] = i
            item['CityD'] = j
            item['Distance'] = dist_twonames(i,j)
            item['Tot_connection'] = len(pairs)
            item['HSR_connection'] = hsr_connection

            item['2ndClass_remaining'] = remainingseats(pairs, '二等座')
            item['2ndClass_price'] = avgprice(pairs, '二等座票价')
            item['2ndClass_sellout'] = noseats(pairs, '二等座')

            item['1stClass_remaining'] = remainingseats(pairs, '一等座')
            item['1stClass_price'] = avgprice(pairs, '一等座票价')
            item['1stClass_sellout'] = noseats(pairs, '一等座')

            item['BusinessClass_remaining'] = remainingseats(pairs, '商务座')
            item['BusinessClass_price'] = avgprice(pairs, '商务座票价')
            item['BusinessClass_sellout'] = noseats(pairs, '商务座')

            ## shortest path between i and j: here we define as the least price
            item['lowest_price'] = minprice(pairs, '二等座票价')

            ## calculate the number of remaining HSR that has seats available
            item['SeatsAvailableLines'] = len(pairs)-item['2ndClass_sellout'] - item['1stClass_sellout'] + item['BusinessClass_sellout']
                        
            ## calculate redundancy between O and D, taking the lowest price as the shortest path
            item['Redundancy'] = item['SeatsAvailableLines']/item['lowest_price']

            ## calculate the remaining seats of the line
            item['TotRemainingSeats'] = item['2ndClass_remaining'] + item['1stClass_remaining'] + item['BusinessClass_remaining']

            ## accumulate network efficiency regarding the price
            dij_effi_price += minprice(pairs, '二等座')

            ## if we define the shortest path as the 

            # print(item)
            networkmetric.append(item)

print(networkmetric)

pandas.DataFrame(networkmetric).to_excel(inputdir+'networkmetric.xlsx', index=False)


### 其他网络指标计算
### total number of nodes in the network. We define the number of nodes as number of stations
networkmetric = pandas.read_excel(inputdir+'networkmetric.xlsx')

tot_nodes = len(list(set(originS.tolist() + destinationS.tolist())))
print(tot_nodes)
### calculate network density
network_den = len(d)/(tot_nodes*(tot_nodes-1))

### calculate network effective density
# print(networkmetric['SeatsAvailableLines'])
print(pandas.to_numeric(networkmetric['SeatsAvailableLines']))
network_effective_den = pandas.to_numeric(networkmetric['SeatsAvailableLines']).sum()/(tot_nodes*(tot_nodes-1))

### calculate network efficiency using the lowest price of each HSR
network_effi_price = pandas.to_numeric(networkmetric['lowest_price'])/(tot_nodes*(tot_nodes-1))
print(pandas.to_numeric(networkmetric['lowest_price']))

### calculate network efficiency using the remaining seats of all the HSR
network_effi_seats = pandas.to_numeric(networkmetric['TotRemainingSeats']).sum()/(tot_nodes*(tot_nodes-1))
