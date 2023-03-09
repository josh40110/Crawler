from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import re
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import pandas as pd
import requests





op = webdriver.ChromeOptions()
op.add_argument("--start-maximized")
# op.add_argument('headless')
driver= webdriver.Chrome('./chromedriver')
driver.get("https://mops.twse.com.tw/mops/web/t51sb10_q1")
def research(market,keyword,industry,year):

    driver.get("https://mops.twse.com.tw/mops/web/t51sb10_q1")
    time.sleep(1)
    driver.find_element(By.XPATH,'//input[@name="r1"][@value="1"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH,"//select[@name='KIND']").click()
    time.sleep(1)

    #點選市場別
    if market =="O":
        marketpath ='//select[@name="KIND"]//option[@value="O"]'
    else:
        marketpath ='//select[@name="KIND"]//option[@value="L"]'
    driver.find_element(By.XPATH,marketpath).click()#L是上市 O是上櫃
    time.sleep(1)

    #點選產業別
    if industry =="全部":
        industrypath ='//select[@name="CODE"]//option[@value=""]'
    elif industry =="存託憑證":
        industrypath ='//select[@name="CODE"]//option[@value="91"]'
    elif industry =="金融業":
        industrypath ='//select[@name="CODE"]//option[@value="17"]'
    driver.find_element(By.XPATH,industrypath).click()
    time.sleep(1)

    #輸入關鍵字
    driver.find_element(By.XPATH,'//input[@name="keyWord"]').send_keys("經理")
    time.sleep(1)
    driver.find_element(By.XPATH,'//input[@name="keyWord2"]').send_keys(str(keyword))
    time.sleep(1)

    #輸入年分
    driver.find_element(By.XPATH,'//input[@name="year"]').clear()
    driver.find_element(By.XPATH,'//input[@name="year"]').send_keys(year)
    time.sleep(1)

    #點選月份
    path = '//select[@name="month1"]//option[@value="{}"]'.format(0)
    driver.find_element(By.XPATH,path).click()
    time.sleep(1)
    driver.find_element(By.XPATH,'//select[@name="end_day"]//option[@value="31"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH,'//div[@id="search_bar1"]//input[@value=" 搜尋 "]').click()
    # year = 105
    AD_year= int(year)+1911
    data = findDetailRequestInfo(AD_year)
    print(data)
    big_dict = {}
    for k in data[0]:
        big_dict[k] = [''.join(list(d[k])) for d in data]
    try:
        df=pd.read_excel(f'ManagerChange.xlsx',index_col=0) 
        df2= pd.DataFrame(big_dict)
        df2=pd.concat([df,df2],ignore_index=True)
        df2.to_excel(f'ManagerChange.xlsx',encoding='utf_8_sig') 
    except FileNotFoundError:
        df= pd.DataFrame(big_dict)
        df.to_excel(f'ManagerChange.xlsx',encoding='utf_8_sig') 

def findDetailRequestInfo(year):
    time.sleep(5)
    allDataXpath = driver.find_elements(By.XPATH,"//form[@action='/mops/web/ajax_t05st01']//input[@value='詳細資料']")
    detailDatalength = len(allDataXpath)
    test =[]
    for i in range(0,detailDatalength):
        RequestInfo = allDataXpath[i].get_attribute("onclick").split(";")
        for j in range(len(RequestInfo)):
            if("seq_no")in RequestInfo[j] :
                reg = re.search(r'document\.fm\.seq_no\.value="([^"]+)"', RequestInfo[j])
                #這個正則表達式使用了 [^"]+ 來匹配一個或多個非雙引號的字元，並使用了圓括號將這個部分包裝起來，以便後面可以使用 group(1) 方法取出這個部分。
                seq_no_value =reg.group(1)
        for j in range(len(RequestInfo)):
            if("spoke_time")in RequestInfo[j] :
                reg = re.search(r'document\.fm\.spoke_time\.value="([^"]+)"', RequestInfo[j])
                spoke_time_value =reg.group(1)
        for j in range(len(RequestInfo)):
            if("spoke_date")in RequestInfo[j] :
                reg = re.search(r'document\.fm\.spoke_date\.value="([^"]+)"', RequestInfo[j])
                spoke_date_value =reg.group(1)
        for j in range(len(RequestInfo)):
            if(".i.")in RequestInfo[j] :
                reg = re.search(r'document\.fm\.i\.value="([^"]+)"', RequestInfo[j])
                i_value =reg.group(1)
        for j in range(len(RequestInfo)):
            if("co_id")in RequestInfo[j] :
                reg = re.search(r'document\.fm\.co_id\.value="([^"]+)"', RequestInfo[j])
                co_id_value =reg.group(1)
        for j in range(len(RequestInfo)):
            if("TYPEK")in RequestInfo[j] :
                reg = re.search(r'document\.fm\.TYPEK\.value="([^"]+)"', RequestInfo[j])
                TYPEK_value =reg.group(1)
        seq_no=seq_no_value
        spoke_time =spoke_time_value
        spoke_date =spoke_date_value
        i=i_value
        co_id =co_id_value
        tYPEK=TYPEK_value
        NewRequestInfo = {'seq_no':seq_no,'spoke_time':spoke_time,'spoke_date':spoke_date,'i':i,'co_id':co_id,'TYPEK':tYPEK,'year':year}
        time.sleep(2)
        result =getIntoInfo(NewRequestInfo) 
        test.append(result) if result !=None else print("None")
    return test



def getIntoInfo(NewRequestInfo):
    ua = UserAgent()
    headers = {'User-Agent':ua.google}
    url = "https://mops.twse.com.tw/mops/web/ajax_t05st01"
    seq_no = NewRequestInfo['seq_no']
    spoke_time=NewRequestInfo['spoke_time']
    spoke_date=NewRequestInfo['spoke_date']
    year=NewRequestInfo['year']

    i=NewRequestInfo['i']
    co_id=NewRequestInfo['co_id']
    tYPEK =NewRequestInfo['TYPEK']
    payload = {
        "step" :"2",
        "colorchg" :"1",
        "co_id" :co_id,
        "TYPEK" :tYPEK,
        "off" :"1",
        "firstin" :"1",
        "i" :i,
        "year" :year,
        "month" :"1",
        "spoke_date" :spoke_date,
        "spoke_time" :spoke_time,
        "seq_no" :seq_no,
        "b_date" :"1",#開始日期
        "e_date" :"31",#結束日期
        "t51sb10" :"t51sb10",
    }
    res = requests.post(url, data=payload, headers=headers).content
    soup = BeautifulSoup(res, "html.parser")
    try:
        company_code = soup.find('input',{"name":'Q1V'})['value']
    except TypeError:
        return None
    datatime = soup.find('input',{"name":'Q2V'})['value']
    test=  soup.find('pre',{"style":'text-align:left !important; font-family:細明體 !important;'}).text
    print(test)
    title=  soup.find('pre',{"style":'font-family:0�;'}).text

    try:    
        pattern = r"舊任者姓名及簡歷:(.*?)新任者姓名及簡歷"
        match = re.search(pattern, test, re.DOTALL)
        old_manager = match.group(1)
    except AttributeError:
        try:
            pattern = r"舊任者姓名、級職及簡歷:(.*?)新任者姓名、級職及簡歷"
            match = re.search(pattern, test, re.DOTALL)
            old_manager = match.group(1)
        except AttributeError:
            try:
                pattern = r"舊任者姓名:(.*?)舊任者簡歷"
                match = re.search(pattern, test, re.DOTALL)
                old_manager_name = match.group(1)
                pattern = r"舊任者簡歷:(.*?)新任者姓名"
                match = re.search(pattern, test, re.DOTALL)
                old_manager_cv= match.group(1)
                old_manager=old_manager_name+"/"+old_manager_cv
            except AttributeError:
                return None
    try:
        pattern = r"新任者姓名及簡歷:(.*?)異動情形"
        match = re.search(pattern, test, re.DOTALL)
        new_manager = match.group(1)
    except AttributeError:
        try: 
            pattern = r"新任者姓名、級職及簡歷:(.*?)異動情形"
            match = re.search(pattern, test, re.DOTALL)
            new_manager = match.group(1)
        except AttributeError:
            try:
                pattern = r"新任者姓名:(.*?)新任者簡歷"
                match = re.search(pattern, test, re.DOTALL)
                new_manager_name = match.group(1)
                pattern = r"新任者簡歷:(.*?)異動情形"
                match = re.search(pattern, test, re.DOTALL)
                new_manager_cv= match.group(1)
                new_manager=new_manager_name+"/"+new_manager_cv
            except AttributeError:
                return None
    try:
        pattern = r"異動情形（請輸入「辭職」、「職務調整」、「資遣」、「退休」、「死亡」或「新\n任」）:(.*?)異動原因"
        match = re.search(pattern, test, re.DOTALL)
        change_situation = match.group(1)
    except AttributeError:
        try:
            pattern = r"異動情形（請輸入「辭職」、「職務調整」、「資遣」、「退休」、「死亡」、「新\n任」或「解任」）:(.*?)異動原因"
            match = re.search(pattern, test, re.DOTALL)
            change_situation = match.group(1)
        except AttributeError:
            try:
                pattern = r"異動情形（請輸入「辭職」、「解任」、「任期屆滿」、「職務調整」\n、「資遣」、「退休」、「逝世」或「新任」）:(.*?)異動原因"
                match = re.search(pattern, test, re.DOTALL)
                change_situation = match.group(1)
            except AttributeError:
                try:
                    pattern = r"異動情形(.*?)異動原因"
                    match = re.search(pattern, test, re.DOTALL)
                    change_situation = match.group(1)
                except AttributeError:
                    return None

    try:
        pattern = r"異動原因:(.*?)新任生效日期"
        match = re.search(pattern, test, re.DOTALL)
        change_reason = match.group(1)
    except AttributeError:
        pattern = r"異動原因:(.*?)生效日期"
        match = re.search(pattern, test, re.DOTALL)
        change_reason = match.group(1)
    print(f"'標題':{title},'公司代號':{company_code},'西元年':{datatime},'舊任者姓名及簡歷':{old_manager},'新任者姓名及簡歷':{new_manager},'異動情形':{change_situation},'異動原因':{change_reason}")
    return {'標題':{title},'公司代號':{company_code},'西元年':{datatime},'舊任者姓名及簡歷':{old_manager},'新任者姓名及簡歷':{new_manager},'異動情形':{change_situation},'異動原因':{change_reason}}

if __name__ =="__main__":
    for i in range(105,111):
        research("L","退休","全部",str(i))

'''單一表單傳送內容'''
# <input type="hidden" name="step" value="2">
# <input type="hidden" name="colorchg" value="1">
# <input type="hidden" name="co_id" value="">
# <input type="hidden" name="TYPEK" value="">
# <input type="hidden" name="off" value="1">
# <input type="hidden" name="firstin" value="1">
# <input type="hidden" name="i" value="">
# <input type="hidden" name="year" value="2016">
# <input type="hidden" name="month" value="1">
# <input type="hidden" name="spoke_date" value="">
# <input type="hidden" name="spoke_time" value="">
# <input type="hidden" name="seq_no" value="">
# <input type="hidden" name="b_date" value="1">
# <input type="hidden" name="e_date" value="4">
# <input type="hidden" name="t51sb10" value="t51sb10">
# <input type="hidden" name="h00" value="廣積">
# <input type="hidden" name="h01" value="8050">
# <input type="hidden" name="h02" value="20160104">
# <input type="hidden" name="h04" value="公告本公司總經理異動">
# <input type="hidden" name="h03" value="174519">
# <input type="hidden" name="h05" value="1">
# <input type="hidden" name="h06" value="otc">
# <input type="hidden" name="h07" value="25">
# <input type="hidden" name="h08" value="20160104">
# <input type="hidden" name="h09" value="2">

'''clean dataframe'''
# df=pd.read_excel(f'ManagerChange.xlsx',index_col=0) 
# columns= list(df.columns.values)
# stopwords= ['1.','2.','3.','4.','5.','6.','7.','8.','9.']
# newdf = pd.DataFrame()
# for name in columns:
#     newlist = []
#     for data in df[str(name)]:

#         for index in range(0,len(stopwords)):
#             if stopwords[index] in str(data):
#                 newlist.append(str(data).strip(stopwords[index]))
#                 break
#             if index ==(len(stopwords)-1):
#                 newlist.append(str(data))
#     newdf[str(name)]=newlist
# newdf.to_excel(f'ManagerChange.xlsx',encoding='utf_8_sig') 
# new_situation=[]
# for i in newdf['異動情形']:
#     if ":" not in  i:
#         new_situation.append(i)
#     else:
#         i = i.split(":")[1]
#         new_situation.append(i)
# newdf['異動情形']=new_situation
# newdf.to_excel(f'ManagerChange.xlsx',encoding='utf_8_sig') 








