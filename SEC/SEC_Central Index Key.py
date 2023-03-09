import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import re
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import pandas as pd
import requests
from selenium.common.exceptions import StaleElementReferenceException

'''This part is about performing a perfect comparison of company names between the 
source Excel "df" and the FI Excel "objectdf" ,export ['Cust_cik_number']'''
# df = pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx")
# objectdf = pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\financial inform.xlsx")
# compareFI =[]
# dfcompanyList= set(df['Cust_name'].values)
# objectdfcompanyList= set(objectdf['Company Name'].values)
# #比對兩表都有的公司
# findcompany = []
# for i in dfcompanyList:
#     for j in objectdfcompanyList:
#         if i ==j:
#             findcompany.append(str(i))
#             break
# for dfcompany in df['Cust_name']:
#     if dfcompany in findcompany:
#         #有在比對符合的清單裡
#         for objectIndex in range(0,len(objectdf['Company Name'])):
#             if dfcompany==objectdf['Company Name'][objectIndex]:
#                 if str(objectdf['CIK Number'][objectIndex]) != 'nan':
#                     CIK = objectdf['CIK Number'][objectIndex]
#                     zero_length = "0"*(10-len(str(int(CIK))))
#                     compareFI.append(str(zero_length+str(int(CIK))))
#                     print(f'index:{dfcompany},found,CIK:{str(zero_length+str(CIK))}')
#                     break
#                 else:
#                     compareFI.append("nan")
#                     break
#     else:
#         compareFI.append(0)
# df['Cust_cik_number']=compareFI
# df.to_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx")


'''source Excel removes stopwords ,export ['strip_Cust_name']&['stopwords']'''
# df = pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx",index_col=0)
# stopwordsdf  = list(pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\組織名.xlsx")['llc'])
# df_striped =[]
# df_striped_words =[]
# for i in range(0,len(df)):
#     name =  df['Cust_name'][i]
#     if (df['Cust_cik_number'][i]==0.0):#去除掉暫留字可能在名子裡面的公司名稱
#         if (name.split(" ")[0]!=name):
#             for stopwordIndex in range(len(stopwordsdf)):
#                 if stopwordsdf[stopwordIndex] == name.split(" ")[-1]:
#                     stripname =name[0:-len(stopwordsdf[stopwordIndex])]
#                     stripname = stripname.strip(", ")
#                     df_striped.append([name,stripname,stopwordsdf[stopwordIndex]])
#                     break
#                 elif (stopwordIndex==(len(stopwordsdf)-1)):
#                     df_striped.append([name,name,None])#純找全名
#         else:
#             df_striped.append([name,name,None])#純找全名
#     else:
#         df_striped.append([None,None,None])#有資料
# df_striped_words = [i[1] for i in df_striped]
# stopwords = [i[2] for i in df_striped]
# df['strip_Cust_name']=df_striped_words
# df['stopwords']=stopwords
# df.to_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx")

'''If the company name in the source is to be web crawled, no web crawling will be performed 
if the ['Cust_cik_number'] column for that record has a value greater than 0.0 ,export [compare]'''
#driver
df = pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx")
stopwordsdf  = list(pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\組織名.xlsx")['llc'])
op = webdriver.ChromeOptions()
op.add_argument("--start-maximized")
# op.add_argument('headless')
driver= webdriver.Chrome(r"C:\Users\joshAnn\Desktop\Project\Python\chromedriver.exe")
driver.get("https://www.sec.gov/edgar/search/#")
driver.find_element(By.XPATH,'//a[@id="show-full-search-form"]').click()
compare =[]#cik,situation

for i in range(19308, len(df['strip_Cust_name'])):
    print(i)
    if df['Cust_cik_number'][i]!=0.0:#如果cik有資料
        print("Has Value")
        compare.append(["Has Value","Has Value"])
        continue
    if (i!=0) and(df['strip_Cust_name'][i] ==df['strip_Cust_name'][i-1]):#如果和上一行長一樣
        compare.append([compare[i-1][0],compare[i-1][1]])
        continue
    if (i!=0):
        if df['Cust_cik_number'][i-1]==0.0:
            if  (len(df['strip_Cust_name'][i]) ==len(df['strip_Cust_name'][i-1])):#如果和上一行長度一樣
                compare.append([compare[i-1][0],compare[i-1][1]])
                continue
    if df['Cust_cik_number'][i]==0.0:
        time.sleep(2)
        driver.find_element(By.XPATH,'//input[@id="entity-full-form"]').clear()
        time.sleep(2)
        driver.find_element(By.XPATH,'//input[@id="entity-full-form"]').send_keys(df['strip_Cust_name'][i])
        time.sleep(2)
        #related search
        try:
            newRelaList=[]
            test =driver.find_elements(By.XPATH,'//div[@class="entity-hints border border-dark border-top-0 rounded-bottom"]//table[@id="asdf"]//tr')#如果出現複數related words
            if test!=[]:
                RelaList =[]
                for relatedData in test:
                    RelaList.append([relatedData.text.split("CIK")[0].split("(")[0],relatedData.text.split("CIK")[1]])
                for name,CIK in RelaList:
                    if (name.split(" ")[0]!=name):
                        for stopwordIndex in range(len(stopwordsdf)):
                            if stopwordsdf[stopwordIndex] == name.split(" ")[-1]:
                                stripname =name[0:-len(stopwordsdf[stopwordIndex])]
                                stripname = stripname.strip(", ")
                                newRelaList.append([stripname,CIK])
                                break
                            elif (stopwordIndex==(len(stopwordsdf)-1)):
                                newRelaList.append([name,CIK])#純找全名
                    else:
                        newRelaList.append([name,CIK])#純找全名
            else:
                compare.append(['cannot find related keyword','cannot find related keyword'])
        except StaleElementReferenceException:
            time.sleep(2)
            test =driver.find_element(By.XPATH,'//div[@class="entity-hints border border-dark border-top-0 rounded-bottom"]//table[@id="asdf"]//tr')#如果出現非複數related words
            newRelaList=[]
            if test!=[]:
                RelaList =[]
                for relatedData in test:#放入related words和CIK
                    RelaList.append([relatedData.text.split("CIK")[0].split("(")[0],relatedData.text.split("CIK")[1]])
                for name,CIK in RelaList:
                    if (name.split(" ")[0]!=name):
                        for stopwordIndex in range(len(stopwordsdf)):
                            if stopwordsdf[stopwordIndex] == name.split(" ")[-1]:
                                stripname =name[0:-len(stopwordsdf[stopwordIndex])]
                                stripname = stripname.strip(", ")
                                newRelaList.append([stripname,CIK])
                                break
                            elif (stopwordIndex==(len(stopwordsdf)-1)):
                                newRelaList.append([name,CIK])#純找全名
                    else:
                        newRelaList.append([name,CIK])#純找全名
            else:
                compare.append(['cannot find related keyword','cannot find related keyword'])
        for resultIndex in range(len(newRelaList)):
            result = newRelaList[resultIndex]
            time.sleep(2)
            if result == df['strip_Cust_name'][i]:#如果related words = source Compname
                driver.find_element(By.XPATH,'//input[@id="entity-full-form"]').clear()
                time.sleep(2)
                driver.find_element(By.XPATH,'//input[@id="entity-full-form"]').send_keys(df['strip_Cust_name'][i])
                time.sleep(2)
                driver.find_element(By.XPATH,'//button[@id="search"]').click()
                time.sleep(4)
                text = driver.find_element(By.XPATH,'//div[@id="results"]').text
                if "No results" in text:
                    compare.append([result[0],result[1]])
                else:
                    compare.append([result[0],result[1]])
                break
            else:
                driver.find_element(By.XPATH,'//input[@id="entity-full-form"]').clear()
                time.sleep(2)
                driver.find_element(By.XPATH,'//input[@id="entity-full-form"]').send_keys(result[0])
                time.sleep(2)
                driver.find_element(By.XPATH,'//button[@id="search"]').click()
                time.sleep(4)
                text = driver.find_element(By.XPATH,'//div[@id="results"]').text
                if "No results" in text:
                    compare.append([result[0],result[1]])
                else:
                    compare.append([result[0],result[1]])
                break
    print(compare[i])

for i in compare:
    print(i)
newCIK = [i[1] for i in compare]
newName = [i[0] for i in compare]
df['CrawlCIK'] =newCIK
df['CrawlName'] =newName
df.to_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx")


'''source Compare FI , step1 The first step in comparing the source Excel with the FI Excel 
    is to export the non-duplicated FI company names,[objectdictionary].'''
df = pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx",index_col=0)
objectdf = pd.read_excel(r"C:\Users\joshAnn\Desktop\獻良\financial inform.xlsx")
#All to uppercase
newobjectdf = list([[i,objectdf['Company Name'][i]] for i in range(len(objectdf['Company Name']))])
objectdictionary = []
for index,test in newobjectdf:
    newword = ''.join([i for i in test if i.isalpha()])
    newword = newword.lower()
    objectdictionary.append([index,newword])

'''The second step involves comparing the similarity of each character in the strings!!! 
    In general, a similarity score of at least 50 is required for a successful match. 

    Let A be the source company name and B be the FI company name. 
    Two scenarios may arise: if the length of string A is less than that of string B, 
    a loop will be used to compare the characters of the two strings until the end of string A 
    is reached. If any character in string B fails to match during the comparison, 
    the similarity score will be checked to see if it is greater than or equal to 50. 
    If not, the comparison will be broken. Once the comparison is complete without any errors, 
    the similarity score will be output in 「descending order from 100 to 50」 to ensure the best match is 
    found. 
    In the second scenario, if the length of string B is less than that of string A, 
    string B will end first, and the similarity score of string B will be checked to 
    determine whether to output the result, provided that the similarity score is 
    greater than or equal to 50.'''
totalscoreList=[]
for test in range(0,len(df['Cust_name'])):
    if df['Cust_cik_number'][test]!=0.0:
        totalscoreList.append([df['Cust_name'][test],"Has Value"])
        continue
    print(test)
    newword = ''.join([i for i in df['Cust_name'][test] if (i.isalpha())or(i.isnumeric())])
    newword = newword.lower()#df word
    score = len(newword)#100分
    scoreList=[]
    for index,object in objectdictionary:
        count = 0
        for l in range(0,score):
            try:
                if newword[l]==object[l]:
                    count+=1
                else:
                    if ((count/score)*100<=50):#開始出現錯誤 比對率小於50%就跳過
                        break
                    if ((count/score)*100>=50):
                        scoreList.append([index,object,count])#開始出現錯誤 比對率大於50%直接塞進去
                        break

                if l ==(score-1):#比對字串字元數>source長度 且比對完 看有沒有50%
                    if ((count/score)*100>=50):
                        scoreList.append([index,object,count])
                        break
            except IndexError:#blockhrinc 比較小的blockh去比對(O) 有500分但字元數比較少
                if ((count/score)*100>=50):
                    scoreList.append([index,object,count])
                    break
    result = sorted(scoreList,key=lambda x: x[2],reverse=True)
    totalscoreList.append([df['Cust_name'][test],result])
resultList =[]
for i in totalscoreList:
    if (i[1]==[]):
        resultList.append(["Doesn't Compared","Doesn't Compared"])
        continue
    if i[1]=="Has Value":
        resultList.append(["Has Value","Has Value"])
    else:
        name = objectdf['Company Name'][i[1][0][0]]
        CIK = objectdf['CIK Number'][i[1][0][0]]
        resultList.append([name,CIK])

df['CompareExcelName']=[i[0] for i in resultList]
df['CompareExcelCIK']=[i[1] for i in resultList]
df.to_excel(r"C:\Users\joshAnn\Desktop\獻良\0301.xlsx")






