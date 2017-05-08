import openpyxl
from bs4 import BeautifulSoup
import re
import pyautogui, time
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait


wb = openpyxl.load_workbook('Master Sheet.xlsx')
wb.get_sheet_names()
MasterSheet= wb.get_sheet_by_name('Sheet1')

i=0
Zipholder = []

for rowcellObj in MasterSheet['A2':MasterSheet.max_row]:
    for cellObj in rowcellObj:
        #print(cellObj.value)
        Zipholder.append(cellObj.value)
        i=i+1
        
print('---End of Zip Codes---')


#Takes zips from Data structure and checks DG site to see if New Low Pricing is available in that Zip

#user_agent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.57 Safari/537.36"
#dcap = dict(DesiredCapabilities.PHANTOMJS)
#dcap["phantomjs.page.settings.userAgent"] = user_agent

browser = webdriver.Chrome()
browser.get('http://www2.dollargeneral.com/Savings/Circulars/Pages/index.aspx')
browser.maximize_window()
browser.switch_to.frame("flipp-iframe")

x=0

zipElem = browser.find_element_by_xpath('//*[@id="postal_code_input"]')

print(Zipholder[x])
zipElem.send_keys(Zipholder[x])
zipElem.submit()


Selectzip = browser.find_element_by_xpath('//*[@id="store_select_area"]/div/div/table/tbody/tr[1]/td[3]/form/button')
Selectzip.click()

WeeklyAdzip = browser.find_element_by_xpath('//*[@id="other_flyer_runs"]/div/div/div/div[2]/table/tbody/tr[2]/td[1]/div/div/img')
WeeklyAdzip.click()
nlp=0
NewzipHolder =[[x],[nlp]]

TestValue = 'DG Digital Coupon'

NewzipHolder = [0]*len(Zipholder)

while x <= len(Zipholder):


    SelectStore = browser.find_element_by_xpath('//*[@id="wishabi-flyerarea"]/div[2]/div/div[1]/div[2]/div[1]/div/div/div/div/div/h4')
    SelectStore.click()
    ChangeStore = browser.find_element_by_xpath('//*[@id="wishabi-flyerarea"]/div[2]/div/div[1]/div[2]/div[2]/div[3]')
    ChangeStore.click()

    try:
        zipElem = browser.find_element_by_xpath('//*[@id="postal_code_input"]')
        zipElem.click()
        zipElem.click()
        zipElem.click()
        print("1 try "+str(Zipholder[x]))
        time.sleep(1)
        pyautogui.typewrite(['backspace','backspace','backspace','backspace','backspace','backspace','backspace'])
        pyautogui.hotkey('ctrl','a')
        zipElem.send_keys(Zipholder[x])
        zipElem.submit()

        try:

            Selectzip = browser.find_element_by_xpath('//*[@id="store_select_area"]/div/div/table/tbody/tr[1]/td[3]/form/button')
            Selectzip.click()

            try:
                WeeklyAdzip = browser.find_element_by_xpath('//*[@id="other_flyer_runs"]/div/div/div/div[2]/table/tbody/tr[2]/td[1]/div/div/img')
                WeeklyAdzip.click()

                wait = WebDriverWait(browser, 10)
                html = browser.page_source
                soup = BeautifulSoup(html, "html.parser")
                soup_string = str(soup)
                matches = re.findall(TestValue, soup_string)

                if matches:
                    NewzipHolder[x]=Zipholder[x],"True"

                else:
                    NewzipHolder[x] = Zipholder[x],"False"

                print(x)

            except:
                wait = WebDriverWait(browser, 10)
                html = browser.page_source
                soup = BeautifulSoup(html, "html.parser")
                soup_string = str(soup)
                matches = re.findall(TestValue, soup_string)

                if matches:
                    NewzipHolder[x] = Zipholder[x],"True"

                else:
                    NewzipHolder[x] = Zipholder[x], "False"

        except:
            pass
    except:
        pass
    print(x)
    x = x + 1

print (NewzipHolder)
