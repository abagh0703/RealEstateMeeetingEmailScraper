import mechanize
from bs4 import BeautifulSoup
import requests
import xlsxwriter
from operator import itemgetter
import time

from selenium import webdriver
import selenium
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

url="http://www.icsc.org/2015EDM/login"
people=[]
pageTemplate="num="
curPage=0
count=0
#mechanize method
br = mechanize.Browser()
br.set_handle_redirect(True)
br.addheaders = [('User-agent', 'Firefox')]
br.set_handle_robots(False)   # ignore robots
br.set_handle_refresh(False)  # can sometimes hang without this
br.open(url)

driver = webdriver.Firefox()
driver.get(url)
driver.find_elements_by_name('icsc_email_id')[1].send_keys('')
driver.find_elements_by_name('icsc_password')[1].send_keys('')
driver.find_elements_by_name('icsc_password')[1].submit()
#driver.get('http://www.icsc.org/2015EDM/whos-coming/attendees/?paginate=members&filter=YTo3OntzOjU6ImxpbWl0IjtzOjI6IjI1IjtzOjEwOiJtYXhyZXN1bHRzIjtzOjI6IjI1IjtzOjE0OiJzaG93LWNvbXBhbmllcyI7czoyOiJubyI7czo4OiJvcmRlcl9ieSI7czo5OiJsYXN0X25hbWUiO3M6NDoic29ydCI7czo5OiJsYXN0X25hbWUiO3M6Njoib2Zmc2V0IjtpOjQ4NzU7czo1OiJhbHBoYSI7czowOiIiO30&offset=4900')
driver.get('http://www.icsc.org/2015EDM/whos-coming/attendees/?paginate=members&filter=YTo3OntzOjU6ImxpbWl0IjtzOjI6IjI1IjtzOjEwOiJtYXhyZXN1bHRzIjtzOjI6IjI1IjtzOjE0OiJzaG93LWNvbXBhbmllcyI7czoyOiJubyI7czo4OiJvcmRlcl9ieSI7czo5OiJsYXN0X25hbWUiO3M6NDoic29ydCI7czo5OiJsYXN0X25hbWUiO3M6Njoib2Zmc2V0IjtpOjE0MjU7czo1OiJhbHBoYSI7czowOiIiO30&offset=2175')
userLink = driver.find_elements_by_css_selector('p.attendee-title a')
emailCount = 0
emailSkip = 9
emailDone = False
while(True):
    personCount = 0
    userLink = driver.find_elements_by_css_selector('p.attendee-title a')
    for link in userLink:
        dict = {}
        newLinksElement = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "p.attendee-title a")))
        newLinks = driver.find_elements_by_css_selector('p.attendee-title a')
        driver.get(newLinks[personCount].get_attribute('href'))
        personCount += 1
        if (emailSkip > 0):
            emailSkip = emailSkip - 1
            driver.back()
            continue
        try:
            nameElement = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "html.js.no-touch body div#content-banner.profile div.wrapper.cleared div.main div.text.cleared h2")))
            dict['name'] = nameElement.text
        except TimeoutException:
            dict['name'] = ''
        # dict['name'] = driver.find_element_by_css_selector('html.js.no-touch body div#content-banner.profile div.wrapper.cleared div.main div.text.cleared h2').text
        try:
            genPosTest = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "html.js.no-touch body div#content-banner.profile div.wrapper.cleared div.main div.text.cleared p")))
            genPosition = driver.find_element_by_css_selector('html.js.no-touch body div#content-banner.profile div.wrapper.cleared div.main div.text.cleared p')
        except TimeoutException:
            dict['company'] = ''
            dict['genPos'] = ''
            dict['country'] = ''
            dict['state'] = ''
            dict['address'] = ''
            dict['phone'] = ''
            dict['fax'] = ''
            dict['compUrl'] = ''
            driver.back()
            people.append(dict)
            continue
        genPosition = genPosition.text
        genPosition = genPosition.split('\n')
        
        company = driver.find_element_by_css_selector('html.js.no-touch body div#content-banner.profile div.wrapper.cleared div.main div.text.cleared p a').text
        companyIdx = -1
        for idx, val in enumerate(genPosition):
            if (company in val):
                comapnyIdx = idx
                dict['company'] = val
                if (idx > 0):
                    dict['genPos'] = genPosition[0]
                else:
                    dict['genPos'] = ''
        dict['country'] = genPosition[len(genPosition)-1]
        dict['state'] = genPosition[len(genPosition)-2]
        address = ''
        for x in range(comapnyIdx+1,len(genPosition)-2):
            address += genPosition[x] + ' '
        dict['address'] = address
        try:
            contact = driver.find_element_by_css_selector('html.js.no-touch body div#content-banner.profile div.wrapper.cleared div.sidebar p:nth-child(2)').text
        except NoSuchElementException:
            driver.back()
            people.append(dict)
            continue
        contact = contact.split('\n')
        for val in contact:
            if ('Phone') in val:
                dict['phone'] = val[7:]
            elif ('Fax') in val:
                dict['fax'] = val[5:]
            elif ('Company URL') in val:
                dict['compUrl'] = val[13:]
        try:
            email = driver.find_element_by_css_selector('a.display-email')
            email.click()
            time.sleep(10)
            email = email.text
            print email
            dict['email'] = email
            emailCount +=1
        except NoSuchElementException:
            driver.back()
            people.append(dict)
            continue
        if (emailCount == 103):
            emailDone = True
            break
        people.append(dict)
        driver.back()
    if (emailDone):
        break
    try:
        waitElement = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "groups-items")))
        nextPage = driver.find_element_by_css_selector('span.next a')
        driver.get(nextPage.get_attribute('href'))
        print 'next page'
    except (TimeoutException, NoSuchElementException) as e:
        print "Loading took too much time!"
        break

# Printing out the results

query = 'egan3'
fileName=query+'.xlsx'
workbook = xlsxwriter.Workbook(query+'.xlsx')
worksheet = workbook.add_worksheet()
row=1
col=0
worksheet.write(row,0, "Name")
worksheet.write(row,3,"Position")
worksheet.write(row,2,"Company")
worksheet.write(row,4,"Company URL")
##worksheet.write(row,4,"Address")
##worksheet.write(row,5,"State")
##worksheet.write(row,6,"Country")
worksheet.write(row,8,"Phone")
##worksheet.write(row,8,"Fax")
worksheet.write(row,6,"Email")
for item in people:
    row+=1
    if (item.has_key("name")):
        col=0
        worksheet.write(row,col,item["name"].strip())
    if (item.has_key("genPos")):
        col=3
        worksheet.write(row,col,item["genPos"].strip())
    if (item.has_key("company")):
        col=2
        worksheet.write(row,col,item["company"].strip())
    if (item.has_key("compUrl")):
        col=4
        worksheet.write(row,col,item["compUrl"].strip())
##    if (item.has_key("address")):
##        col=4
##        worksheet.write(row,col,item["address"].strip())
##    if (item.has_key("state")):
##        col=5
##        worksheet.write(row,col,item["state"].strip())
##    if (item.has_key("country")):
##        col=6
##        worksheet.write(row,col,item["country"].strip())
    if (item.has_key("phone")):
        col=8
        worksheet.write(row,col,item["phone"].strip())
##    if (item.has_key("fax")):
##        col=8
##        worksheet.write(row,col,item["fax"].strip())
    if (item.has_key("email")):
        col=6
        worksheet.write(row,col,item["email"].strip())
workbook.close()

print ('done')
