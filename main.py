from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import xlwt
PATH="D:/python_practice/find_client_data/chromedriver.exe"
driver=webdriver.Chrome(PATH)
driver.get('https://www.yes123.com.tw/admin/index.asp')
driver.find_element_by_xpath('//*[@id="find_key1"]').send_keys('電子工程師')
driver.find_element_by_xpath('/html/body/div[2]/form/div/div/a/div').click()
company_name_list=['','']
workbook=xlwt.Workbook(encoding='ascii')
worksheet=workbook.add_sheet("工作表一")
worksheet.write(0,0,"公司名稱")
worksheet.write(0,1,"公司類別")
worksheet.write(0,2,"公司說明")
worksheet.write(0,3,"員工數量")
worksheet.write(0,4,"資本額")
worksheet.write(0,5,"公司電話")
worksheet.write(0,6,"公司地址")
worksheet.write(0,7,"公司網站")
listcount=0
for page in range(1,100):
    soup=BeautifulSoup(driver.page_source,"html5lib")
    allPage=soup.find('div','main_content')
    company_divs=allPage.find_all('div','Job_opening_item_title')
    company_url_list=[]
    for company_id_a_href in company_divs:
        company_id=str(company_id_a_href.h5.a)
        company_p_id=company_id[17:45]
        url="https://www.yes123.com.tw/wk_index/comp_info.asp?"+company_p_id
        company_url_list.append(url)
    print(company_url_list)
    for url in company_url_list:
        collect=False
        driver.get(url=url)
        company_page_information_html=BeautifulSoup(driver.page_source,"html5lib")
        company_name=str(company_page_information_html.find('title').text.replace('|【工作職缺與徵才簡介】 yes123 求職網',''))\
                                                                         .replace('【工作職缺與徵才簡介】 yes123 求職網','')
        for i in range (0,len(company_name_list)):
            if(company_name==company_name_list[i]):
                collect=False
                break
            else:
                collect=True
        if collect==True:
            listcount+=1
            company_name_list.append(company_name)
            try:
                company_page_html=company_page_information_html.find('div','job_explain s_mt20 mt d-flex justify-content-between')\
                                                               .find('ul').find_all('li')
                company_class='無資料'
                company_instruction='無資料'
                company_employee_number='無資料'
                company_capital='無資料'
                company_telephone='無資料'
                company_address='無資料'
                company_website='無資料'
                for i in range (0,len(company_page_html)):
                    try:
                        title=str(company_page_html[i].find('span','left_title').text)
                        if(title=='行業類別：'):company_class=company_page_html[i].find('span','right_main').text
                        elif(title=='行業說明：'):company_instruction=company_page_html[i].find('span','right_main').text
                        elif(title=='員工人數：'):company_employee_number=company_page_html[i].find('span','right_main').text
                        elif(title=='資本金額：'):company_capital=company_page_html[i].find('span','right_main').text
                        elif(title=='企業電話：'):company_telephone=company_page_html[i].find('span','right_main').text
                        elif(title=='企業地址：'):company_address=company_page_html[i].find('span','right_main').a.text
                        elif(title=='企業網址：'):company_website=company_page_html[i].find('span','right_main').a.text
                        else:continue
                    except AttributeError:
                        continue
                print(company_name)
                print(company_class)
                print(company_instruction)
                print(company_employee_number)
                print(company_capital)
                print(company_telephone)
                print(company_address)
                print(company_website)
                print('----------------------------')
                worksheet.write(listcount, 0, company_name)
                worksheet.write(listcount, 1, company_class)
                worksheet.write(listcount, 2, company_instruction)
                worksheet.write(listcount, 3, company_employee_number)
                worksheet.write(listcount, 4, company_capital)
                worksheet.write(listcount, 5, company_telephone)
                worksheet.write(listcount, 6, company_address)
                worksheet.write(listcount, 7, company_website)
            except AttributeError:
                continue
    driver.get('https://www.yes123.com.tw/admin/index.asp')
    driver.find_element_by_xpath('//*[@id="find_key1"]').send_keys('電子工程師')
    driver.find_element_by_xpath('/html/body/div[2]/form/div/div/a/div').click()
    select=Select(driver.find_element_by_xpath('//*[@id="inputState"]'))
    select.select_by_index(page)
workbook.save('客戶資料.xls')