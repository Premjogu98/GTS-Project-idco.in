from selenium import webdriver
import time
import re
import sys , os
from datetime import datetime
import Global_var
import wx
from Insert_On_databse import insert_in_Local
import string
import ctypes
import html
app = wx.App()
def Choromedriver():
    browser = webdriver.Chrome(executable_path=str(f"C:\\chromedriver.exe"))
    browser.get('http://english.idco.in/idcodnnmodule/tendernew.aspx')
    browser.maximize_window()
    time.sleep(2)
    for table_id in ['gvTender','gvQuotation','gvEOI']:
        next_page = False
        next_page_count = 2
        td = 2
        while next_page == False:
            publish_date = browser.find_elements_by_xpath(f'//*[@id="{table_id}"]/tbody/tr/td[2]')
            for i in range(len(publish_date)-1):
                for publish_date in browser.find_elements_by_xpath(f'//*[@id="{table_id}"]/tbody/tr[{str(td)}]/td[2]'):
                    browser.execute_script("arguments[0].scrollIntoView();", publish_date)
                    publish_date = publish_date.get_attribute('innerText').strip()
                    break
                datetime_object_pub = datetime.strptime(publish_date, '%d-%b-%Y')
                User_Selected_date = datetime.strptime(str(Global_var.From_Date), '%Y-%m-%d')
                timedelta_obj = datetime_object_pub - User_Selected_date
                day = timedelta_obj.days
                if day >= 0:
                    for Closing_date in browser.find_elements_by_xpath(f'//*[@id="{table_id}"]/tbody/tr[{str(td)}]/td[3]'):
                        Closing_date = Closing_date.get_attribute('innerText').strip()
                        break
                    for Opening_date in browser.find_elements_by_xpath(f'//*[@id="{table_id}"]/tbody/tr[{str(td)}]/td[4]'):
                        Opening_date = Opening_date.get_attribute('innerText').strip()
                        break
                    for Department in browser.find_elements_by_xpath(f'//*[@id="{table_id}"]/tbody/tr[{str(td)}]/td[6]'):
                        Department = Department.get_attribute('innerText').strip()
                        break
                    for Title in browser.find_elements_by_xpath(f'//*[@id="{table_id}"]/tbody/tr[{str(td)}]/td[5]/a'):
                        Title_text = Title.get_attribute('innerText').strip()
                        Title.click()
                        time.sleep(4)
                        break
                    get_htmlsource = ''
                    for get_htmlsource in browser.find_elements_by_xpath('//*[@class="modal-body"]'):
                        get_htmlsource = get_htmlsource.get_attribute('outerHTML').strip()
                        break
                    pdf_document_text = ''
                    if len(browser.window_handles) != 1:
                        browser.switch_to.window(browser.window_handles[0])
                        browser.close()
                        browser.switch_to.window(browser.window_handles[0])
                    visible = True
                    while visible == True:
                        try:
                            for pdf_document in browser.find_elements_by_xpath('//*[@onclick="SetTarget();"]'):
                                pdf_document_text = pdf_document.get_attribute('outerHTML').replace('\n','').strip()
                                pdf_document_text = pdf_document_text.partition('href="')[2].partition('"')[0]
                                pdf_document.click()
                                time.sleep(2)
                                visible = False
                        except:
                            visible = True
                            print('Error On PDF DOWNLOAD Links')
                            time.sleep(1)
                    browser.switch_to.window(browser.window_handles[1])
                    pdf_url = browser.current_url
                    get_htmlsource = get_htmlsource.replace(pdf_document_text,pdf_url).replace('href="../','href="http://english.idco.in/').replace('\n','')
                    browser.close()
                    browser.switch_to.window(browser.window_handles[0])
                    Scraping_data(get_htmlsource,Title_text,Department,Opening_date,Closing_date,publish_date)
                    Global_var.Total += 1
                    td += 1
                    print(f'Total: {Global_var.Total} Deadline Not given: {Global_var.deadline_Not_given} duplicate: {Global_var.duplicate} inserted: {Global_var.inserted} expired: {Global_var.expired} QC Tenders: {Global_var.QC_Tenders}')
                    for Close_model in browser.find_elements_by_xpath('//*[@data-dismiss="modal"]'):
                        Close_model.click()
                        break
                    time.sleep(2)
                else:
                    next_page = True
            if next_page == False:
                if table_id == 'gvTender':
                    for nextpage in browser.find_elements_by_xpath(f'//*[@class="pagination-ys"]/td/table/tbody/tr/td[{str(next_page_count)}]/a'):
                        browser.execute_script("arguments[0].scrollIntoView();", nextpage)
                        nextpage.click()
                        time.sleep(2)
                        next_page_count += 1
                        td = 2
                        break
                elif table_id == 'gvEOI':
                    for nextpage in browser.find_elements_by_xpath(f'//*[@id="gvEOI"]/tbody/tr[12]/td/table/tbody/tr/td[2]/a'):
                        browser.execute_script("arguments[0].scrollIntoView();", nextpage)
                        nextpage.click()
                        time.sleep(2)
                        td = 2
                        break
    wx.MessageBox(f'Total: {Global_var.Total}\nDeadline Not given: {Global_var.deadline_Not_given}\nduplicate: {Global_var.duplicate}\ninserted: {Global_var.inserted}\nexpired: {Global_var.expired}\nQC Tenders: {Global_var.QC_Tenders}','english.idco.in', wx.OK | wx.ICON_INFORMATION)
    browser.close()
    sys.exit()
def Scraping_data(get_htmlsource,Title_text,Department,Opening_date,Closing_date,publish_date):
    SegField = []
    for data in range(45):
        SegField.append('')

    SegField[1] = ''
    SegField[2] = ''
    if Opening_date != "":
        Opening_date = datetime.strptime(str(Opening_date).strip(), "%d-%b-%Y")
        Opening_date = Opening_date.strftime("%Y-%m-%d")
        # SegField[3] = orgaddress+ "~" + orgemail + "~" + orgtelfax + "~" + docstartdata + "~" + opendata;
        SegField[3] = "NA" + "~" + "NA" + "~" + "NA" + "~" + "NA" + "~" + Opening_date.strip()
    else:
        SegField[3] = "NA" + "~" + "NA" + "~" + "NA" + "~" + "NA" + "~" + "NA"

    SegField[12] = 'ODISHA INDUSTRIAL INFRASTRUCTURE DEVELOPMENT CORPORATION'

    tender_id = Title_text.partition('No-')[2].partition('Date-')[0].strip()
    SegField[13] = tender_id

    SegField[14] = "2"  # notice_type

    Title_text = string.capwords(str(Title_text))
    SegField[19] = Title_text.strip().replace('"','').replace('“','')
    Department = string.capwords(str(Department))

    SegField[18] = f'{SegField[19]}<br>\nDepartment: {Department.strip()}<br>\nOpening Date: {Opening_date}<br>\nClosing Date: {Closing_date}<br>\nPublish Date: {publish_date}'
    
    if Closing_date != "":
        deadline = datetime.strptime(str(Closing_date).strip(), "%d-%b-%Y")
        deadline = deadline.strftime("%Y-%m-%d")
        SegField[24] = deadline.strip()

    SegField[7] = "IN"
    SegField[20] = ""
    SegField[22] = ""
    SegField[26] = "0.0"  # earnest_money
    SegField[27] = "0"  # Financier

    SegField[28] = 'http://english.idco.in/idcodnnmodule/tendernew.aspx'

    SegField[31] = "idco.in"
    SegField[42] = SegField[7]
    SegField[43] = ""

    for SegIndex in range(len(SegField)):
        print(SegIndex, end=' ')
        print(SegField[SegIndex])
        SegField[SegIndex] = html.unescape(str(SegField[SegIndex]))
        SegField[SegIndex] = str(SegField[SegIndex]).replace("'", "''")

    if len(SegField[19]) >= 200:
        SegField[19] = str(SegField[19])[:200]+'...'

    if len(SegField[18]) >= 1500:
        SegField[18] = str(SegField[18])[:1500]+'...'

    if SegField[19] == '':
        wx.MessageBox(' Short Desc Blank ','ebusiness.kockw.com', wx.OK | wx.ICON_ERROR)
    else:
        # check_date(get_htmlsource, SegField)
        pass
        
def check_date(get_htmlSource, SegField):
    deadline = str(SegField[24])
    curdate = datetime.now()
    curdate_str = curdate.strftime("%Y-%m-%d")
    try:
        if deadline != '':
            datetime_object_deadline = datetime.strptime(deadline, '%Y-%m-%d')
            datetime_object_curdate = datetime.strptime(curdate_str, '%Y-%m-%d')
            timedelta_obj = datetime_object_deadline - datetime_object_curdate
            day = timedelta_obj.days
            if day > 0:
                insert_in_Local(get_htmlSource, SegField)
                # create_filename(get_htmlSource , SegField)
            else:
                print("Expired Tender")
                Global_var.expired += 1
        else:
            print("Deadline Not Given")
            Global_var.deadline_Not_given += 1
            wx.MessageBox(' Deadline Not Given ','ipms.ppadb.co.bw', wx.OK | wx.ICON_INFORMATION)
            # insert_in_Local(get_htmlSource, SegField)
    except Exception as e:
        exc_type , exc_obj , exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname , "\n" ,exc_tb.tb_lineno)

Choromedriver()