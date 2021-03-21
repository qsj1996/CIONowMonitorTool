import traceback
from boto3 import Session
from xlrd.book import open_workbook_xls
import xlrd
import smtplib
import time
import tkinter
from pathlib import Path
from email.header import Header
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import PIL.Image as Image
import json
import os

#判断截图文件夹是否存在:不存在创建，存在判断文件夹是否为空，不为空清空
def setscreenshotfile(filepath):
    if not Path(filepath).exists():
        os.makedirs(filepath)
    return

def getscreen(screenpath):
    driver.get_screenshot_as_file(screenpath)
    return

def getRelativeValue(zoomValue):
    screen = tkinter.Tk()
    xi = int(screen.winfo_screenwidth())
    yi = int(screen.winfo_screenheight())
    relative = ((xi / 1920 + yi / 1080) / 2)*zoomValue
    relativeValue = '%.2f' % relative
    return relativeValue

def addpng(msgRoot, path, name):
    fp = open(path, 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', name)
    msgRoot.attach(msgImage)

def getConfigFromJson():
    with open("./config.json",'r') as load_f:
        configDict = json.load(load_f)
    return configDict

def chooseProjects(driver, selectedProjects):
    driver.find_element_by_xpath(r"//select[@id='project-filter']/following::div[1]/button").click()
    time.sleep(0.5)
    for i in selectedProjects:
        time.sleep(0.1)
        driver.find_element_by_xpath(r"//select[@id='project-filter']/following::div[1]/button/following::div[1]/div[1]/input").clear()
        driver.find_element_by_xpath(r"//select[@id='project-filter']/following::div[1]/button/following::div[1]/div[1]/input").send_keys(i)
        time.sleep(0.5)
        driver.find_element_by_xpath(r"//select[@id='project-filter']/following::div[1]//input[@data-name='selectAll']").click()
    time.sleep(0.1)

def monitorComplianceOps():
    #  ComplianceOps
    driver.get("https://cionow.accenture.com/ComplianceOps")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "maindash")))
    chooseProjects(driver,selectedProjects)
    driver.find_element_by_xpath(r"//span[@id='cp-filter-btn']").click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "maindash")))

    time.sleep(1)
    driver.execute_script("document.body.style.zoom='" + getRelativeValue(0.80) + "'")
    time.sleep(1)
    elAll = driver.find_element_by_id('maindash')
    driver.execute_script("arguments[0].scrollIntoView();", elAll)
    # elAll.screenshot(r'.\Screenshot\ComplianceOps.png')
    driver.save_screenshot(screenshotPath + r'\ComplianceOps.png')
    time.sleep(1)

    #  ComplianceOps 裁剪图片
    complianceImg = Image.open(screenshotPath + r'\ComplianceOps.png')
    cSize = complianceImg.size
    croppedComplianceImg = complianceImg.crop((cSize[0]/4, 0, 3*cSize[0]/4, cSize[1])) # (left, upper, right, lower)
    croppedComplianceImg.save(screenshotPath + r'\ComplianceOps_cropped.png')
    time.sleep(5)


def monitorDeliverService():
    #  DeliverService
    driver.get("https://cionow.accenture.com/DeliverService")
    # downloadDivXpath = "//div[@id='download-ToolbarButton']"
    WebDriverWait(driver, 60).until(lambda d: d.find_element_by_xpath("//iframe"))
    iframe = driver.find_element(By.XPATH, "//iframe")
    driver.switch_to.frame(iframe)

    WebDriverWait(driver, 20).until(lambda d: d.find_element_by_xpath("//*[@id='download-ToolbarButton']"))
    WebDriverWait(driver, 20).until(lambda d: d.find_element_by_xpath("//div[@id='tableau_base_widget_LegacyCategoricalQuickFilter_4']//div[@class='tabComboBoxButtonHolder']"))

    driver.find_element_by_xpath("//div[@id='tableau_base_widget_LegacyCategoricalQuickFilter_4']//div[@class='tabComboBoxButtonHolder']").click()

    driver.find_element_by_xpath("//a[@title='(All)']/preceding-sibling::input").click()

    for i in selectedProjects:
        time.sleep(0.1)
        driver.find_element_by_xpath(r"//textarea[@title='Search (Enter)' and @id]").clear()
        driver.find_element_by_xpath(r"//textarea[@title='Search (Enter)' and @id]").send_keys(i)
        time.sleep(1)
        selectedXpath = "//a[contains(@title,'"+i+"')]/preceding-sibling::input"
        driver.find_element_by_xpath(selectedXpath).click()

    driver.find_element_by_xpath("//span[text()='Apply']").click()
    time.sleep(2)
    WebDriverWait(driver, 20).until(lambda d: d.find_element_by_xpath("//div[@id='loadingGlassPane'and contains(@style,'display: none')]"))

    # locator = (By.XPATH, "//div[@id='tableau_base_widget_LegacyCategoricalQuickFilter_4']//div[@class='tabComboBoxButtonHolder']")
    # WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator))

    driver.find_element_by_xpath("//div[@class='tab-glass clear-glass tab-widget']").click()
    driver.find_element_by_xpath("//*[@id='download-ToolbarButton']").click()

    WebDriverWait(driver, 20).until(lambda d: d.find_element_by_xpath("//button[text()='Image']"))
    driver.find_element_by_xpath("//button[text()='Image']").click()

    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "loadingSpinner")))
    WebDriverWait(driver, 60).until_not(EC.visibility_of_element_located((By.ID, "loadingSpinner")))


    
def monitorPlanAndManageService():
    #  PlanAndManageService
    driver.get("https://cionow.accenture.com/PlanAndManageService")
    xpath = "//h3[contains(text(),'Forecast & Actuals')]"
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))

    
    chooseProjects(driver,selectedProjects)
    driver.find_element_by_xpath(r"//span[@onclick='planmanage_filter()']").click()

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))

    el1 = driver.find_element_by_id('planmanage-row-1')
    driver.execute_script("arguments[0].scrollIntoView();", el1)
    el1.screenshot(screenshotPath + r'\PlanAndManageService1.png')
    time.sleep(0.1)

    el2 = driver.find_element_by_id('planmanage-row-2')
    driver.execute_script("arguments[0].scrollIntoView();", el2)
    el2.screenshot(screenshotPath + r'\PlanAndManageService2.png')
    time.sleep(0.1)

    el3 = driver.find_element_by_id('planmanage-row-3')
    driver.execute_script("arguments[0].scrollIntoView();", el3)
    el3.screenshot(screenshotPath + r'\PlanAndManageService3.png')
    time.sleep(0.1)

    hrele = driver.find_element_by_xpath("//hr[@style='border-top:1px solid #aeaeae;']")
    driver.execute_script("arguments[0].scrollIntoView();", hrele)
    hrele.screenshot(screenshotPath + r'\hrEle.png')
    time.sleep(0.1)

    #  PlanAndManageService 拼接图片
    img1 = Image.open(screenshotPath + r'\PlanAndManageService1.png')
    img2 = Image.open(screenshotPath + r'\PlanAndManageService2.png')
    img3 = Image.open(screenshotPath + r'\PlanAndManageService3.png')
    hrImg = Image.open(screenshotPath + r'\hrEle.png')
    size1 = img1.size
    size2 = img2.size
    size3 = img3.size
    Sizehr = hrImg.size

    imgNew = Image.new('RGB', (size1[0], size1[1] + size2[1] + size3[1] +(2 * Sizehr[1]) ))

    loc1, lochr1, loc2, lochr2, loc3 = (0, 0),(0, size1[1]), (0, size1[1]+Sizehr[1]),(0, size1[1]+Sizehr[1]+ size2[1]), (0, size1[1] + size2[1]+2*Sizehr[1])

    imgNew.paste(img1,loc1)
    imgNew.paste(hrImg,lochr1)
    imgNew.paste(img2,loc2)
    imgNew.paste(hrImg,lochr2)
    imgNew.paste(img3,loc3)
    imgNew.save(screenshotPath + r'\PlanAndManageService.png')

def SendReportEmail():
    #Emil section
        receivers = [Emailto, Emailcc]
        msgRoot = MIMEMultipart('related')
        msgRoot['From'] = Header(EmailSender, 'utf-8')
        msgRoot['To'] = Header(Emailto, 'utf-8')
        msgRoot['Cc'] = Header(Emailcc, 'utf-8')
        subject = 'CIO Now Report ' + timenow
        msgRoot['Subject'] = Header(subject, 'utf-8')
        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)
        HTMLBody = '''
            <p>Hi All,</p>
            <p>Please get below screenshot and table of '''+projectsName+''' Team CIO Now report on '''+timenow+'''.</p>
            <h3>Compliance</h3>
            <p><img src="cid:Compliance"></p>
            <h3>PlanAndManageService</h3>
            <p><img src="cid:PlanAndManageService"></p>
            <h3>DeliverService</h3>
            <p><img src="cid:DeliverService"></p>
            <br/><br/>
            <u style="font-family:Segoe Script;style="font-size: 20px;color:#ef7c06">CIO CAMS Team</u>
            '''
        msgAlternative.attach(MIMEText(HTMLBody, 'html', 'utf-8'))
        # add png
        addpng(msgRoot, screenshotPath + r'\ComplianceOps_cropped.png', '<Compliance>')
        addpng(msgRoot, screenshotPath + r'\PlanAndManageService.png', '<PlanAndManageService>')
        addpng(msgRoot, screenshotPath + r'\Deliver Service.png', '<DeliverService>')
        smtpObj = smtplib.SMTP()
        smtpObj.connect('63.240.188.61', 25)
        smtpObj.sendmail(EmailSender, receivers, msgRoot.as_string())
        smtpObj.quit()

try:
    configs = getConfigFromJson()
    timenow = str(time.strftime("%Y-%m-%d", time.localtime()))
    datafile = configs["datafile"]

    EmailSender = configs["emailFrom"]
    failEmailToList = configs["failEmailToList"]
    failEmailCcList = configs["failEmailCcList"]

    monitorDetails = configs["monitorDetails"]

    for item in monitorDetails:
        Emailto = item["emailToList"]
        Emailcc = item["emailCcList"]
        projectsName = item["projectsName"]
        selectedProjects = item["selectedProjects"]
        #创建当日截图文件夹
        screenshotPath = ".\\" + projectsName +"\\Screenshot"+timenow
        fullPath = os.path.abspath(screenshotPath)
        setscreenshotfile(fullPath)

        # 配置webdriver
        option = webdriver.ChromeOptions()
        option.add_argument(r'--user-data-dir='+datafile+'') 
        option.add_argument('--profile-directory=Default')
        prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': fullPath}
        option.add_experimental_option('prefs', prefs)
        # option.add_argument('--download.default_directory='+screenshotPath)
        driver = webdriver.Chrome(options=option)
        driver.maximize_window()

        monitorPlanAndManageService()
        monitorDeliverService()
        monitorComplianceOps()

        #关闭浏览器
        driver.quit()
        
        # SendReportEmail()

except smtplib.SMTPException as e:
    receivers = [failEmailToList,failEmailCcList]
    message = MIMEText('邮件发送失败: '+(traceback.format_exc()), 'plain', 'utf-8')
    message['From'] = Header(EmailSender, 'utf-8')
    message['To'] = Header(failEmailToList, 'utf-8')
    message['Cc'] = Header(failEmailCcList, 'utf-8')
    subject = 'RE:[CAMS Daily Report]发送失败'
    message['Subject'] = Header(subject, 'utf-8')
    smtpObj = smtplib.SMTP()
    smtpObj.connect('63.240.188.61', 25)
    smtpObj.sendmail(EmailSender, receivers, message.as_string())