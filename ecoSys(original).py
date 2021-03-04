from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.support import expected_conditions as EC
import time
import xlrd

class AutoFill:
    def __init__(self, url):
        self.url = url
        self.driver = webdriver.Chrome('chromedriver.exe')
        self.driver.maximize_window() #maximize window
        self.driver.implicitly_wait(10) #waiting to load
        self.driver.get(self.url) #link

    def FillTextByID(self, id, content):
       self.driver.find_element(By.ID, id).send_keys(content)

    def FillTextByXpath(self, xpath, content):
       self.driver.find_element(By.XPATH, xpath).send_keys(content)
    
    def FillInputClickByXpath(self, xpathParent, xpathChild):
       self.driver.find_element(By.XPATH, xpathParent).click()
       time.sleep(2)
       self.driver.find_element(By.XPATH, xpathChild).click()
       time.sleep(2)


    def ClickButtonByID(self, id):
        self.driver.find_element(By.ID, id).click()
    
    def ClickButtonByXpath(self, xpath):
        self.driver.find_element(By.XPATH, xpath).click()

    def SelectScrollBarByID(self, id, content):
        item = Select(self.driver.find_element(ByID, id))
        item.select_by_visible_text(content)
        
    def SelectScrollBarByXpath(self, xpath, content):
        item = Select(self.driver.find_element(By.XPATH, xpath))
        item.select_by_visible_text(content)

    def DeleteAndFillText(self, xpath, date):
        for i in range(0, 10):
            self.driver.find_element(By.XPATH, xpath).send_keys(Keys.BACKSPACE)

        self.FillTextByXpath(xpath, date)
    
    def FindButtonCss(self):
         link = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "rbDecorated")))
         link.click()
        
    def test(self, driver, xpath):
        driver.find_element(By.XPATH, xpath).click()

    def ClickButttonInNewDocument(self, xpathFirst, xpathSecond, xpathThird):
        save = self.driver
        self.ClickButtonByXpath(xpathFirst)
        self.driver.switch_to.frame(0)
        self.ClickButtonByXpath(xpathSecond)
        self.ClickButtonByXpath(xpathThird)
        self.driver.switch_to.default_content()

def readDataFromExcel():
    path = 'readData.xlsx'
    inputWorkbook = xlrd.open_workbook(path)
    inputWorksheet = inputWorkbook.sheet_by_index(1)
    row = inputWorksheet.nrows
    

    ecoSys = []

    for i in range(0,row):
        ecoSys.append(inputWorksheet.cell_value(i, 0))
    return ecoSys

if __name__ == "__main__":
    ecoSys = readDataFromExcel()
    auto = AutoFill("https://dichvucong.moit.gov.vn/Login.aspx")

    # Đăng Nhập
    auto.FillTextByID("ctl00_cplhContainer_txtLoginName", "0316543468")
    auto.FillTextByID("ctl00_cplhContainer_txtPassword", "Vtashipping")
    auto.ClickButtonByID("ctl00_cplhContainer_btnLogin")

   
    # click khai báo
    auto.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_grdViewDefault']/tbody/tr[2]/td[4]/a")

    # Tắt thông báo
    auto.ClickButtonByID("timer")

    # Click khai báo
    auto.ClickButtonByXpath("//*[@id='ctl00_Menu1_radMenu']/ul/li[1]/ul/li[1]/div/a")


    # Bắt đầu điền vào bản khai báo

    # Form
    auto.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbFormCO_Input']", "//*[@id='ctl00_cplhContainer_cmbFormCO_DropDown']/div/ul/li[7]")

    # Importing country
    auto.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbMarket_Input']", "//*[@id='ctl00_cplhContainer_cmbMarket_DropDown']/div/ul/li[3]")


    # Export Declaration Number
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_plhCustomsNumber0_txtInvoiceNumber']", "035433333333")

    # ngày
    auto.DeleteAndFillText("//*[@id='ctl00_cplhContainer_plhCustomsNumber0_radDpkInvoiceDate_dateInput']", "24/12/2020")

    # Consignee’s name 
    auto.FillTextByID("ctl00_cplhContainer_PersionNameImportEnglish", "MUSA JUTE FIBERS")
    
    #Address line 1 
    auto.FillTextByID("ctl00_cplhContainer_AddressEnglishImport", "24, F.ARAJI PARA ROAD, MOYLAPOTA MORE BESIDE KFC")

    #Address line 2
    auto.FillTextByID("ctl00_cplhContainer_AddressEnglishImport2", "KHULNA-9100, BANGLADESH")

    #Transportation type
    auto.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbTransportMethod_Input']", "//*[@id='ctl00_cplhContainer_cmbTransportMethod_DropDown']/div/ul/li[2]")

    #Port of Loading  
    auto.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbSenderPlace_Input']", "//*[@id='ctl00_cplhContainer_cmbSenderPlace_DropDown']/div[1]/ul/li[7]")

    #Port of Discharge
    auto.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbReceiverPlace_Input']", "//*[@id='ctl00_cplhContainer_cmbReceiverPlace_DropDown']/div[1]/ul/li[9]")

    #Vessel’s Name/Aircraft etc
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtShipName']", "5")

    #Departure date 
    auto.DeleteAndFillText("//*[@id='ctl00_cplhContainer_txtTransportDate_dateInput']", "24/12/2020")
 
    # Click Add/Update Item
    auto.ClickButtonByID("ctl00_cplhContainer_radbtnSelectTabGoods_input")

    # Goods-----------------------------
    # Click Origin criterion
    auto.ClickButttonInNewDocument("//*[@id='ctl00_cplhContainer_rpvGoods']/div[1]/div[2]/div[2]/div/img", "//*[@id='chkWO']", "//*[@id='ctl00_cplhContainer_radToolBarDefault']/div/div/div/ul/li[1]/a/span/span/span/span")





    """
    # Exporting HS code
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbHSCode_Input']", "01012100")

    # Goods description
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtName']", "1")

    # Quantity
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtUnitValue']", '2')
    # QUnit
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbUnit_Input']", 'DOSE')

    # Gross weight
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtGwValue']", "2")
    # GUnit
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbGwUnitId_Input']", "DOSE")

    # Invoice number
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtInvoiceItem']", "1")

    # Data
    auto.DeleteAndFillText("//*[@id='ctl00_cplhContainer_radDpkInvoiceItemDate_dateInput']", "6/12/2000")

    # Importing HS code
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbHSCodeOutsite_Input']", "8516101000")


    # Mark and number on package
    #auto.FillTextInNewDocument("//*[@id='ctl00_cplhContainer_txtShippingMark']", "2")

    # Package Quantity
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtBoxValue']", "3")
    # UPackage Quantity
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtBoxValue']", "CASE")

    # FOB value
    auto.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtCurrencyValue']", "4")
    # Crc FOB value
    #auto.FillTextByXpath("//*[@id="ctl00_cplhContainer_cmbItemCurrencyUnit_Input"]", "")
"""


