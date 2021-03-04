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
import readData as read
import tkinter.filedialog as filedialog
import tkinter as tk



class Auto:
    def __init__(self, url):
        self.url = url
        self.driver = webdriver.Chrome('chromedriver.exe')
        self.driver.maximize_window() #maximize window
        self.driver.implicitly_wait(5) #waiting to load
        self.driver.get(self.url) #link

    def FillTextByID(self, id, content):
       self.driver.find_element(By.ID, id).send_keys(content)

    def FillTextByXpath(self, xpath, content):
       self.driver.find_element(By.XPATH, xpath).send_keys(content)
    
    def FillInputClickByXpath(self, xpathParent, xpathChild):
       self.driver.find_element(By.XPATH, xpathParent).click()
       time.sleep(1)
       self.driver.find_element(By.XPATH, xpathChild).click()
       time.sleep(1)


    def ClickButtonByID(self, id):
        self.driver.find_element(By.ID, id).click()
    
    def ClickButtonByXpath(self, xpath):
        self.driver.find_element(By.XPATH, xpath).click()


    def DeleteAndFillText(self, xpath, date):
        self.driver.find_element(By.XPATH, xpath).clear()
        self.FillTextByXpath(xpath, date)

    def ClickButttonInNewDocument(self, xpathFirst, xpathSecond, xpathThird):

        self.ClickButtonByXpath(xpathFirst)
        self.driver.switch_to.frame(0)
        self.ClickButtonByXpath(xpathSecond)
        self.ClickButtonByXpath(xpathThird)
        self.driver.switch_to.default_content()


    def deleteAndFillUnit(self, xpathText, content):
        self.driver.find_element(By.XPATH, xpathText).clear()

        self.FillTextByXpath(xpathText, content)
        

    def clearText(self,xpath):
        self.driver.find_element(By.XPATH, xpath).clear()


class Form(Auto):
    def __init__(self, path):
        self.url = "https://dichvucong.moit.gov.vn/Login.aspx" 
        Auto.__init__(self, self.url)
        self.path = path

    def Login(self):
        #Login
        self.FillTextByID("ctl00_cplhContainer_txtLoginName", "0316543468")
        self.FillTextByID("ctl00_cplhContainer_txtPassword", "Vtashipping")
        self.ClickButtonByID("ctl00_cplhContainer_btnLogin")

        #Move to C/O form
        self.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_grdViewDefault']/tbody/tr[2]/td[4]/a")
        self.ClickButtonByID("timer")
        self.ClickButtonByXpath("//*[@id='ctl00_Menu1_radMenu']/ul/li[1]/ul/li[1]/div/a")


    def COForm(self, numOfTK, data): 
        
        # Form E
        if data[int(2*int(data[0])+1)] == 'Form E':
            self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbFormCO_Input']", '//*[@id="ctl00_cplhContainer_cmbFormCO_DropDown"]/div/ul/li[7]')
            time.sleep(1)
            if data[int(2*int(data[0])+2)] == 'China':
                self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbMarket_Input']", "//*[@id='ctl00_cplhContainer_cmbMarket_DropDown']/div/ul/li[3]")
            elif data[int(2*int(data[0])+2)] == 'Myanmar':
                self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbMarket_Input']", '//*[@id="ctl00_cplhContainer_cmbMarket_DropDown"]/div/ul/li[7]')
            elif data[int(2*int(data[0])+2)] == 'Thailand':
                self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbMarket_Input']", '//*[@id="ctl00_cplhContainer_cmbMarket_DropDown"]/div/ul/li[10]')
        elif data[int(2*int(data[0])+1)] == 'Form EUR.1':
            self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbFormCO_Input']", '//*[@id="ctl00_cplhContainer_cmbFormCO_DropDown"]/div/ul/li[17]')
            time.sleep(1)
            if data[int(2*int(data[0])+2)] == 'Germany':
                self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbMarket_Input']", '//*[@id="ctl00_cplhContainer_cmbMarket_DropDown"]/div/ul/li[14]')
        

        for i in range(0, int(int(data[0])-1)):
            self.ClickButtonByXpath('//*[@id="ctl00_cplhContainer_btnAddRowCustomsNumber"]')
            time.sleep(1)

        # Exporter's business name
        self.DeleteAndFillText("//*[@id='ctl00_cplhContainer_PersionNameExportEnglish']",data[2*data[0]+3])

        # Address line 1 (exporter)
        self.DeleteAndFillText("//*[@id='ctl00_cplhContainer_AddressEnglishExport']",data[2*data[0]+4])

        # Address line 2 (exporter)
        self.DeleteAndFillText("//*[@id='ctl00_cplhContainer_AddressEnglishExport2']",data[2*data[0]+5])

        # Transportation type
        self.FillInputClickByXpath("//*[@id='ctl00_cplhContainer_cmbTransportMethod_Input']", "//*[@id='ctl00_cplhContainer_cmbTransportMethod_DropDown']/div/ul/li[2]")
        # Consignee’s name 
        self.FillTextByID("ctl00_cplhContainer_PersionNameImportEnglish", data[2*data[0]+6])

        # Address line 1 
        self.FillTextByID("ctl00_cplhContainer_AddressEnglishImport", data[2*data[0]+7])

        # Address line 2
        self.FillTextByID("ctl00_cplhContainer_AddressEnglishImport2", data[2*data[0]+8])
        
        # Port of Loading  
        self.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbSenderPlace_Input']", "CANG CAT LAI")
        time.sleep(2)
        self.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_cmbSenderPlace_DropDown']/div[1]/ul/li")
        time.sleep(1)

        # Port of Discharge//
        self.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbReceiverPlace_Input']", data[2*data[0]+9])
        time.sleep(2)
        self.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_cmbReceiverPlace_DropDown']/div[1]/ul/li")
        time.sleep(1)
        
        # Vessel’s Name/Aircraft etc 
        self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtShipName']", data[2*data[0]+10])

        # Departure date 
        self.DeleteAndFillText("//*[@id='ctl00_cplhContainer_txtTransportDate_dateInput']", data[2*data[0]+11])
    
        # Click Add/Update Item
        self.ClickButtonByID("ctl00_cplhContainer_radbtnSelectTabGoods_input")

    def GoodsForm(self, loopTime, data):
        #QUnit
        self.deleteAndFillUnit("//*[@id='ctl00_cplhContainer_cmbUnit_Input']", 'KILOGRAM')

        # GUnit
        self.deleteAndFillUnit("//*[@id='ctl00_cplhContainer_cmbGwUnitId_Input']", "KILOGRAM")

        # UPackage Quantity
        self.deleteAndFillUnit("//*[@id='ctl00_cplhContainer_cmbBoxUnitId_Input']", "CARTON")

        # Exporting HS code/ 
        self.FillTextByXpath("//*[@id='ctl00_cplhContainer_cmbHSCode_Input']", data[2*data[0]+13])
        time.sleep(2)
        self.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_cmbHSCode_DropDown']/div[1]/ul/li")


        time.sleep(2)

        # Invoce Number
        self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtInvoiceItem']", int(data[2*data[0]+14]))

        # Date
        self.FillTextByXpath("//*[@id='ctl00_cplhContainer_radDpkInvoiceItemDate_dateInput']", data[2*data[0]+15])

        # Origin criterion
        self.ClickButttonInNewDocument("//*[@id='ctl00_cplhContainer_rpvGoods']/div[1]/div[2]/div[2]/div/img", "//*[@id='chkWO']", "//*[@id='ctl00_cplhContainer_radToolBarDefault']/div/div/div/ul/li[1]/a/span/span/span/span")
        time.sleep(1)
    
        for i in range(0 , int(loopTime)):
            if i > 0:
                time.sleep(1)
            # Goods description
            self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtName']", str(data[data[2*data[0]+16]+(6*i)]))
             # Shipping mark
            self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtShippingMark']", data[data[2*data[0]+17]+(6*i)])
            # Quantity
            self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtUnitValue']", int(data[data[2*data[0]+18]+(6*i)]))
            # Gross weight
            self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtGwValue']", int(data[data[2*data[0]+19]+(6*i)]))

            # Package Quantity
            self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtBoxValue']", int(data[data[2*data[0]+20]+(6*i)]))

            # FOB value
            self.FillTextByXpath("//*[@id='ctl00_cplhContainer_txtCurrencyValue']", int(data[data[2*data[0]+21]+(6*i)]))

            #save
            self.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_btnAddItem']")
            time.sleep(2)
            
            self.clearText('//*[@id="ctl00_cplhContainer_txtShippingMark"]')
           
        
        time.sleep(1)    
        self.ClickButtonByXpath("//*[@id='ctl00_cplhContainer_ckbShowOnCO']")



    def run(self):
        self.Login()
        data = read.DataFromExcel(self.path)
        numofTK = data[0]
        self.COForm(numofTK, data)
        loopTime = data[13]
        self.GoodsForm(loopTime, data)


def save_info(event):
    path = link_entry.get()
    form = Form(path)
    form.run()

def input():
    link_path = tk.filedialog.askopenfilename()
    link_entry.delete(1, tk.END)  # Remove current text in entry
    link_entry.insert(0, link_path)  # Insert the 'path'

if __name__ == "__main__":
    master = tk.Tk()
    top_frame = tk.Frame(master)
    bottom_frame = tk.Frame(master)
    line = tk.Frame(master, height = 1, width = 400, bg = "grey80", relief = "groove", )


    link_text = tk.Label(top_frame, text = "Input File Path:")
    link_entry = tk.Entry(top_frame, text = "", width = 40)
    browse = tk.Button(top_frame, text = "Browse", command = input)

    begin_button = tk.Button(bottom_frame, text = "Begin!", command = save_info)

    master.bind("<Return>",save_info)

    top_frame.pack(side=tk.TOP)
    line.pack(pady=10)
    bottom_frame.pack(side=tk.BOTTOM)

    link_text.pack(pady=5)
    link_entry.pack(pady=5)
    browse.pack(pady=5)

    begin_button.pack(pady=20, fill=tk.X)

    master.mainloop()

