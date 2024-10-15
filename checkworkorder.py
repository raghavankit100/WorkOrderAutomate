from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

CaseWorkpackage=[
'BINECVS',
'BINECWX',
'BIOASLW',
'BINECXC',
]


driver = webdriver.Chrome()
driver.get('https://repository.xchanging.com/web/')
driver.maximize_window()

time.sleep(20)

Account_input = driver.find_element(By.ID , "nativeRepositoryPluginDojo_IMRLayout_0_LoginPane_XCusername")
Account_input.send_keys("LPCI")

username_input = driver.find_element("name","nativeRepositoryPluginDojo_IMRLayout_0_LoginPane_XCusername2")
username_input.send_keys("IND16Q")

time.sleep(2)

password_input = driver.find_element("name", "nativeRepositoryPluginDojo_IMRLayout_0_LoginPane_password")
password_input.send_keys("Partnership@452")
password_input.send_keys(Keys.ENTER)

time.sleep(15)

workPackage_click = driver.find_element("id", "dijit__TreeNode_3_label")
workPackage_click.click()
time.sleep(2)

NotShown_Workpackage=[]

for value in CaseWorkpackage:

    Search_Value= driver.find_element("id", "iMRSearchTemplatePluginDojo_SearchForm_1_ecm.widget.SearchCriterian_0")
    Search_Value.clear()
    Search_Value.send_keys(value)
    time.sleep(2)

    Search_click = driver.find_element("id", "dijit_form_Button_27_label")
    Search_click.click()
    time.sleep(12)


    verify_Contact= driver.find_element("id", "contractDetail_titleBarNode")
    if(verify_Contact.text=='Contract Details'):
        pass
    else:
        NotShown_Workpackage.append(value)
        
    time.sleep(5)
    search_icon=driver.find_element(By.CSS_SELECTOR, ".dijitReset.dijitInline.dijitIcon.iconNode.IMRSearchTemplatePluginLaunchIcon")
    search_icon.click()
    time.sleep(5)

    Search_criteria=driver.find_element(By.XPATH,'//*[@id="dijit_layout_ContentPane_33"]/div[1]/div/table/tbody/tr/td[1]/div')
    Search_criteria.click()
    time.sleep(2)



PushWorkpackages = list(set(CaseWorkpackage) - set(NotShown_Workpackage))

