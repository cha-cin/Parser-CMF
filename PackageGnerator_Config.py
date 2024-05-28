from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import json
import pandas as pd
import pandas as pd
import json
import os
current_dir=os.path.dirname(os.path.abspath(__file__))
print(current_dir)
excel_file = current_dir+"\PackageData.xlsx"

df_Config = pd.read_excel(excel_file, sheet_name='Configuration', usecols="A:B")
print(df_Config)
Configuration = {}
for index, row in df_Config.iterrows():
    key = str(row['Item'])
    values = str(row['Value'])
    if values=="nan":
        Configuration[key] = ""
    else:
        Configuration[key] = values
print(Configuration)


df_ServiceResource = pd.read_excel(excel_file, sheet_name='ServiceResource', usecols="A:F")
ServiceResource = {}
for index, row in df_ServiceResource.iterrows():
    key = str(row['Service'])
    values = [str(row['Replicas']), str(row['MinMemory']), str(row['MaxMemory']),str(row['MinCPU']), str(row['MaxCPU'])]
    if "nan" in values:
        print("There is null values in ServiceResource")
        exit()
    ServiceResource[key] = values
    

df_Volumes = pd.read_excel(excel_file, sheet_name='Volumes', usecols="A:E")
Volumes = {}
for index, row in df_Volumes.iterrows():
    key = str(row['Item'])
    values = [str(row['Type']), str(row['Size']), str(row['Var1']),str(row['Var2'])]
    if "nan" in values[0:2]:
        print("There is null values in ServiceResource")
        exit()
    Volumes[key] = values
print(Volumes)
download_dir = str(Configuration["Destination Download Path"]) #Destination Download Path
print(download_dir)
chrome_options = webdriver.ChromeOptions()
prefs = {"download.default_directory": download_dir}
chrome_options.add_experimental_option("prefs", prefs)
url=Configuration["Installation URL"] #Need to change to corrrect Installation url
print(current_dir+"\chromedriver.exe")
#driver = webdriver.Chrome(current_dir+"\chromedriver.exe") 
driver = webdriver.Chrome(options=chrome_options)  


if "@micron.com" in Configuration["Your Micron Email"]:
    driver.get(url) #use chrome friver to open
    time.sleep(10)
    xpath = f"//input[@name='username']"
    UsernameItem = driver.find_element(By.XPATH,xpath)
    #WritePropertyeditorItem.click()
    UsernameItem.clear()
    UsernameItem.send_keys(Configuration["Your Micron Email"])
    button = driver.find_element(By.XPATH, "//button[@class='cmf-btn cmf-btn-primary']")
    time.sleep(2)
    button.click()
    time.sleep(50) #wait to login on your phone
    actions = ActionChains(driver) #ActionChains can real simulate the mouse to click
else:
    driver.get(url) #use chrome friver to open
    #time.sleep(15)
    with open(current_dir+"\Cookie.json") as f:
        cookies = json.load(f)
        for cookie in cookies:
            print(cookie)
            if 'sameSite' in cookie:
                cookie.pop('sameSite')  
            if 'domain' in cookie:
                cookie["domain"]=".criticalmanufacturing.com"  #cookie preprocessing
            driver.add_cookie(cookie)
        driver.get(url) #use cookie to open
    time.sleep(15)
    actions = ActionChains(driver) #ActionChains can real simulate the mouse to click
def WritePropertyeditor(label,keys,type): #Use for type text into the element
    xpath = f"//cmf-core-business-controls-propertyeditor[@data-varid='{label}']//input[@type='{type}']"
    WritePropertyeditorItem = driver.find_element(By.XPATH,xpath)
    #WritePropertyeditorItem.click()
    WritePropertyeditorItem.clear()
    WritePropertyeditorItem.send_keys(keys)
def WritePropertyeditorAndEnter(path,keys,type): #Use for type text into the element and click enter
    xpath = f"//cmf-core-business-controls-propertyeditor[@{path}]//input[@type='{type}']"
    WritePropertyeditorItem = driver.find_element(By.XPATH,xpath)
    #WritePropertyeditorItem.click()
    WritePropertyeditorItem.clear()
    time.sleep(1)
    WritePropertyeditorItem.send_keys(keys)
    time.sleep(1)
    WritePropertyeditorItem.send_keys(Keys.ARROW_DOWN)
    time.sleep(1)
    WritePropertyeditorItem.send_keys(Keys.RETURN)
    time.sleep(2)
def ClickPropertyeditor(path,flag): #Use for click the element
    try:
        xpath = f"//cmf-core-business-controls-propertyeditor[@{path}]//label[@class='switch']"
        ClickPropertyeditorItem = driver.find_element(By.XPATH,xpath)
        ClickPropertyeditorItem_input = ClickPropertyeditorItem.find_element(By.XPATH,"//input[@type='checkbox']")
        data_value = ClickPropertyeditorItem_input.get_attribute("data-value")
        print(ClickPropertyeditorItem_input.get_attribute("outerHTML"))
        print(str(data_value).lower())
        flag=str(flag).lower()
        if str(data_value).lower() != flag:
            print("Switch to "+ flag)
            ClickPropertyeditorItem.click()
    except:
        print("Error and click again")
        time.sleep(1)
        ClickPropertyeditor(path,flag)

def ClickPropertyeditorRadioOptionByDiv(path,option): #Use for click the element (RadioOption exsit in <div>)
    try:
        xpath = f"//cmf-core-business-controls-propertyeditor[@{path}]"
        Item = driver.find_element(By.XPATH,xpath)
        ClickItem = Item.find_element(By.XPATH, f"//div[contains(text(),'{option}')]")
        ClickItem.click()
    except:
        print("Error and click again")
        time.sleep(1)
        ClickPropertyeditorRadioOptionByDiv(path,option)

def ClickPropertyeditorRadioOptionBySpan(path,option): #Use for click the element (RadioOption exsit in <span>)
    try:
        xpath = f"//cmf-core-business-controls-propertyeditor[@{path}]"
        Item = driver.find_element(By.XPATH,xpath)
        ClickItem = Item.find_element(By.XPATH, f"//span[contains(text(),'{option}')]")
        ClickItem.click()
    except:
        print("Error and click again")
        time.sleep(1)
        ClickPropertyeditorRadioOptionBySpan(path,option)
def ClickPropertyviewerBySpan(path,option): #Use for click the element (RadioOption exsit in <span>)
    try:
        xpath = f"//cmf-core-business-controls-propertyviewer[@{path}]"
        Item = driver.find_element(By.XPATH,xpath)
        ClickItem = Item.find_element(By.XPATH, f"//span[contains(text(),'{option}')]")
        ClickItem.click()
    except:
        print("Error and click again")
        time.sleep(1)
        ClickPropertyviewerBySpan(path,option)
def GoNext(): #Click Next button
    time.sleep(1)
    button = driver.find_element(By.XPATH, "//button[@class='cmf-btn cmf-btn-primary']")
    actions.move_to_element(button).click().perform()
    time.sleep(6)
def GoNextFast(): #Click Next button
    time.sleep(1)
    button = driver.find_element(By.XPATH, "//button[@class='cmf-btn cmf-btn-primary']")
    actions.move_to_element(button).click().perform()
    time.sleep(2)
def GoBack(): #Click Back button
    time.sleep(1)
    button = driver.find_element(By.XPATH, "//button[@class='cmf-btn cmf-btn-secondary']")
    actions.move_to_element(button).click().perform()
    time.sleep(4)

def EditServiceResource(label,ReplicaNumber,MinMemoryValue,MaxMemoryValue,MinCPUValue, MaxCPUValue): # fill the Min Max Memory, Min Max CPU in
    try:
        time.sleep(1)
        DataGroupId=f"data-group-{label}"
        PanelbarElementXpath = f"//cmf-core-controls-panelbar[@data-labelid='{DataGroupId}']"
        PanelbarElement = driver.find_element(By.XPATH, PanelbarElementXpath)
        actions.move_to_element(PanelbarElement).click().perform()
        time.sleep(2)
        # ReplicaId=f"Kubernetes_{label}_Replicas"
        # ReplicasXpath = f"//cmf-core-business-controls-propertyeditor[@data-varid='{ReplicaId}']//input[@type='text']"
        # Replicas = PanelbarElement.find_element(By.XPATH,ReplicasXpath)
        # Replicas.clear()
        # Replicas.send_keys(str(ReplicaNumber))
        MinMemoryId=f"Kubernetes_{label}_MinMemory"
        MinMemoryXpath = f"//cmf-core-business-controls-propertyeditor[@data-varid='{MinMemoryId}']//input[@type='text']"
        MinMemory = PanelbarElement.find_element(By.XPATH,MinMemoryXpath)
        MinMemory.clear()
        MinMemory.send_keys(str(MinMemoryValue))

        MaxMemoryId=f"Kubernetes_{label}_MaxMemory"
        MaxMemoryXpath = f"//cmf-core-business-controls-propertyeditor[@data-varid='{MaxMemoryId}']//input[@type='text']"
        MaxMemory = PanelbarElement.find_element(By.XPATH,MaxMemoryXpath)
        MaxMemory.clear()
        MaxMemory.send_keys(str(MaxMemoryValue))

        MinCPUId=f"Kubernetes_{label}_MinCPU"
        MinCPUXpath = f"//cmf-core-business-controls-propertyeditor[@data-varid='{MinCPUId}']//input[@type='text']"
        MinCPU = PanelbarElement.find_element(By.XPATH,MinCPUXpath)
        MinCPU.clear()
        MinCPU.send_keys(str(MinCPUValue))


        MaxCPUId=f"Kubernetes_{label}_MaxCPU"
        MaxCPUXpath = f"//cmf-core-business-controls-propertyeditor[@data-varid='{MaxCPUId}']//input[@type='text']"
        MaxCPU = PanelbarElement.find_element(By.XPATH,MaxCPUXpath)
        MaxCPU.clear()
        MaxCPU.send_keys(str(MaxCPUValue))
    except:
        EditServiceResource(label,ReplicaNumber,MinMemoryValue,MaxMemoryValue,MinCPUValue, MaxCPUValue)

#Crawler Start
GoNext()
GoBack() #Handle some error so go next first and go back.

#1. Package
WritePropertyeditorAndEnter("data-label='Deployment Package'",Configuration["Deployment Package"],"text")
ClickPropertyeditorRadioOptionByDiv("data-label='Configuration Level'",Configuration["Configuration Level"])
WritePropertyeditorAndEnter("data-label='License'",Configuration["License"],"text")

GoNext()
ClickPropertyeditorRadioOptionBySpan("data-label='Database Mode'",Configuration["Database Mode"])
ClickPropertyeditor("data-label='Connect to a central Traefik'",Configuration["Connect to a central Traefik"])

GoNext()
WritePropertyeditorAndEnter("data-label='Target'",Configuration["Target"],"text")

GoNext()
ReadandUnderstood = driver.find_element(By.XPATH,"//div[@class='icheckbox_minimal']")
actions = ActionChains(driver)
actions.move_to_element(ReadandUnderstood).click().perform()

GoNext()
#Configuration
#General Data
WritePropertyeditor("SYSTEM_NAME",Configuration["SYSTEM_NAME"],"text")
WritePropertyeditor("TENANT_NAME",Configuration["TENANT_NAME"],"text")
WritePropertyeditor("APPLICATION_PUBLIC_HTTP_ADDRESS",Configuration["APPLICATION_PUBLIC_HTTP_ADDRESS"],"text")
ClickPropertyeditor("data-varid='APPLICATION_PUBLIC_HTTP_TLS_ENABLED'",Configuration["APPLICATION_PUBLIC_HTTP_TLS_ENABLED"])
WritePropertyeditor("INSTALLATION_DATA_VOLUME_PATH",Configuration["INSTALLATION_DATA_VOLUME_PATH"],"text")
GoNext()


WritePropertyeditor("DATABASE_ONLINE_MSSQL_ADDRESS",Configuration["DATABASE_ONLINE_MSSQL_ADDRESS"],"text")
WritePropertyeditor("DATABASE_ONLINE_MSSQL_USERNAME",Configuration["DATABASE_ONLINE_MSSQL_USERNAME"],"text")
WritePropertyeditor("DATABASE_ONLINE_MSSQL_PASSWORD",Configuration["DATABASE_ONLINE_MSSQL_PASSWORD"],"password")

WritePropertyeditor("DATABASE_ODS_MSSQL_ADDRESS",Configuration["DATABASE_ODS_MSSQL_ADDRESS"],"text")
WritePropertyeditor("DATABASE_ODS_MSSQL_USERNAME",Configuration["DATABASE_ODS_MSSQL_USERNAME"],"text")
WritePropertyeditor("DATABASE_ODS_MSSQL_PASSWORD",Configuration["DATABASE_ODS_MSSQL_PASSWORD"],"password")

WritePropertyeditor("DATABASE_DWH_MSSQL_ADDRESS",Configuration["DATABASE_DWH_MSSQL_ADDRESS"],"text")
WritePropertyeditor("DATABASE_DWH_MSSQL_USERNAME",Configuration["DATABASE_DWH_MSSQL_USERNAME"],"text")
WritePropertyeditor("DATABASE_DWH_MSSQL_PASSWORD",Configuration["DATABASE_DWH_MSSQL_PASSWORD"],"password")


WritePropertyeditor("DATABASE_AS_MSAS_ADDRESS",Configuration["DATABASE_AS_MSAS_ADDRESS"],"text")
WritePropertyeditor("DATABASE_AS_MSAS_USERNAME",Configuration["DATABASE_AS_MSAS_USERNAME"],"text")
WritePropertyeditor("DATABASE_AS_MSAS_PASSWORD",Configuration["DATABASE_AS_MSAS_PASSWORD"],"password")

GoNext()

ClickPropertyeditor("data-varid='SECURITY_PORTAL_STRATEGY_LOCAL_AD_ENABLED'",Configuration["SECURITY_PORTAL_STRATEGY_LOCAL_AD_ENABLED"])
WritePropertyeditor("SECURITY_PORTAL_STRATEGY_LOCAL_AD_DEFAULT_DOMAIN",Configuration["SECURITY_PORTAL_STRATEGY_LOCAL_AD_DEFAULT_DOMAIN"],"text")
WritePropertyeditor("SECURITY_PORTAL_STRATEGY_LOCAL_AD_SERVER_ADDRESS",Configuration["SECURITY_PORTAL_STRATEGY_LOCAL_AD_SERVER_ADDRESS"],"text")

WritePropertyeditor("SECURITY_PORTAL_STRATEGY_LOCAL_AD_SERVER_BASE_DN",Configuration["SECURITY_PORTAL_STRATEGY_LOCAL_AD_SERVER_BASE_DN"],"text")
WritePropertyeditor("SECURITY_PORTAL_STRATEGY_LOCAL_AD_USERNAME",Configuration["SECURITY_PORTAL_STRATEGY_LOCAL_AD_USERNAME"],"text")
WritePropertyeditor("SECURITY_PORTAL_STRATEGY_LOCAL_AD_PASSWORD",Configuration["SECURITY_PORTAL_STRATEGY_LOCAL_AD_PASSWORD"],"password")

GoNext()
WritePropertyeditor("REPORTING_SSRS_WEB_PORTAL_URL",Configuration["REPORTING_SSRS_WEB_PORTAL_URL"],"text")
WritePropertyeditor("REPORTING_SSRS_WEB_SERVICE_URL",Configuration["REPORTING_SSRS_WEB_SERVICE_URL"],"text")

WritePropertyeditor("REPORTING_SSRS_USERNAME",Configuration["REPORTING_SSRS_USERNAME"],"text")
WritePropertyeditor("REPORTING_SSRS_PASSWORD",Configuration["REPORTING_SSRS_PASSWORD"],"password")
GoNextFast() #5
GoNextFast() #6
GoNextFast() #7
GoNextFast() #8
GoNextFast() #9 Email
WritePropertyeditor("EMAIL_FROM_ADDRESS",Configuration["EMAIL_FROM_ADDRESS"],"text")
WritePropertyeditor("EMAIL_SMTP_ADDRESS",Configuration["EMAIL_SMTP_ADDRESS"],"text")
WritePropertyeditor("EMAIL_SMTP_PORT",Configuration["EMAIL_SMTP_PORT"],"text")
GoNextFast()() #10
GoNext() #11 Service 
ClickPropertyeditor("data-varid='ACTION_MANAGER_BOOT_SYNC_ENABLED'",Configuration["ACTION_MANAGER_BOOT_SYNC_ENABLED"])
WritePropertyeditor("Kubernetes_Services_Global_DNS",Configuration["DNS"],"text")
GoNext()
for Service, Resource in ServiceResource.items():
    EditServiceResource(Service, Resource[0], Resource[1], Resource[2], Resource[3], Resource[4])
GoNext()
for Item, Value in Volumes.items():
    if Value[0]=="NFS": #type=NFS
        try:
            WritePropertyeditorAndEnter(f"data-varid='Kubernetes_Volume_Type_{Item}'",Value[0],"text")
            time.sleep(1)
            WritePropertyeditor(f"Kubernetes_Volume_Size_{Item}",Value[1],"text")
            WritePropertyeditor(f"Kubernetes_Volume_NFS_Server_{Item}",Value[2],"text")
            WritePropertyeditor(f"Kubernetes_Volume_NFS_Path_{Item}",Value[3],"text")
        except:
            time.sleep(1)
            WritePropertyeditorAndEnter(f"data-varid='Kubernetes_Volume_Type_{Item}'",Value[0],"text")
            time.sleep(1)
            WritePropertyeditor(f"Kubernetes_Volume_Size_{Item}",Value[1],"text")
            WritePropertyeditor(f"Kubernetes_Volume_NFS_Server_{Item}",Value[2],"text")
            WritePropertyeditor(f"Kubernetes_Volume_NFS_Path_{Item}",Value[3],"text")

    if Value[0]=="Existing Storage Class": #type=Existing Storage Class
        try:
            WritePropertyeditorAndEnter(f"data-varid='Kubernetes_Volume_Type_{Item}'",Value[0],"text")
            time.sleep(1)
            WritePropertyeditor(f"Kubernetes_Volume_Size_{Item}",Value[1],"text")
            WritePropertyeditor(f"Kubernetes_Volume_Existing_Storage_Class_Name_{Item}",Value[2],"text")
        except:
            time.sleep(1)
            WritePropertyeditorAndEnter(f"data-varid='Kubernetes_Volume_Type_{Item}'",Value[0],"text")
            time.sleep(1)
            WritePropertyeditor(f"Kubernetes_Volume_Size_{Item}",Value[1],"text")
            WritePropertyeditor(f"Kubernetes_Volume_Existing_Storage_Class_Name_{Item}",Value[2],"text")
GoNext()
time.sleep(15)
ClickPropertyviewerBySpan("data-label='Deployment Artifact'","CustomerEnvironment") #Download Package
time.sleep(30)
driver.quit()





