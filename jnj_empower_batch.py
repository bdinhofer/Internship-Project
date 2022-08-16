import os
import pandas as pd
import zipfile
import re
import win32com.client
import time
import warnings
from pathlib import Path
from configparser import ConfigParser
from datetime import date

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException,UnexpectedAlertPresentException,NoAlertPresentException,ElementNotInteractableException,ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

pd.options.mode.chained_assignment = None  # default='warn'
warnings.filterwarnings("ignore", category=DeprecationWarning) 

request_path = os.path.join(os.getcwd(),"JnJ Media Request Template.xlsx")
downloads_path = str(Path.home() / "Downloads")
optima_login_flag = False

config = ConfigParser()
config.read(os.path.join(os.getcwd(),"config.ini"))

network_username = config.get("Network","username")
network_password = config.get("Network","password")
egnyte_username = config.get("Egnyte","username")
egnyte_password = config.get("Egnyte","password")
outlook_folder_path = config.get("Outlook","empower_folder_path")

egnyte_url = "https://cmicompas.egnyte.com/timeout.noauth"
empower_url = "https://empower.cmicompas.com/Account/Login"
empower_pull_url = "https://empowerhcp.cmicompas.com/"

service = Service(executable_path=ChromeDriverManager().install())
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
browser = webdriver.Chrome(service=service,options=options)
browser.maximize_window()

def find_empower_folder(outlook_folder_path,outlook):
    try:
        curr_folder =  outlook.Folders[outlook_folder_path.split("/")[0]]
        for path_section in outlook_folder_path.split('/')[1:]:
            try: 
                curr_folder = curr_folder.Folders[path_section]
            except:
                print(path_section +" is invalid")
                return None
        return curr_folder
    except:
        print("Email is invalid")
    return curr_folder

def switch_to_tab(url_keyword,browser):
    current_w = browser.current_window_handle
    for w in browser.window_handles:
        browser.switch_to.window(w)
        if url_keyword in browser.current_url:
            break
        else:
            browser.switch_to.window(current_w)

def get_download_link(num_of_emails):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    links_folder = find_empower_folder(outlook_folder_path,outlook)
    messages = links_folder.Items
    while len(messages) == num_of_emails:
        time.sleep(5)
    message = messages[len(messages) - 1]
    time.sleep(1)
    print("New email found!")
    return re.search(r'(?<=href=").*(?=" original)', message.HTMLBody)[0].strip(), re.search("(?<=Buy File:</b> <i>).*\.txt", message.HTMLBody)[0].strip()[:-4]

def unzip(f):
    file_path = os.path.join(downloads_path,f)
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(downloads_path)
    os.remove(file_path)

def process_tl(f):
    file_path = os.path.join(downloads_path,f)
    df = pd.read_csv(file_path,sep="\t",dtype=str)
    df["ME_ID"] = df["ME_ID"].str.zfill(10)
    df["NPI_ID"] = df["NPI_ID"].str.zfill(10)
    df["Postal_Code"] = df["Postal_Code"].str.zfill(5)
    df.to_csv(file_path,sep="\t",index=False)

def process_and_upload_tl(curr_egnyte_path,num_of_emails):
    download_link,base_file_name = get_download_link(num_of_emails)
    print("Link retrieved, now downloading file!")
    file_name = base_file_name+".txt"
    zip_file_name = base_file_name+".zip"
    browser.get(download_link)
    while max([os.path.join(downloads_path, basename) for basename in os.listdir(downloads_path)], key = os.path.getctime) != os.path.join(downloads_path, zip_file_name):
        time.sleep(5)
    unzip(zip_file_name)
    process_tl(file_name)
    print("Uploading file to Egnyte!")
    browser.execute_script('''window.open("{0}","_blank");'''.format(curr_egnyte_path))
    switch_to_tab("egnyte",browser)
    time.sleep(3)
    browser.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/button[1]").click() # Upload Button
    time.sleep(1) 
    egnyte_folder = browser.find_element(By.CLASS_NAME, "folder-items")
    num_of_files = len(egnyte_folder.find_elements(By.XPATH, "//div[@role='row']"))
    browser.find_elements_by_xpath("//input")[-1].send_keys(os.path.join(downloads_path,file_name))
    while len(egnyte_folder.find_elements(By.XPATH, "//div[@role='row']")) == num_of_files:
        time.sleep(5)
    browser.refresh()
    time.sleep(5)
    load_complete_flag = False
    while load_complete_flag == False:
        try:
            browser.find_element_by_xpath("//span[text()='{0}']".format(file_name)).click()
            load_complete_flag = True
        except (ElementNotInteractableException,NoSuchElementException):
            time.sleep(3)
    #browser.find_element_by_xpath("//span[text()='{0}']".format(file_name)).click()
    time.sleep(3)
    browser.find_element_by_xpath("//button[text()='Share...']").click()
    time.sleep(1)
    browser.find_element_by_xpath("//div[text()='Share Link']").click()
    time.sleep(2)
    try:
        browser.find_element_by_xpath("/html/body/div[3]/div/div/div/form/div/div[2]/button[1]").click() #Get Link Button
    except:
        browser.find_element_by_xpath("//svg[text()='Get Link']").click()
    time.sleep(1)
    link = browser.find_element_by_xpath("//textarea").text
    time.sleep(1)
    browser.find_element_by_xpath("//button[text()='Close']").click()
    time.sleep(1)
    os.remove(os.path.join(downloads_path,file_name))
    print("Share link retrieved!")
    return link.replace("\n","<br>"), file_name

def send_media_confirmation(request_info,request_path):
    today = date.today()
    total_emails = []
    emails = [i for i in request_info["requested_by"].apply(lambda x:x.split(";"))]
    for email in emails:
        total_emails += email
    outlook = win32com.client.Dispatch("Outlook.Application")
    Msg = outlook.CreateItem(0)
    Msg.To = ';'.join(list(set(total_emails)))
    Msg.Subject = f'Execution File Request - Confirmation Email {today.month}/{today.day}/{today.year}'

    Msg.HTMLBody =  f"""\
    <html>
    <head></head>
    <body>
    <p>Hello,<br> <br> 
    Attached is a document confirming all the execution file requests that have been fulfilled along with the corresponding file name sent to the supplier.
        <br><br>Please let me know if you have any questions.
        <br><br>Thanks,
    </p>
    </body>
    </html>
    """
    Msg.Attachments.Add(request_path)
    Msg.Save()

    print("Confirmation email sent to Media Team")


def send_email(main_email,cc_emails,brand,client,optima_vehicle_name,total_count,placement_description,link):
    print("Sending email!") 
    try:
        client = client.split(' ')[0] 
    except:
        pass  
    outlook = win32com.client.Dispatch("Outlook.Application")
    Msg = outlook.CreateItem(0)
    Msg.To = main_email
    if not pd.isna(cc_emails):
        Msg.CC = cc_emails
    Msg.Subject = f'Execution File Request {client} - {brand}'
    Msg.HTMLBody =  f"""\
    <html>
    <head></head>
    <body>
    <p>Hello,<br> <br> 
    I have attached a link to the Execution file for {client} - {brand}. This file contains prescriber information that includes available identifiers, names, addresses, and corresponding details needed for the upcoming program deployment and the associated required data feeds post-deployment.
    <br><br> Vehicle Name: {optima_vehicle_name} 
    <br> Record Count: {total_count}
    <br> Placement Description: {placement_description}<br><br>

    {link}

        <ul>
    <li>The CMI planning team should be able to provide details how and when the buy files should be used for execution.</li>
    <li>Please note, CMI planning teams are not permitted to view HCP level data per CMI policy, please deactivate or remove links prior to including planning team on email chain</li>
    </ul>
        <br><br>Upon receiving this email, please confirm receipt of receiving the files by replying to all here.
        <br><br>Please let me know if you have any questions.
        <br><br>Thanks,
    </p>
    </body>
    </html>
    """
    Msg.Save()
    print("Email Written")

def empower_pull(brand, client, target_list_id, me_flag,placement_id,vehicle_name,channel,reachable_or_active,segments,optima_login_flag,total_count):
    #Select brand
    select = Select(browser.find_element_by_id("brandSelectAlt"))
    time.sleep(1)
    try:
        select.select_by_visible_text(brand + ' - ' + client) 
        time.sleep(2)
    except NoSuchElementException:
        print("The brand/client, '" + brand + ' - ' + client + "', was not found in Empower.  Check to make sure spelling of brand and client name is correct")
        return "", True, "", "", optima_login_flag

    time.sleep(1)
    #Select list and detect deactivated list
    select = Select(browser.find_element_by_id("listSelectAlt")) 
    target_list_options = browser.find_elements(By.XPATH, "//*[@id='listSelectAlt']/option")
    for option in target_list_options:
        try:
            if target_list_id in option.text:
                select.select_by_value("import-" + str(target_list_id))
        except NoSuchElementException:
            print("The Target List, '" + target_list_id + "', was not found. Check Spelling.")
            return "", True, "", "", optima_login_flag
    time.sleep(1)

    #Selecting the apply button, testing for deactivated target list
    browser.find_element_by_id("changeList").click()
    try:
        browser.find_element(By.ID, "toast-container")
        print("The Target List, '" + target_list_id + "', has been deactivated. Check that you are using the correct one.")
        return "", True, "", "", optima_login_flag
    except NoSuchElementException:
        pass
    time.sleep(5)
    browser.maximize_window()
    load_complete_flag = False
    while load_complete_flag == False:
        try:
            browser.find_element_by_id("exportBrandLink").click()
            load_complete_flag = True
        except (ElementNotInteractableException,NoSuchElementException):
            time.sleep(3)
    #Open export menu
    #browser.find_element_by_id("exportBrandLink").click() 
    time.sleep(3)
    #Check buy file option and enter placement id
    load_complete_flag = False
    while load_complete_flag == False:
        try:
            browser.find_element_by_id("exportFileBuy").click()
            load_complete_flag = True
        except (ElementNotInteractableException,NoSuchElementException):
            time.sleep(3)

    input_pid = False
    while not input_pid:
        try:
            browser.find_element_by_id("exportPlacementId").send_keys(placement_id)
            input_pid = True
        except ElementNotInteractableException:
            browser.find_element_by_xpath("//i[@class='material-icons tiny close-pop-up']").click()
            browser.find_element_by_id("exportBrandLink").click()
            browser.find_element_by_id("exportFileBuy").click()

    browser.find_element_by_xpath("//*[@id='exportBrand']/div[1]").click()
    time.sleep(3)
    try: 
        browser.find_element(By.ID, "toast-container")
        print("The Placement ID, '" + placement_id + "', doesn't exist or is incorrect. Check that you are using the correct one.")
        return "", True, "", "", optima_login_flag
    except (NoSuchElementException,UnexpectedAlertPresentException):
        pass
    time.sleep(2)
    try:
        browser.switch_to.alert.accept()
    except NoAlertPresentException:
        pass
    #Select vehicle
    vehicle_select_name = vehicle_name+' ['+channel+"]"
    browser.find_element_by_xpath("//li[span/text()='Vehicle Reach'][i/text()='add_box']").find_element_by_xpath("i").click()  #open column filter menu
    time.sleep(1)
    if reachable_or_active=="active":
        browser.find_element_by_xpath("//input[@value='active']").click()

    unclicked = True
    while unclicked:
        try:
            browser.find_element_by_xpath("//button[@class='ui-multiselect ui-widget ui-state-default ui-corner-all']").click()
            unclicked = False
        except ElementNotInteractableException:
            continue 

    time.sleep(3) 
    emp_v_names = browser.find_elements(By.XPATH, "/html/body/div[25]/ul/li")
    vehicle_error_flag = True
    for v in emp_v_names:
        if v.text.lower() == vehicle_select_name.lower():
            v.click()
            browser.find_element_by_xpath("//*[@id='export-filters-prompt']/div[3]/button[2]/b").click()
            vehicle_error_flag = False
            break
    if vehicle_error_flag == True:
        print("Could not find, '" + vehicle_select_name + "' in Empower. Check vehicle name and channel.")
        return "", True, "", "", optima_login_flag
    #Select segments
    segment_title_errors = []
    segment_name_errors = []
    for segment_title in list(segments.keys()):
        try:
            browser.find_element_by_xpath("//li[span/text()="+"'"+segment_title+"'"+"][i/text()='add_box']").find_element_by_xpath("i").click()
            time.sleep(1)
            unclicked = True
            while unclicked:
                try:
                    browser.find_element_by_xpath("//button[@class='ui-multiselect ui-widget ui-state-default ui-corner-all']").click()
                    unclicked = False
                except ElementNotInteractableException:
                    continue
        except NoSuchElementException:
            segment_title_errors.append(segment_title)
            continue
        time.sleep(3)
        for segment_name in segments[segment_title]:
            try:
                browser.find_element(By.XPATH, "//label[span/text()="+"'"+segment_name+"'"+"]").click()
                time.sleep(0.5)
            except NoSuchElementException:
                segment_name_errors.append(segment_name)
                continue
        browser.find_element_by_xpath("//*[@id='export-filters-prompt']").click()
        submitted_flag = False
        while not submitted_flag:
            try:
                browser.find_element_by_xpath("//*[@id='export-filters-prompt']/div[3]/button[2]/b").click()
                submitted_flag = True
            except ElementClickInterceptedException:
                browser.find_element_by_xpath("//*[@id='export-filters-prompt']").click()
    if len(segment_title_errors) > 0 or len(segment_name_errors) > 0: 
        if len(segment_title_errors) > 0:
            print("The Following Segment Title's could not be found: " + str(segment_title_errors))
        if len(segment_name_errors) > 0:
            print("The Following Segment Names could not be found: " + str(segment_name_errors))
        return "", True, "", "", optima_login_flag
 
    #Check if count in Empower matches request count
    time.sleep(1)
    empower_record_count = int(browser.find_element(By.XPATH, "//*[@id='exportBrand']/div[2]/div[1]/div[3]/div/span/span").text.replace(' ', ''))
    if int(total_count) == empower_record_count:
        browser.find_element_by_xpath("//*[@id='submitBtn']").click() #Submitting to next page
    else:
       print("Total record count doesn't match. Empower shows, {} records, while you entered {}. Check segment titles and segment names.".format(empower_record_count, total_count))
       return "", True, "", "", optima_login_flag

    #Select appropriate upload fields
    select = Select(browser.find_element_by_xpath("//*[@id='uploadFields']/li[16]/select"))
    select.select_by_visible_text("AMA Specialty")

    for ix in range(len(segments.keys())):
        select = Select(browser.find_element_by_xpath("//*[@id='uploadFields']/li["+str(18+ix)+"]/select")) 
        select.select_by_visible_text(list(segments.keys())[ix])

    #Open Optima
    browser.switch_to.default_content()
    browser.execute_script('''window.open("https://optima.compasonline.com/", "");''')
    browser.switch_to.window(browser.window_handles[1])
    #Sign in only once
    if not optima_login_flag:
        browser.find_element(By.ID, "UserName").send_keys(network_username)
        browser.find_element(By.ID, "LoginPassword").send_keys(network_password)
        browser.find_element(By.ID, "UserIDSubmit").click()
        optima_login_flag = True
    browser.switch_to.frame("menu")
    browser.find_element(By.ID, "Companies").click() #Clicking companies button
    browser.switch_to.default_content()
    browser.switch_to.frame("content")
    browser.find_element(By.ID, "ctl00_MainContent_CompanySearchAnchor_Orders").click() 
    browser.find_element(By.ID, "ctl00_MainContent_JobSearchAnchor_Media").click()
    browser.find_element(By.ID, "ctl00_MainContent_PlacementID").send_keys(placement_id)
    browser.find_element(By.ID, "ctl00_MainContent_Search").click()
    time.sleep(5)
    #Read Empower table for campaign details
    load_complete_flag = False
    while load_complete_flag == False:
        try:
            table = browser.find_element(By.ID, "ctl00_MainContent_MediaJobPlacementsControl")
            load_complete_flag = True
        except NoSuchElementException:
            time.sleep(3)
    rows = table.find_elements(By.XPATH, '//*[@id="ctl00_MainContent_MediaJobPlacementsControl"]/tbody/tr')
    for row in range(len(rows)):
        columns = rows[row].find_elements(By.XPATH, '//*[@id="ctl00_MainContent_MediaJobPlacementsControl"]/tbody/tr[' + str(row + 1) +']/td')
        campaign_name = columns[1].text
        placement_description = columns[8].text
        optima_vehicle_name = columns[6].text
        break
    #Switch back to Empower and add campaign name
    browser.close()
    browser.switch_to.window(empower_window)
    browser.find_element_by_xpath("//*[@id='uploadFields']/li[17]/input[2]").send_keys(campaign_name)

    #Change ME to not show 
    if me_flag == '0':
        select = Select(browser.find_element(By.XPATH, "//*[@id='uploadFields']/li[4]/select"))
        select.select_by_visible_text("--Select Field to Match--")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    links_folder = find_empower_folder(outlook_folder_path,outlook)
    num_of_emails = len(links_folder.Items)
    #Submit request
    browser.find_element(By.XPATH, '//*[@id="buttonSubmitMatchedFields"]/b').click()
    browser.switch_to.alert.accept()
    #Refresh Empower page
    browser.get(empower_pull_url)

    return num_of_emails,False,placement_description,optima_vehicle_name,optima_login_flag

def process_buy_file(me_flag,placement_id,vehicle_name,channel,reachable_or_active,target_list_id,client,brand,main_email,cc_emails,segments,curr_egnyte_path,optima_login_flag,total_count):
    browser.switch_to.window(empower_window)
    print("Pulling list from Empower")
    num_of_emails, error_flag, placement_description,optima_vehicle_name,optima_login_flag = empower_pull(brand, client, target_list_id, me_flag,placement_id,vehicle_name,channel,reachable_or_active,segments,optima_login_flag,total_count)
    file_name = ''
    if not error_flag:
        current_link, file_name = process_and_upload_tl(curr_egnyte_path,num_of_emails)
        browser.close()
        send_email(main_email,cc_emails,brand,client,optima_vehicle_name,total_count,placement_description,current_link)
    return optima_login_flag, error_flag, file_name

def open_xlsx_file(path):
    file_opened = False
    while not file_opened:
        try:
            file = pd.read_excel(path,dtype=str)
            file_opened = True
        except PermissionError:
            print("Error - File is still opened")
            input("Please save and close file then press Enter to continue")
    return file

valid_outlook_path = False
outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
links_folder = find_empower_folder(outlook_folder_path,outlook)
if links_folder is None:
    browser.quit()
    input("Email folder path is invalid, please close program and make corrections")
else:
    valid_outlook_path = True

egnyte_login = False
empower_login = False
browser.get(egnyte_url)
username_elem = browser.find_element_by_id("loginUsername")
username_elem.send_keys(egnyte_username)
browser.find_element_by_class_name("set-username-btn").click()
time.sleep(1)
try: 
    browser.find_element_by_xpath("//span[text()='No account found']")
    browser.quit()
    input("Egnyte username invalid, please close program and make corrections")
except NoSuchElementException:
    time.sleep(3)
    password_elem = browser.find_element_by_id("j_password")
    password_elem.send_keys(egnyte_password)
    browser.find_element_by_id("loginBtn").click()
    try_login = True
    while try_login:
        try:
            browser.find_element_by_xpath("//span[text()='Invalid password']")
            browser.quit()
            input("Egnyte password invalid, please close program and make corrections")
        except NoSuchElementException:
            try:
                browser.find_element_by_xpath("//div[@id='login-g-recaptcha']")
                password_elem = browser.find_element_by_id("j_password")
                password_elem.send_keys(egnyte_password)
                input("Please complete reCAPTCHA test then hit Enter in program to continue")
                browser.find_element_by_id("loginBtn").click()
            except NoSuchElementException:
                print("Successfully logged into Egnyte")
                egnyte_login = True
                try_login = False
        
if egnyte_login:
    browser.execute_script('''window.open("{0}","_blank");'''.format(empower_url))
    switch_to_tab("egnyte",browser)
    browser.close()
    browser.switch_to.window(browser.window_handles[0])
    username_elem = browser.find_element_by_id("user-input")
    password_elem = browser.find_element_by_id("pw-input")

    username_elem.send_keys(network_username)
    password_elem.send_keys(network_password)

    browser.find_element_by_class_name("signin-button").click()
    time.sleep(1)
    try:
        browser.find_element_by_xpath("//span[text()='Professionals']").click()
        empower_login = True
    except NoSuchElementException:
        browser.quit()
        input("Invalid Empower credentials, please close program and make corrections")

    if empower_login:
        print("Successfully logged into Empower!")
        switch_to_tab("empower.cmicompas.com",browser)
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        empower_window = browser.current_window_handle

if valid_outlook_path and empower_login and egnyte_login:
    request_info = open_xlsx_file(request_path)
    request_info.reset_index(inplace=True)
    segment_titles_count = request_info.filter(regex="SegmentTitle").shape[1]
    counter = 0
    while list(set(request_info["processed"].values)) != ["1"]:
        counter+=1
        print("Run #{} -----------------------------------------------------------------------------------------------------".format(counter))
        for ix,row in request_info[request_info["processed"]!="1"].iterrows():
            print("Now processing request #{}".format(int(row["index"])+1))
            brand = row["brand"].strip()
            client = row["client"].strip()
            target_list_id  = row["target_list_id"].strip()
            me_flag = row["me_flag"].strip()
            placement_id = row["placement_id"].strip()
            vehicle_name = row["vehicle_name"].strip()
            channel = row["channel"].strip()
            reachable_or_active = row["reachable_or_active"].strip()
            main_email = row["main_email"]
            cc_emails = row["cc_emails"]
            curr_egnyte_path = row["egnyte_directory"]
            total_count= row["total_count"]
        
            segments = {}
            for i in range(1,segment_titles_count+1):
                segment = row["SegmentTitle"+str(i)]
                segment_names_count = request_info.filter(regex="segment"+str(i)+"_name").shape[1]
                if not pd.isna(segment):
                    segment_names = []
                    for j in range(1,segment_names_count+1):
                        segment_name = row["segment"+str(i)+"_name"+str(j)]
                        if not pd.isna(segment_name):
                            segment_names.append(segment_name)
                    segments[segment] = segment_names

            optima_login_flag, error_flag, file_name = process_buy_file(me_flag,placement_id,vehicle_name,channel,reachable_or_active,target_list_id,client,brand,main_email,cc_emails,segments,curr_egnyte_path,optima_login_flag,total_count)
    
            if not error_flag:
                request_info.at[ix,"processed"] = "1"
                request_info.at[ix,"file_name"] = file_name
                print("Successfully processed request #{}".format(ix+1))
                    
            file_written = False
            while not file_written:
                try:
                    request_info_to_write = request_info.drop(columns=["index"])
                    request_info_to_write.to_excel(request_path, index=False)
                    file_written = True
                except PermissionError:
                    print("Error - File is still opened")
                    input("Please save and close file then press Enter to continue")

            #Hard reset after every request
            for w in  browser.window_handles:
                if w != empower_window:
                    browser.switch_to.window(w)
                    browser.close()
            browser.switch_to.window(empower_window)
            browser.get(empower_pull_url)
        
        if list(set(request_info["processed"].values)) != ["1"]:
            print("Please address issues in listed requests")
            input("Once adjustments have been made to the original file, press Enter to rerun the process")
            request_info = open_xlsx_file(request_path)
            request_info.reset_index(inplace=True)


#send_media_confirmation(request_info[request_info["processed"]=="1"],request_path)

browser.quit()
input("All requests successfully executed, press Enter to close program")