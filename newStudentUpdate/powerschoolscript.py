from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import os
import subprocess
import time
import socket
from selenium.webdriver.chrome.options import Options
import os
import pandas as pd
from datetime import datetime
from datetime import timedelta
import json
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from cryptography.fernet import Fernet

 
def main():

    chrome_options= webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=chrome")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")\
    
 #generate key and open key.key

#opens key.key and assigns the key stored as key

# 
 
 #chrome_prefs = {"download.default_directory": r"C:\path\to\Downloads"} # (windows)
 #chrome_options.experimental_options["prefs"] = chrome_prefs

    
    ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__)))
    driver.f
    
    if os.path.exists('filekey.key'):
        k = open('filekey.key')
        key = k.read()
        

    if os.path.exists('pslogin.json'):
        print("Logging In...")
    else:
        key = getLogin()
        encrypt(key)
    decrypt(key)
    f = open('pslogin.json')
    login= json.load(f)
    grade=2
    driver = webdriver.Chrome(options=chrome_options)
 #
 # url = "https://powerschool.npd117.net/admin/pw.html"
 
    logIn(driver,login)
    findSchool(driver,login)
    f.close()
    while grade < 6:
    
    
        hn=os.environ["USERNAME"]
        path1= "C:/Users/"+hn+ "/Downloads/"
   #read_file=pd.DataFrame()
        currGrade="grade_"+str(grade)
    
   # print("Now pulling Grade "+ str(grade))
        findGrade(driver,currGrade)
        export(driver)
        if grade == 5:
            driver.quit()  
    
        grade+=1
    encrypt(key)
 #print(grade)
    
    grade-=1
    filePos=grade-grade
    while grade > 1:
        convertToXlsx(path1, grade, filePos,ROOT_DIR)
        grade-=1
        filePos+=2
def getLogin():
   
    username=input('Enter Username: ')
    password=input('Enter Password: ')
    url = input('Enter URL: ')
    school = input('Enter school(GO or OR): ')
    if school == 'OR' or 'or':
        school= 'school_5'
    elif school == 'GO' or 'go':
        school='school_4'
    elif school == 'CN' or 'cn':
        school='school_3'
    dictionary = {
        "username": username,
        "password": password,
        "url": url,
        "school": school
    }
    
    json_object = json.dumps(dictionary, indent=4)
    with open("pslogin.json", "w") as outfile:
     outfile.write(json_object) 
    
    key = Fernet.generate_key()
    
 
# string the key in a file
    with open('filekey.key', 'wb') as filekey:
        filekey.write(key)
    return key
    
 
        
def encrypt(key):
    # opening the key
    with open('filekey.key', 'rb') as filekey:
        key = filekey.read()
 
    # using the generated key
    fernet = Fernet(key)
 
    # opening the original file to encrypt
    with open('pslogin.json', 'rb') as file:
        original = file.read()
     
    # encrypting the file
    encrypted = fernet.encrypt(original)
 
# opening the file in write mode and
# writing the encrypted data
    with open('pslogin.json', 'wb') as encrypted_file:
        encrypted_file.write(encrypted)
def decrypt(key):
    # using the key
    fernet = Fernet(key)
 
# opening the encrypted file
    with open('test.json', 'rb') as enc_file:
        encrypted = enc_file.read()
 
# decrypting the file
    decrypted = fernet.decrypt(encrypted)
 
# opening the file in write mode and
# writing the decrypted data
    with open('test.json', 'wb') as dec_file:
        dec_file.write(decrypted)
def logIn(driver, login):
    #uses user input to login
    
    url=login['url']
    userName=login['username']
    passWord = login['password']
    
    driver.get(url)
    input= driver.find_element("id","fieldUsername")
    input.send_keys(userName)
    input= driver.find_element("id","fieldPassword")
    input.send_keys(passWord)
    input.submit()
    time.sleep(3)#change this to wait for element to load
def findSchool(driver,login):
    school = login['school']
    print("Finding school...")
    url = driver.current_url
    driver.get(url)
    schoolpicker=driver.find_element("id","adminSchoolPicker")
    schoolpicker.click()
    time.sleep(7)
    schoolpicker=driver.find_element("id",school)
    schoolpicker.click()
    time.sleep(8)
def findGrade(driver,currGrade):
    print("Pulling " + currGrade+ ": ")
    url = driver.current_url
    driver.get(url)
    
    gradeLvl=driver.find_element("id", "content-main")
   
    gradeLvl=driver.find_element("id",currGrade)
    gradeLvl.click()
    time.sleep(3)
    url = driver.current_url
    driver.get(url)
    time.sleep(3)
    
    functions=driver.find_element("id","selectFunctionDropdownButtonStudent")
    functions.click()
    functions=driver.find_element("id","lnk_studentsQuickExport")
    functions.click()
    time.sleep(3)
def export(driver):
    print("Exporting...")
    url = driver.current_url
    driver.get(url)
    export=driver.find_element("id","tt")
    export.clear()#clear field
    time.sleep(5)
    #modify with necessary data flags
    export.send_keys("SchoolEntryDate \n SchoolID \n Student_Number \nState_StudentNumber \n Last_Name \nFirst_Name \n^(*period_info;AM-PM(A);section_number)   \n^(*period_info;MA(A);section_number) \n^(*period_info;MA(A);teacher_name)  \n^(*period_info;ELA(A);section_number) \n^(*period_info;ELA(A);teacher_name) \nGender \n Grade_Level \n S_IL_STU_Demographics_X.Home_Language_Code \n IL_LEP \n S_IL_STU_X.Disability \n S_IL_STU_Plan504_X.Participant \n IL_FER \n Ethnicity\n IL_LII ")
    export=driver.find_element("id","btnSubmit")
    export.click()
    time.sleep(10)
    export=driver.find_element("id","branding-powerschool")
    export.click() 
def convertToXlsx(path1,grade,filePos, ROOT_DIR):
    os.chdir(path1)
    files = filter(os.path.isfile, os.listdir(path1))
    files = [os.path.join(path1, f) for f in files] # add path to each file
    files.sort(key=lambda x: -os.stat(x).st_mtime)
    openFile=files[filePos]
    #checks if file exists
    if os.path.exists(openFile):
        print('The file exists')
    else:
     print('The specified file does NOT exist')
    daytime = datetime.now()
    df = pd.read_csv(openFile ,sep='\t', lineterminator='\r')
    xlsxFile="UpdatedStudentListGrade"+str(grade)+'_'+daytime.strftime("%m%d%Y")+ ".xlsx"
    df.to_excel(xlsxFile, index = False)
   # read_file = pd.read_excel(xlsxFile, sheet_name="Sheet1")
    pd.set_option('display.max_columns', None)
    Fri = daytime.today() - timedelta(days=3)
    Thurs  = daytime.today() - timedelta(days=4)
    Wedn  = daytime.today() - timedelta(days=5)
    Tues = daytime.today() - timedelta(days=6)
    Mon = daytime.today() - timedelta(days=7)
    FriMod = Fri.strftime ('%m/%d/%Y') # format the date to ddmmyyyy
    ThursMod = Thurs.strftime ('%m/%d/%Y') # format the date to ddmmyyyy
    WedMod= Wedn.strftime ('%m/%d/%Y') # format the date to ddmmyyyy
    TuesMod= Tues.strftime ('%m/%d/%Y') # format the date to ddmmyyyy
    Monmod= Mon.strftime ('%m/%d/%Y') # format the date to ddmmyyyy
    M=df.loc[(df['SchoolEntryDate'] == Monmod)] 
    T= df.loc[(df['SchoolEntryDate'] == TuesMod)] 
    W=df.loc[(df['SchoolEntryDate'] == WedMod)]
    TH=df.loc[(df['SchoolEntryDate'] == ThursMod)]
    F=df.loc[(df['SchoolEntryDate'] == FriMod) ]
    days=[M,T,W,TH,F]
    week= pd.concat(days)
    week.to_excel("NewStudentInfoGrade"+str(grade)+'_'+daytime.strftime("%m%d%Y")+ ".xlsx")
    os.remove(openFile) 
    os.chdir(ROOT_DIR)

if __name__=="__main__":
    main()