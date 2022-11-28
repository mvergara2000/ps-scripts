import json
import selenium
import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
def main():
    chrome_options= webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=chrome")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__)))
 
    if os.path.exists('pslogin.json'):
        print("Logging In...")
    else:
        getLogin()
    studentID= input("Enter student ID (ex: 483729,874932,727265): ")
    f= open("pslogin.json")
    login= json.load(f)
   
    driver = webdriver.Chrome(options=chrome_options)
    hn=os.environ["USERNAME"]
    path1= "C:/Users/"+hn+ "/Downloads/"
    logIn(driver,login)
    findStudents(driver, studentID)
    export(driver)
    convertToCSV(path1,ROOT_DIR)

def getLogin():
   
    username=input('Enter Username: ')
    password=input('Enter Password: ')
    url = input('Enter URL: ')

    dictionary = {
        "username": username,
        "password": password,
        "url": url,
        
    }
    
    json_object = json.dumps(dictionary, indent=4)
    with open("pslogin.json", "w") as outfile:
     outfile.write(json_object)  
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
def findStudents(driver, studentID):
    
   
    #students = ""
    x=0
    url = driver.current_url
    driver.get(url)
    while True:
        
        time.sleep(3)
        findStudent= driver.find_element("id", "studentSearchInput")
        findStudent.clear()
        time.sleep(3)
        
        id = studentID.split(",")[x]
       
        x+=1  
        print("Adding " + id)
        
        findStudent.send_keys(id)
        time.sleep(3)
        findStudent=driver.find_element("id","addRemoveStudent0")
        time.sleep(3)
        findStudent.click()
        try:
           studentID.split(",")[x]
        except:    
            break
        url = driver.current_url
    
    
    time.sleep(5)
  
def export(driver):

    time.sleep(5)
    functions=driver.find_element("id","selectFunctionDropdownButtonStudent")
    functions.click()
    functions=driver.find_element("id","lnk_studentsQuickExport")
    functions.click()
    time.sleep(5)
    print("Exporting...")
    url = driver.current_url
    driver.get(url)
    export=driver.find_element("id","tt")
    export.clear()#clear field
    time.sleep(5)
    #modify with necessary data flags
    export.send_keys("SchoolID\nGrade_Level\nLast_Name\nFirst_Name\nLast_Name\nHome_Room\nStudent_Number\nStudent_Web_ID\n Student_Web_password\nU_LD_Account_Management.Computer_ID\nU_LD_Account_Management.Computer_Password\nU_LD_Account_Management.Student_Email\nEnroll_Status\nU_StudentsUserFields.Covid_Attendance\nU_StudentsUserFields.chromebook_inventory\nU_StudentsUserFields.chromebook_serial\n")
    export=driver.find_element("id","btnSubmit")
    export.click()
    time.sleep(10)
    
def convertToCSV(path1, ROOT):
    os.chdir(path1)
    files = filter(os.path.isfile, os.listdir(path1))
    files = [os.path.join(path1, f) for f in files] # add path to each file
    files.sort(key=lambda x: -os.stat(x).st_mtime)
    openFile=files[0]
    #checks if file exists
    if os.path.exists(openFile):
        print('The file exists')
    else:
     print('The specified file does NOT exist')
    os.chdir(ROOT)
    df = pd.read_csv(openFile ,sep='\t', lineterminator='\r')
    CSVFile="DYMO_LABEL_EXPORT.csv"
    df.to_csv(CSVFile, index = False)
   # read_file = pd.read_excel(xlsxFile, sheet_name="Sheet1")
    pd.set_option('display.max_columns', None)
   
    os.remove(openFile) 
    
if __name__=="__main__":
    main()