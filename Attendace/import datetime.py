from datetime import datetime
from datetime import timedelta
import os
import PyPDF2
import io
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import( 
    PieChart,
    Reference
)
import pandas as pd
import math


today=datetime.now()
todayFormatted=today.strftime('%m/%d/%Y')
forFile = today.strftime(('%m%d%Y'))
df=pd.DataFrame(columns = ['Date', 'Grade', 'Classroom', 'Day Membership Possible','# of students Absent', 'Day Attendance','# of students Attending'])
if os.path.exists('classRoomAttendanceIncentive.xlsx'):
    print('The File Already Exists')
else:
    file_name = "classRoomAttendanceIncentive.xlsx"
    df.to_excel(file_name)
hn=os.environ["USERNAME"]
path1= "C:/Users/"+hn+ "/Downloads/"
#writer = pd.ExcelWriter('classRoomAttendanceIncentive.xlsx', engine='openpyxl')
#wb = load_workbook(filename='classRoomAttendanceIncentive.xlsx')
reader = pd.read_excel("classRoomAttendanceIncentive.xlsx", engine="openpyxl")
book=load_workbook("classRoomAttendanceIncentive.xlsx")


os.chdir(path1)
pdfFile= open("ClassAttendanceAudit.pdf", 'rb') 
pdfReader= PyPDF2.PdfFileReader(pdfFile)
pages = pdfReader.numPages
t=0
membership=0
attendance=0
absent=0

for x in range(pages):
    pageObj = pdfReader.getPage(x)
    lines = pageObj.extractText().splitlines()
    
    for line in lines:
       
        #print(todayFormatted)
        
        
        if line.split()[0]=="Teacher:":
            if t == 0:
               
                teacher = line.split()[1].replace(',','')
                
                t+=1 
                

        if line.split()[0]=="Section:":
           
            if t == 1:
                grade=line.split()[1]

                t+=1
                    
        if line.split()[0]=="Total":
            #print(line.split()[2])
            if line.split()[1]== "Membership:":
                membership=int(line.split()[2])
                
                
            elif line.split()[1]== "Attendance:":
                attendance=int(line.split()[2])
                #print(y)
                absent = membership-attendance 
                #"# of students Absent: "+ 
                percent= attendance/membership
                percent=percent*100
                df = df.append({'Date':todayFormatted, 'Classroom':teacher,'Grade':grade,'Day Membership Possible':membership,'# of students Absent': absent, 'Day Attendance':attendance,'# of students Attending':math.ceil(percent)},ignore_index = True)
                #print(str(absent)) 
            t=0
pdfFile.close()
os.remove("ClassAttendanceAudit.pdf")
os.chdir(os.path.realpath(os.path.join(os.path.dirname(__file__))))

if not forFile in book.sheetnames:
    book.create_sheet(forFile)
try:
    book.remove_sheet(book['Sheet1'])   
except:
    pass
with pd.ExcelWriter("classRoomAttendanceIncentive.xlsx", engine="openpyxl") as writer:
    writer.book = book
    writer.sheets.update(dict((ws.title, ws) for ws in book.worksheets))

    df.to_excel(writer,sheet_name = forFile,index=False, startrow=len(reader)-34)


book.close()
