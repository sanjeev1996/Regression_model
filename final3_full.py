import re
import docx2txt
import xlwt
import xlrd
from xlutils.copy import copy
import os
filename=[]
i_1=1
i_2=1
file_name = os.listdir("C:/Users/Amazing/Desktop/data")
for i in range(len(file_name)):
    b="C:/Users/Amazing/Desktop/data/"+file_name[i]
    doc = docx2txt.process(b)
    name=re.search(r'Patient\s? ?Name\s? ?: ?\w*.?\s?\s?\w*.?',doc)
    if name:
        name=name.group()
        name_index=re.search(r':',name)
        name_index=name_index.span()
        name1=name[name_index[1]:len(name)]
    print(name1)


    Start_date=re.search(r'Treatment Start Date ?: ?\d*.\d*.\d*',doc)
    if Start_date:
        Start_date=Start_date.group()
        Start_date_index=re.search(r':',Start_date)
        Start_date_index=Start_date_index.span()
        Start_date1=Start_date[Start_date_index[1]:len(Start_date)]
    print(Start_date1)


    Age=re.search(r'Age ?/ ?Gender ?: ?\d*/\w*',doc)
    if Age:
        Age=Age.group()
        Age_index=re.search(r':',Age)
        Age_index=Age_index.span()
        Age=Age[Age_index[1]:len(Age)]
        Age_index=re.search(r'/',Age)
        Age_index=Age_index.span()
        Age1=Age[0:Age_index[0]]
        sex=Age[Age_index[1]:len(Age)]
#    print(sex)
 #   print(Age1)

    Location=re.search(r'Location ?:? ?\w* ?',doc)
    if Location:
        Location=Location.group()
        Location_index=re.search(r':',Location)
        Location_index=Location_index.span()
        Location=Location[Location_index[1]:len(Location)].strip()
    print(Location)

    doc1=re.search(r'Problem Summary',doc)
    if doc1:
        doc1_index=doc1.span()
        doc1=doc[doc1_index[0]:len(doc)]
        doc2=re.split("\n+", doc1)
        Problem=doc2[0].lstrip('Problem Summary ?:')
#    print(Problem)

    doc3=re.search(r'Date',doc1)
    doc3_index=doc3.span()
    a=doc1[doc3_index[0]:len(doc1)]
    a=re.split("\n+", a)
    j=0
    coulumn_extension=6
    
    if (Location=="Delhi"):
        rb=xlrd.open_workbook('C:/Users/Amazing/Desktop/Delhi.xls')
    elif (Location=="Mumbai"):
        rb=xlrd.open_workbook('C:/Users/Amazing/Desktop/Mumbai.xls')
    else: rb=xlrd.open_workbook('C:/Users/Amazing/Desktop/Delhi.xls')
    wb=copy(rb)
    w_sheet=wb.get_sheet(0)
    w_sheet.write(0,0,'Patient Name')
    w_sheet.write(0,1,'Age')
    w_sheet.write(0,2,'Sex')
    w_sheet.write(0,3,'Tinnitus problem')
    w_sheet.write(0,4,'Start_date')
    w_sheet.write(0,5,'Location')
    if (Location=="Delhi"):        
        w_sheet.write(i_1,0,name1)
        w_sheet.write(i_1,1,Age1)
        w_sheet.write(i_1,2,sex)
        w_sheet.write(i_1,3,Problem)
        w_sheet.write(i_1,4,Start_date1)
        w_sheet.write(i_1,5,Location)
    elif (Location=="Mumbai"):
        w_sheet.write(i_2,0,name1)
        w_sheet.write(i_2,1,Age1)
        w_sheet.write(i_2,2,sex)
        w_sheet.write(i_2,3,Problem)
        w_sheet.write(i_2,4,Start_date1)
        w_sheet.write(i_2,5,Location)
    else:
        w_sheet.write(i_1,0,name1)
        w_sheet.write(i_1,1,Age1)
        w_sheet.write(i_1,2,sex)
        w_sheet.write(i_1,3,Problem)
        w_sheet.write(i_1,4,Start_date1)
        w_sheet.write(i_1,5,Location)


    for i in range(0,len(a),4):
    #..............Start_Date...........
        Start_Date=a[j]
        if (Start_Date=="Date"):
            j+=3
        Start_Date=a[j]
        if (Start_Date[0]=="P"):
           if (Location=="Delhi"):
               i_1+=1
               wb.save('C:/Users/Amazing/Desktop/Delhi.xls')
           elif (Location=="Mumbai"):
               i_2+=1
               wb.save('C:/Users/Amazing/Desktop/Mumbai.xls')
           else:
               i_1+=1 
               wb.save('C:/Users/Amazing/Desktop/Delhi.xls')
           break
        Start_Date=a[j]
        print("Start_Date:"+Start_Date)
        j+=1
    #.............Caller_Name...........
        Caller_Name=a[j]
  #      print('Caller_Name'+Caller_Name)
        Date_and_Caller_name=Start_Date+" , "+Caller_Name
  #      print(Date_and_Caller_name)
        j+=1
    #..............Treatment...........
        Treatment=a[j]
        Treatment=a[j].lstrip('Pattern:')
        Treatment=Treatment.split()
        Treatment=' '.join(Treatment)
   #     print('Treatment'+Treatment)
        j+=1
    #..............Response...........
        Response=a[j]
    #    print("Response"+Response)
        j+=1
        w_sheet.write(0,coulumn_extension,'Date and Caller name')
        w_sheet.write(0,coulumn_extension+1,'TreatmentNB')
        w_sheet.write(0,coulumn_extension+2,'Response')
        if (Location=="Delhi"):
            w_sheet.write(i_1,coulumn_extension,Date_and_Caller_name)
            w_sheet.write(i_1,coulumn_extension+1,Treatment)
            w_sheet.write(i_1,coulumn_extension+2,Response)
        elif (Location=="Mumbai"):
            w_sheet.write(i_2,coulumn_extension,Date_and_Caller_name)
            w_sheet.write(i_2,coulumn_extension+1,Treatment)
            w_sheet.write(i_2,coulumn_extension+2,Response)
        else:
            w_sheet.write(i_1,coulumn_extension,Date_and_Caller_name)
            w_sheet.write(i_1,coulumn_extension+1,Treatment)
            w_sheet.write(i_1,coulumn_extension+2,Response)
        coulumn_extension=coulumn_extension+3
        if (j==len(a)):
            if (Location=="Delhi"):
                i_1+=1    
                wb.save('C:/Users/Amazing/Desktop/Delhi.xls')
            elif (Location=="Mumbai"):
                i_2+=1
                wb.save('C:/Users/Amazing/Desktop/Mumbai.xls')
            else:
                i_1+=1
                wb.save('C:/Users/Amazing/Desktop/Delhi.xls')                
            break
