from bs4 import BeautifulSoup

import pyodbc
import glob
import os
import csv



# change the folder location
try:
    list_of_files = glob.glob('/Users/Varun/Desktop/Consultant Prject/files/*')
    latest_file = max(list_of_files, key=os.path.getmtime)
except FileNotFoundError:
    print("File not found")


print("file" + latest_file)
a=[]



soup = BeautifulSoup(open(latest_file,"r"), "html.parser")



x=soup.find('table',attrs={'class':'c59'}).findAll('tr')[1:]
time=soup.find('td', attrs={'class': 'c5'})
date_uploaded=time.text.split(':')[1].strip()
dateuploaded=date_uploaded.rsplit(' ', 1)[0]
print(dateuploaded)
l1=[]
for s in x:

    if s.attrs ==  {'class': ['c29']}:
        if(len(l1)>0):
            l1.append(dateuploaded)
            a.append(l1)

        l1 = []
        cell=s.find_all('td')
        for cells in cell:
            l1.append(cells.text.strip())

    elif s.attrs ==  {'class': ['c44']}:
        revenues=s.find_all('td')
        for revenue in revenues:
            #print(revenue.attrs.get('class'))
            if ['c45'] ==  revenue.attrs.get('class'):
                l1.append(revenue.text.split(':')[-1].strip())

            if ['c47'] == revenue.attrs.get('class'):
             #   print(revenue.text.split(':')[1:3])
                counselor=revenue.text.split(':')[1:3]
                practice=counselor[0].split(',')[0].strip()
                counsolername=counselor[1].strip()
                l1.append(practice)
                l1.append(counsolername)
        l1.append(revenue.text.strip())
    elif s.attrs ==  {'class': ['c49']}:
        margins=s.find_all('td')
        for margin in margins:

            if ['c47'] == margin.attrs.get('class'):
                Service = margin.text.split(':')[1:4]
                serviceLine=Service[0].split(',')[0].strip()
                technologyGroup=Service[1].split(',')[0].strip()
                sector=Service[2].split(',')[0].strip()
                l1.append(serviceLine)
                l1.append(technologyGroup)
                l1.append(sector)
            else:
                YTDInfo = margin.text.split(':')[-1].strip()
                l1.append(YTDInfo)

    else:
        if s.attrs != {'class': ['c53']}:
            extracell = s.find_all('td')
            for extra in extracell:
                l1.append(extra.text.split(':')[-1].strip())







if(len(l1)>0):
    l1.append(dateuploaded)
    a.append(l1)




try:
    for columnvalues in a:
        if (len(columnvalues) != 21):
            print(columnvalues)
            print(len(columnvalues))
except:
    print (" Wrong no of parameters")




# for data in csv format

try:
    with open('/Users/Varun/Desktop/Consultant Prject/output/report.csv', 'a') as csv_file:
        writer = csv.writer(csv_file)
        for columnval in a:
            writer.writerow(columnval)
except IOError:
    print("Error reading file")








    # for data in DB;  change the server name , keep database name "demo_powerbi"
		

# cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
#                       "Server=LAPTOP-VOVBO1GN\SQLEXPRESS;"
#                       "Database=demo_powerbi;"
#                       "Trusted_Connection=yes;")
#
# cursor=cnxn.cursor()
# SQLCommand=("insert into employee"
#             "(EmployeeName,Type,RollOffBucket ,ATOPercent,ATODate ,Grade ,Email ,LastClientRate ,YTDUrve ,PracticeName ,CounselorName ,Qualifiaton ,YTDMargin,FunctionalAreas ,YTDRevenue ,ServiceLineSkills ,TechnologyGroup ,Sector ,YTDSales ,Skills ,DateUploaded )"
#           "values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
# for columnvalues in a:
#     cursor.execute(SQLCommand,columnvalues)
#     cnxn.commit()
#
# cnxn.close()


