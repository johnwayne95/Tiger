import gspread
from oauth2client.service_account import ServiceAccountCredentials

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

import unittest, time, re
import datetime
import calendar
import glob
import fnmatch
import os
import sys
import csv
import math
import Techs
import Units

#IMPORTANT GLOBAL VARIABLES
#DATES
today = (datetime.datetime.today()).date()
now = (today - datetime.timedelta(1)) #THIS IS ACTUALLY YESTERDAY, BUT THAT'S WHAT WE USE
beginningofweek = (now - datetime.timedelta(days=now.weekday()))

lastweekmon = (today - datetime.timedelta(days=now.weekday(), weeks=1))
lastweeksun = (lastweekmon + datetime.timedelta(days=6))

month = now.replace(day=1)
endofmonthday = calendar.monthrange(now.year, now.month)
endofmonth = now.replace(day=endofmonthday[1])

quarternumber = int((now.month-1) / 3 + 1)
quarter = (datetime.datetime(now.year, 3 * quarternumber - 2, 1)).date()
qmonth2 = quarter.replace(month = quarter.month+1)
qmonth3 = quarter.replace(month = quarter.month+2)
endofquarterday = calendar.monthrange(qmonth3.year, qmonth3.month)
endofquarter = qmonth3.replace(day=endofquarterday[1])

year = now.replace(month=1, day=1)
endofyear = now.replace(month=12, day=31)

springfieldtotal = 0.0

#TECHS
ryanm = Techs.Tech("RYAN M", "10")
josh = Techs.Tech("JOSH", "12")
alan = Techs.Tech("ALAN", "14")
logand = Techs.Tech("LOGAN D", "16")
derek = Techs.Tech("DEREK P", "18")

logang = Techs.Tech("LOGAN G", "22")
bobb = Techs.Tech("BOB B", "24")

erikf = Techs.Tech("ERIK F", "30")
amandaj = Techs.Tech("AMANDA J", "32")
jimmie = Techs.Tech("JIMMIE", "34")

samk = Techs.Tech("SAM K", "40")
rickb = Techs.Tech("RICK B", "42")
joez = Techs.Tech("JOE Z", "44")
cierrab = Techs.Tech("CIERRA", "46")

#ARRAY OF TECHS
techs = [ryanm, josh, alan, logand, derek, logang, bobb, erikf, amandaj, jimmie, samk, rickb, joez, cierrab]


#BUSINESS UNITS
plumbing = Units.Unit("Plumbing", "5", "6", "15")
electric = Units.Unit("Electric", "9", "28", "37")
hvac = Units.Unit("HVAC", "13", "17", "26")

#ARRAY OF BUSINESS UNITS
units = [plumbing, electric, hvac]

#CSR ARRAYS
#NAME[TAG, REV.ROW#, CONV.ROW#, JOBS, OPPS, SOLD, REVENUE]
#    [ 0      1          2       3     4     5       6   ]
#                                            ^SOLD is currently not in use, just in here for when servicetitan (hopefully) updates
justine = ["JN", "4", "20", 0, 0, 0, 0.0]
lynora = ["LB", "5", "21", 0, 0, 0, 0.0]
nicole = ["NW", "2", "22", 0, 0, 0, 0.0]
teresa = ["TJ", "3", "23", 0, 0, 0, 0.0]

elias = ["ET", "N/A", "25", 0, 0, 0, 0.0]
lisa = ["LV", "N/A", "26", 0, 0, 0, 0.0]
marcie = ["MH", "6", "27", 0, 0, 0, 0.0]

shelley = ["ST", "7", "N/A", 0, 0, 0, 0.0]
ben = ["BM", "8", "25", 0, 0, 0, 0.0]

#ARRAY OF CSR ARRAYS
csrs = [justine, lynora, nicole, teresa, shelley, elias, lisa, marcie, ben]

#CLEAR DOWNLOAD DIRECTORY OF ALL .CSV FILES
def cleardir():
    dirPath = "C:/Users/Administrator/Downloads/"
    fileList = os.listdir(dirPath)
    for fileName in fileList:
        if fileName.endswith(".csv"):
            os.remove(dirPath+"/"+fileName)
    

#GRAB FROM SERVICETITAN
def get_reports():
    cdriver = os.path.join(sys.path[0], 'chromedriver.exe')
    driver = webdriver.Chrome(cdriver)
    driver.get("https://go.servicetitan.com/#/Report")
    driver.find_element_by_id("username").click()
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("username").send_keys("jwayne")
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("password").send_keys("082718O")
    driver.find_element_by_id("password").send_keys(Keys.ENTER)
    time.sleep(5)
    driver.get("https://go.servicetitan.com/#/Search")
    time.sleep(1)
    driver.find_element_by_name("StartFrom").click()
    driver.find_element_by_name("StartFrom").clear()
    driver.find_element_by_name("StartFrom").send_keys("1/1/2018")
    driver.find_element_by_name("StartTo").click()
    driver.find_element_by_name("StartTo").clear()
    driver.find_element_by_name("StartTo").send_keys(endofmonth.strftime("%m/%d/%Y"))
    driver.find_element_by_xpath("(.//*[normalize-space(text()) and normalize-space(.)='What do you want to search for?'])[1]/following::i[1]").click()
    driver.find_element_by_xpath("(.//*[normalize-space(text()) and normalize-space(.)='What do you want to search for?'])[1]/following::button[2]").click()
    driver.find_element_by_link_text("Export to Comma separated (*.csv)").click()
    time.sleep(45)
    driver.close()

#GET ESTIMATES (FOR TOTALS FOR SOME REASON)
#THIS IS NOT CURRENTLY USED THANK GOD
def get_estimates():
    driver = webdriver.Chrome()
    driver.get("https://go.servicetitan.com/#/Report")
    driver.find_element_by_id("username").click()
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("username").send_keys("jwayne")
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("password").send_keys("082718O")
    driver.find_element_by_id("password").send_keys(Keys.ENTER)
    time.sleep(5)
    driver.get("https://go.servicetitan.com/#/Search")
    time.sleep(1)
    driver.find_element_by_name("RecordType").click()
    Select(driver.find_element_by_name("RecordType")).select_by_visible_text("Estimate")
    driver.find_element_by_name("RecordType").click()
    driver.find_element_by_name("StartFrom").click()
    driver.find_element_by_name("StartFrom").clear()
    driver.find_element_by_name("StartFrom").send_keys("1/1/2018")
    driver.find_element_by_name("StartTo").click()
    driver.find_element_by_name("StartTo").clear()
    driver.find_element_by_name("StartTo").send_keys(endofmonth.strftime("%m/%d/%Y"))
    driver.find_element_by_xpath("(.//*[normalize-space(text()) and normalize-space(.)='What do you want to search for?'])[1]/following::i[1]").click()
    driver.find_element_by_xpath("(.//*[normalize-space(text()) and normalize-space(.)='What do you want to search for?'])[1]/following::button[2]").click()
    driver.find_element_by_link_text("Export to Comma separated (*.csv)").click()
    time.sleep(75)
    driver.close()

#OPEN CSV AND PROCESS DATA
def csvgetter():
    global springfieldtotal
    global ryanm, logang, bobb, erikf, amandaj, jimmie, samk, rickb, joez
    global plumbing, electric, hvac
    global techs, units, csrs

    filepath = "C:/Users/Administrator/Downloads/*.csv"
    searchcsv = glob.glob(filepath)
    name = searchcsv[0]

    with open(name, encoding="utf8", newline='') as File:  
        reader = csv.DictReader(File)
        #loop through csv
        for row in reader:
            if row['Completed On'] != '':
                
                jobdate = (datetime.datetime.strptime(row['Completed On'], '%m/%d/%y %I:%M %p')).date()
                #loop through list of techs
                for x in techs:
                    
                    #LAST WEEK TOTALS
                    if(jobdate >= lastweekmon and jobdate <= lastweeksun):
                        if(x.name in row['Sold By']):
                            x.lastweek += float(row['Total'])

                    #MONTLY VALUES
                    if(jobdate >= month and jobdate <= now):
                        if(x.name in row['Sold By']):
                            if("Completed" in row['Status']):
                                x.monthtotal += float(row['Total'])
                                if("Maintenance" in row['Business Unit']):
                                    x.maintenance += 1
                                if("Service" in row['Business Unit']):
                                    x.service += 1
                        #MONTHLY CONVERTIBLE
                        if(x.name in row['Sold By']):
                            if(('No' in row['No Charge'] and 'No' in row['Recall'] and 'No' in row['Warranty']) or float(row['Total']) >= 49.0 ):
                                x.convcalls += 1
                                if("Service" in row['Business Unit']):
                                    x.servconv += 1
                                if("Maintenance" in row['Business Unit']):
                                    x.maintconv += 1
                            #MONTHLY SOLD
                            if(float(row['Total']) >= 49.0):
                                x.soldcalls += 1
                                if("Service" in row['Business Unit']):
                                    x.servsold += 1
                                if("Maintenance" in row['Business Unit']):
                                    x.maintsold += 1

                            if(row['Type'] in ('IAQ System Design', 'IAQ Work', 'SOURCE REMOVAL') and "Completed" in row['Status']):
                                x.iaqconv += 1
                                if(float(row['Total']) > 0.0):
                                    x.iaqcls +=1

                            if(row['Type'] in ('System Design', 'SYSTEM DESIGN', 'QA - INSTALL', 'SYSTEM DESIGN - MARKETED', 'QA - SYSTEM DESIGN', 'INSTALL', 'INSTALL - MARKETED') and "Completed" in row['Status']):
                                x.sdconv += 1
                                if(float(row['Total']) > 0.0):
                                    x.sdcls +=1

            if row['Start'] != '':  
                #END OF MONTH (WIP CALC)
                for x in techs:
                    jobdate = (datetime.datetime.strptime(row['Start'], '%m/%d/%y %I:%M %p')).date()
                    if(jobdate >= month and jobdate <= endofmonth ):
                        if(x.name in row['Sold By']):
                            if("Cancelled" not in row['Status'] and "Completed" not in row['Status']):
                                x.wip += float(row['Total'])

                #LOOP THROUGH LIST OF BUSINESSES
                for unit in units:
                    
                    if(unit.name in row['Business Unit']):

                        #LAST WEEK
                        if(jobdate >= lastweekmon and jobdate <= lastweeksun):
                            unit.lastweek += float(row['Total'])

                        #DAILY VALUES FOR COUNTDOWN ON SLIDES
                        if(jobdate == beginningofweek):
                            unit.mon += float(row['Total']) #MONDAY
                        if(jobdate >= beginningofweek and jobdate <= (beginningofweek + datetime.timedelta(days=1))):
                            unit.tue += float(row['Total']) #TUESDAY
                        if(jobdate >= beginningofweek and jobdate <= (beginningofweek + datetime.timedelta(days=2))):
                            unit.wed += float(row['Total']) #WEDNESDAY
                        if(jobdate >= beginningofweek and jobdate <= (beginningofweek + datetime.timedelta(days=3))):
                            unit.thu += float(row['Total']) #THURSDAY
                        if(jobdate >= beginningofweek and jobdate <= (beginningofweek + datetime.timedelta(days=4))):
                            unit.fri += float(row['Total']) #FRIDAY
                        if(jobdate >= beginningofweek and jobdate <= (beginningofweek + datetime.timedelta(days=5))):
                            unit.sat += float(row['Total']) #SATURDAY
                        if(jobdate >= beginningofweek and jobdate <=(beginningofweek + datetime.timedelta(days=6))):
                            unit.sun += float(row['Total']) #SUNDAY

                        #MONTLY VALUES
                        if(jobdate >= month and jobdate <= now):
                            unit.monthtotal += float(row['Total'])
                            if("Maintenance" in row['Business Unit']):
                                unit.maintenance += 1
                            if("Service" in row['Business Unit']):
                                unit.service += 1
                            #MONTHLY CONVERTIBLE
                            if(('No' in row['No Charge'] and 'No' in row['Recall'] and 'No' in row['Warranty']) or float(row['Total']) >= 50.0 ):
                                unit.conv += 1
                            #MONTHLY SOLD
                            if(float(row['Total']) >= 50.0):
                                unit.sold += 1

                            #WEEKLY VALUES FOR COUNTDOWN ON SLIDES
                            if(jobdate.day <= 7):
                                unit.week1 += float(row['Total']) #WEEK 1
                            if(jobdate.day <= 14):
                                unit.week2 += float(row['Total']) #WEEK 2
                            if(jobdate.day <= 21):
                                unit.week3 += float(row['Total']) #WEEK 3
                            if(jobdate.day <= 28):
                                unit.week4 += float(row['Total']) #WEEK 4
                            if(jobdate.day <= 31):
                                unit.week5 += float(row['Total']) #WEEK 5

                        #END OF MONTH (WIP CALC)
                        if(jobdate >= month and jobdate <= endofmonth ):
                            if("Cancelled" not in row['Status'] and "Completed" not in row['Status']):
                                unit.wip += float(row['Total'])

                        #QUARTERLY TOTALS
                        if(jobdate >= quarter and jobdate <= now):
                            unit.quartertotal += float(row['Total'])
                            unit.month3 += float(row['Total'])
                            if(jobdate < qmonth2):
                                unit.month1 += float(row['Total'])
                            if(jobdate < qmonth3):
                                unit.month2 += float(row['Total'])

                        #Yearly Total for Businesses
                        unit.yearlytotal += float(row['Total'])

                for x in csrs:
                    if(jobdate >= month and jobdate <= now):
                        if(x[0] in row['Tag(s)'] and row['Completed On'] != ""):
                            x[6] += float(row['Total'])
                            x[3] += 1
                            if("Potential Member" in row['Tag(s)']):
                                x[4] += 1

                #STUPID SPRINGFIEND ZIP CODES GARBAGE
                if(row['Zip'] in ('62712', '62711', '62707', '62704', '62703', '62702', '62661', '62650', '62640', '62629')):
                    springfieldtotal += float(row['Total'])

#OPEN ESTIMATES AND PROCESS DATA
#THIS IS ALSO NOT USED AT THIS TIME
def estimatesgetter():
    global springfieldtotal
    global ryanm, logang, bobb, erikf, amandaj, jimmie, samk, rickb, joez
    global plumbing, electric, hvac
    global techs, units, csrs

    filepath = "C:/Users/Administrator/Downloads/*.csv"
    searchcsv = glob.glob(filepath)
    name = searchcsv[0]

    with open(name, encoding="utf8", newline='') as File:  
        reader = csv.DictReader(File)
        #loop through csv
        for row in reader:
            if row['Created On'] != '':
                jobdate = (datetime.datetime.strptime(row['Created On'], '%m/%d/%Y')).date()
                total = float(row['Amount'])

                for tech in techs:

                    #LAST WEEK TOTALS
                    if(jobdate >= lastweekmon and jobdate <= lastweeksun):
                        if(tech.name in row['Sold By']):
                            tech.lastweek += total

                    #MONTHLY TOTALS
                    if(jobdate >= month and jobdate <= now):
                        if(tech.name in row['Sold By']):
                            tech.monthtotal += total
                            
#PUT DATA INTO GOOGLE SHEETS

secret = os.path.join(sys.path[0], 'client_secret.json')
#SoldBy Month
def soldbysheet():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    monthname = now.strftime("%B")
    yearnumber = now.strftime("%y")
    worksheetname = "SoldBy " + monthname + " " + yearnumber
    soldbysheet = client.open("Service Advisor Numbers").worksheet(worksheetname)

    soldbysheet.update_acell('A1', now.strftime("%m/%d/%Y"))

    for tech in techs:

        #INPUT LAST WEEK TOTALS FOR TECHS
        soldbysheet.update_acell('A' + tech.row, tech.lastweek)

        #INPUT MONTHLY (NOT FOR HVAC ADVISORS)
        if('ERIK F' not in tech.name and 'AMANDA J' not in tech.name and 'JIMMIE' not in tech.name):
            range_build = 'K' + tech.row + ":L" + tech.row
            cell_list = soldbysheet.range(range_build)
            cell_list[0].value = tech.maintenance #MAINTENANCE
            cell_list[1].value = tech.service #SERVICE
            soldbysheet.update_cells(cell_list)

        # #INPUT IAQ & SD (HVAC ADVISORS)
        if('ERIK F' in tech.name or 'AMANDA J' in tech.name or 'JIMMIE' in tech.name):
            range_build = 'K' + tech.row + ":L" + tech.row
            cell_list = soldbysheet.range(range_build)
            cell_list[0].value = str(tech.iaqcls) + "/" + str(tech.iaqconv) #IAQ
            cell_list[1].value = str(tech.sdcls) + "/" + str(tech.sdconv) #SD
            soldbysheet.update_cells(cell_list)
        
        range_build = 'O' + tech.row + ":T" + tech.row
        cell_list = soldbysheet.range(range_build)
        cell_list[0].value = tech.convcalls #CONVERTIBLE CALLS
        cell_list[1].value = tech.soldcalls #SOLD CALLS
        cell_list[2].value = tech.servconv #CONV SERVICE
        cell_list[3].value = tech.servsold #SOLD SERVICE
        cell_list[4].value = tech.maintconv #CONV MAINT
        cell_list[5].value = tech.maintsold #SOLD MAINT
        soldbysheet.update_cells(cell_list)

        range_build = 'AD' + tech.row + ":AE" + tech.row
        cell_list = soldbysheet.range(range_build)
        cell_list[0].value = tech.monthtotal #MONTHLY TOTAL
        cell_list[1].value = tech.wip #MONTHLY WIP
        soldbysheet.update_cells(cell_list)

#SOAC sheet
def soacsheet():
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    monthname = now.strftime("%B")
    yearnumber = now.strftime("%y")
    worksheetname = "SOAC " + monthname + " " + yearnumber
    soacsheet = client.open("Service Advisor Numbers").worksheet(worksheetname)

    soacsheet.update_acell('A1', "Summary of all Companies as of: " + now.strftime("%m/%d/%Y"))

    for unit in units:

        #INPUT LAST WEEK TOTAL
        soacsheet.update_acell('A' + unit.soacrow, unit.lastweek)

        range_build = 'K' + unit.soacrow + ":P" + unit.soacrow
        cell_list = soacsheet.range(range_build)
        cell_list[0].value = unit.maintenance #MONTH MAINT
        cell_list[1].value = unit.service #MONTH SERV
        cell_list[2].value = 0 #WEEK (NOT USED)
        cell_list[3].value = 0 #WEEK (NOT USED)
        cell_list[4].value = unit.conv #MONTH CALLS RAN
        cell_list[5].value = unit.sold #MONTH CALLS CLOSED
        soacsheet.update_cells(cell_list)

        range_build = 'X' + unit.soacrow + ":Z" + unit.soacrow
        cell_list = soacsheet.range(range_build)
        cell_list[0].value = unit.monthtotal #MONTHLY TOTAL
        cell_list[1].value = unit.yearlytotal #YEARLY TOTAL
        cell_list[2].value = unit.wip #WIP
        soacsheet.update_cells(cell_list)

#CSR REVENUE
def revenuesheet():
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    worksheetname = "Revenue"
    revenuesheet = client.open("Club Sales Stats").worksheet(worksheetname)

    range_build = 'C2:C8'
    cell_list = revenuesheet.range(range_build)
    cell_list[0].value = nicole[6]
    cell_list[1].value = teresa[6]
    cell_list[2].value = justine[6]
    cell_list[3].value = lynora[6]
    cell_list[4].value = marcie[6]
    cell_list[5].value = shelley[6]
    cell_list[6].value = ben[6]
    revenuesheet.update_cells(cell_list)

#CSR CONVERSION
def conversionsheet():
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    worksheetname = "Club Conversion"
    conversionsheet = client.open("Club Sales Stats").worksheet(worksheetname)

    range_build = 'B17:C25'
    cell_list = conversionsheet.range(range_build)

    #JOBS
    cell_list[0].value = justine[3]
    cell_list[2].value = lynora[3]
    cell_list[4].value = nicole[3]
    cell_list[6].value = teresa[3]
    cell_list[10].value = elias[3]
    cell_list[12].value = lisa[3]
    cell_list[14].value = marcie[3]
    cell_list[16].value = ben[3]

    #OPPORTUNITIES 
    cell_list[1].value = justine[4]
    cell_list[3].value = lynora[4]
    cell_list[5].value = nicole[4]
    cell_list[7].value = teresa[4]
    cell_list[11].value = elias[4]
    cell_list[13].value = lisa[4]
    cell_list[15].value = marcie[4]
    cell_list[17].value = ben[4]

    conversionsheet.update_cells(cell_list)

#SHEET THAT UPDATES SLIDES
def slidessheet():

    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    quarterlygoal = plumbing.quartergoal + electric.quartergoal + hvac.quartergoal
    quarterlytotal = plumbing.quartertotal + electric.quartertotal + hvac.quartertotal
    quarterlyremain = quarterlygoal - quarterlytotal
    quarterpercent = float(quarterlytotal / quarterlygoal)
    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    worksheetname = "Slides"
    slidessheet = client.open("Master Reports Sheet").worksheet(worksheetname)

    range_build = 'A1:A4'
    cell_list = slidessheet.range(range_build)
    cell_list[0].value = "Quarter " + str(quarternumber) + " Goal: ${:,.2f}".format(quarterlygoal)
    cell_list[1].value = "Amount left as of " + now.strftime("%B") + " " + str(now.day) + ":"
    cell_list[2].value = "${:,.2f}".format(quarterlyremain)
    cell_list[3].value = "We are at " + "{:.1%}".format(quarterpercent) + " of Q" + str(quarternumber) + " Goal" 
    slidessheet.update_cells(cell_list)

    for unit in units:

        range_build = 'A'+ unit.slidesrow +':H'+ unit.slidesrowend
        cell_list = slidessheet.range(range_build)

        cell_list[0].value = unit.name + " " + now.strftime("%Y") + " Goals"
        cell_list[8].value = "As of " + now.strftime("%A, %b ") + str(now.day) + now.strftime(" ,%Y")

        #THESE ARE THE BASE GOAL NUMBERS
        weekgoal = (unit.monthgoal)/4.3
        cell_list[17].value = weekgoal #WEEK GOAL
        cell_list[19].value = unit.monthgoal #MONTH GOAL
        cell_list[20].value = "Quarter " + str(quarternumber) #QUARTER NUMBER
        cell_list[21].value = unit.quartergoal #QUARTER GOAL
        cell_list[23].value = unit.yearlygoal #YEAR GOAL

        #DAILY REMAINDERS
        mondayremain = "${:,.2f}".format(weekgoal - unit.mon)
        cell_list[25].value = mondayremain #MONDAY
        cell_list[33].value = "" #TUESDAY
        cell_list[41].value = "" #WEDNESDAY
        cell_list[49].value = "" #THURSDAY
        cell_list[57].value = "" #FRIDAY
        cell_list[65].value = "" #SATURDAY
        cell_list[73].value = "" #SUNDAY

        if(now.weekday() >= 1):
            tuesdayremain = "${:,.2f}".format(weekgoal - unit.tue)
            cell_list[33].value = tuesdayremain #TUESDAY
        if(now.weekday() >= 2):
            wednesdayremain = "${:,.2f}".format(weekgoal - unit.wed)
            cell_list[41].value = wednesdayremain #WEDNESDAY
        if(now.weekday() >= 3):
            thursdayremain = "${:,.2f}".format(weekgoal - unit.thu)
            cell_list[49].value = thursdayremain #THURSDAY
        if(now.weekday() >= 4):
            fridayremain = "${:,.2f}".format(weekgoal - unit.fri)
            cell_list[57].value = fridayremain #FRIDAY
        if(now.weekday() >= 5):
            saturdayremain = "${:,.2f}".format(weekgoal - unit.sat)
            cell_list[65].value = saturdayremain #SATURDAY
        if(now.weekday() >= 6):
            sundayremain = "${:,.2f}".format(weekgoal - unit.sun)
            cell_list[73].value = sundayremain #SUNDAY

        #WEEKLY REMAINDERS
        week1remain = "${:,.2f}".format(unit.monthgoal - unit.week1)
        cell_list[27].value = week1remain #WEEK 1
        cell_list[35].value = "" #WEEK 2
        cell_list[43].value = "" #WEEK 3
        cell_list[51].value = "" #WEEK 4
        cell_list[59].value = "" #WEEK 5

        if(now.day >= 7):
            week2remain = "${:,.2f}".format(unit.monthgoal - unit.week2)
            cell_list[35].value = week2remain #WEEK 2
        if(now.day >= 14):
            week3remain = "${:,.2f}".format(unit.monthgoal - unit.week3)
            cell_list[43].value = week3remain #WEEK 3
        if(now.day >= 21):
            week4remain = "${:,.2f}".format(unit.monthgoal - unit.week4)
            cell_list[51].value = week4remain #WEEK 4
        if(now.day >= 28):
            week5remain = "${:,.2f}".format(unit.monthgoal - unit.week5)
            cell_list[59].value = week5remain #WEEK 5

        #MONTHLY REMAINDERS
        month1remain = "${:,.2f}".format(unit.quartergoal - unit.month1)
        cell_list[29].value = month1remain #MONTH 1
        cell_list[37].value = "" #MONTH 2
        cell_list[45].value = "" #MONTH 3

        if(now >= qmonth2):
            month2remain = "${:,.2f}".format(unit.quartergoal - unit.month2)
            cell_list[37].value = month2remain #MONTH 2
        if(now >= qmonth3):
            month3remain = "${:,.2f}".format(unit.quartergoal - unit.month3)
            cell_list[45].value = month3remain #MONTH 3

        #YEAR REMAINDER
        yearremain = "${:,.2f}".format(unit.yearlygoal - unit.yearlytotal)
        cell_list[31].value = yearremain #YEAR

        slidessheet.update_cells(cell_list)

#WEEKLY MGMT SHEET
def weeklymgmtsheet():
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    monthname = today.strftime("%B")
    worksheetname = monthname + " SMN"
    weeklymgmtsheet = client.open("Weekly Management Meeting Agenda").worksheet(worksheetname)

    remainingmonth = (endofmonth - now).days
    remainingquarter = (endofquarter - now).days
    remainingyear = (endofyear - now).days

    range_build = 'B2:M4'
    cell_list = weeklymgmtsheet.range(range_build)
    #MONTHLY TOTAL
    cell_list[0].value = plumbing.monthtotal
    cell_list[1].value = hvac.monthtotal
    cell_list[2].value = electric.monthtotal
    #QUARTER TOTAL
    cell_list[4].value = plumbing.quartertotal
    cell_list[5].value = hvac.quartertotal
    cell_list[6].value = electric.quartertotal
    #YEARLY TOTAL
    cell_list[8].value = plumbing.yearlytotal
    cell_list[9].value = hvac.yearlytotal
    cell_list[10].value = electric.yearlytotal

    cell_list[11].value = springfieldtotal
    #MONTHLY REMAINDER
    cell_list[12].value = plumbing.monthgoal - plumbing.monthtotal
    cell_list[13].value = hvac.monthgoal - hvac.monthtotal
    cell_list[14].value = electric.monthgoal - electric.monthtotal
    #QUARTERLY REMAINDER
    cell_list[16].value = plumbing.quartergoal - plumbing.quartertotal
    cell_list[17].value = hvac.quartergoal - hvac.quartertotal
    cell_list[18].value = electric.quartergoal - electric.quartertotal
    #YEARLY REMAINDER
    cell_list[20].value = plumbing.yearlygoal - plumbing.yearlytotal
    cell_list[21].value = hvac.yearlygoal - hvac.yearlytotal
    cell_list[22].value = electric.yearlygoal - electric.yearlytotal

    #MONTH PER DAY
    cell_list[24].value = (plumbing.monthgoal - plumbing.monthtotal)/remainingmonth
    cell_list[25].value = (hvac.monthgoal - hvac.monthtotal)/remainingmonth
    cell_list[26].value = (electric.monthgoal - electric.monthtotal)/remainingmonth

    #QUARTER PER DAY
    cell_list[28].value = (plumbing.quartergoal - plumbing.quartertotal)/remainingquarter
    cell_list[29].value = (hvac.quartergoal - hvac.quartertotal)/remainingquarter
    cell_list[30].value = (electric.quartergoal - electric.quartertotal)/remainingquarter

    #YEAR PER DAY
    cell_list[32].value = (plumbing.yearlygoal - plumbing.yearlytotal)/remainingyear
    cell_list[33].value = (hvac.yearlygoal - hvac.yearlytotal)/remainingyear
    cell_list[34].value = (electric.yearlygoal - electric.yearlytotal)/remainingyear
    weeklymgmtsheet.update_cells(cell_list)

    #YEARLYTOTAL
    weeklymgmtsheet.update_acell('N17', (plumbing.yearlytotal + electric.yearlytotal + hvac.yearlytotal))

    #MONTHLY PER DAY
    maintmonthplumb = float(weeklymgmtsheet.acell('B6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('B7', maintmonthplumb/remainingmonth)

    maintmonthhvac = float(weeklymgmtsheet.acell('C6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('C7', maintmonthhvac/remainingmonth)

    maintmonthelec = float(weeklymgmtsheet.acell('D6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('D7', maintmonthelec/remainingmonth)

    servmonthplumb = float(weeklymgmtsheet.acell('B12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('B13', servmonthplumb/remainingmonth)

    servmonthhvac = float(weeklymgmtsheet.acell('C12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('C13', servmonthhvac/remainingmonth)

    servmonthelec = float(weeklymgmtsheet.acell('D12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('D13', servmonthelec/remainingmonth)

    #QUARTERLY PER DAY
    maintquarterplumb = float(weeklymgmtsheet.acell('F6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('F7', maintquarterplumb/remainingquarter)

    maintquarterhvac = float(weeklymgmtsheet.acell('G6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('G7', maintquarterhvac/remainingquarter)

    maintquarterelec = float(weeklymgmtsheet.acell('H6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('H7', maintquarterelec/remainingquarter)

    servquarterplumb = float(weeklymgmtsheet.acell('F12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('F13', servquarterplumb/remainingquarter)

    servquarterhvac = float(weeklymgmtsheet.acell('G12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('G13', servquarterhvac/remainingquarter)

    servquarterelec = float(weeklymgmtsheet.acell('H12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('H13', servquarterelec/remainingquarter)

    #YEARLY PER DAY
    maintyearplumb = float(weeklymgmtsheet.acell('J6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('J7', maintyearplumb/remainingyear)

    maintyearhvac = float(weeklymgmtsheet.acell('K6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('K7', maintyearhvac/remainingyear)

    maintyearelec = float(weeklymgmtsheet.acell('L6').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('L7', maintyearelec/remainingyear)

    servyearplumb = float(weeklymgmtsheet.acell('J12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('J13', servyearplumb/remainingyear)

    servyearhvac = float(weeklymgmtsheet.acell('K12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('K13', servyearhvac/remainingyear)

    servyearelec = float(weeklymgmtsheet.acell('L12').value[1: ].replace(",", "").replace("$", ""))
    weeklymgmtsheet.update_acell('L13', servyearelec/remainingyear)

#GRABS MONTHLY, QUARTERLY, AND YEARLY GOALS FROM GOALS.TXT
def goalstxt():
    goals = os.path.join(sys.path[0], 'goals.txt')
    file = open(goals)
    file.readline()
    file.readline()
    plumbing.monthgoal = float(file.readline())
    file.readline() 
    plumbing.quartergoal = float(file.readline())
    file.readline()
    plumbing.yearlygoal = float(file.readline())

    file.readline()
    electric.monthgoal = float(file.readline())
    file.readline()
    electric.quartergoal = float(file.readline())
    file.readline()
    electric.yearlygoal = float(file.readline())

    file.readline()
    hvac.monthgoal = float(file.readline())
    file.readline()
    hvac.quartergoal = float(file.readline())
    file.readline()
    hvac.yearlygoal = float(file.readline())

    file.close()

def dupsheet():
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    monthname = now.strftime("%B")
    yearnumber = now.strftime("%y")
    worksheetname = "SoldBy " + monthname + " " + yearnumber
    soldbysheet = client.open("Service Advisor Numbers").worksheet(worksheetname)


    print(gspread.__version__)