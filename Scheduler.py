import Reports
import time
import gui
import datetime
import os
import sys

def main():
    now = (datetime.datetime.today()).date()
    gui.setup()

    #GRABS THE GOALS FOR THE MONTH, QUARTER, AND YEAR FROM OUR EXCEL FILE
    Reports.goalstxt()

    #GRABS TECHNICIAN LIST FROM OUR EXCEL FILE OF TECHS WE WANT ON THE SHEET
    Reports.techstxt()

    Reports.csrstxt()

    gui.inputgoals()

    gui.print("Starting Daily Reports...")
    time.sleep(2)

    gui.print("Clearing Directory...")
    #REPORTS- CLEAR DIRECTORY OF ALL .CSV FILES
    Reports.cleardir()

    gui.print("Getting Data File From ServiceTitan...")
    #REPORTS- GO TO SERVICETITAN AND GRAB NEW CSV FILE
    Reports.get_reports()

    gui.print("Processing Data File...")
    #REPORTS- ANALYZES AND SORTS DATA FROM CSV FILE
    Reports.csvgetter()

    gui.print("Clearing Directory Again...")
    #REPORTS- CLEAR DIRECTORY OF ALL .CSV FILES
    Reports.cleardir()

    gui.print("Getting CSR Memberships File From ServiceTitan...")
    #REPORTS- GO TO SERVICETITAN AND GRAB NEW CSV FILE
    Reports.get_memberships()

    gui.print("Processing Memberships File...")
    #REPORTS- ANALYZES AND SORTS DATA FROM CSV FILE
    Reports.membershipsgetter()

    gui.print("Exporting Data to SoldBy Sheet...")
    #REPORTS SOLDBY SHEET
    Reports.soldbysheet()

    gui.print("Exporting Data to SOAC Sheet...")
    #REPORTS SOAC SHEET
    Reports.soacsheet()

    #^ doing these two back-to-back is hitting 95 write requests, will have to wait 100 seconds after them
    gui.print("Waiting 100 seconds to Bypass Google API limit...")

    trey1 = os.path.join(sys.path[0], 'trey_1.ico')
    trey2 = os.path.join(sys.path[0], 'trey_2.ico')

    for i in range(100,0,-1):
        time.sleep(0.25)
        if(i%10 == 0):
            gui.loadingprint(str(i) + " Seconds Left")
        gui.dots()
        gui.changeicon(trey2)
        time.sleep(0.25)
        gui.changeicon(trey1)
        time.sleep(0.25)
        gui.changeicon(trey2)
        time.sleep(0.25)
        gui.changeicon(trey1)

    gui.print("Exporting Data to Slides Sheet...")
    #REPORTS- SLIDES SHEET
    Reports.slidessheet()

    gui.print("Exporting Data to CSR Revenue Sheet...")
    #REPORTS- REVENUE SHEET
    Reports.revenuesheet()

    gui.print("Exporting Data to CSR Club Conversion Sheet...")
    #REPORTS- CONVERSION RATE SHEET
    Reports.conversionsheet()

    if(now.weekday() == 0):
        gui.print("It's Monday! Exporting Data to Weekly Management Meeting Sheet...")
        #REPORTS- WEEKLY MGMT MEETING SHEET
        Reports.weeklymgmtsheet()
    gui.print("ALL DONE! You may now close this window.")
    gui.update() #KEEPS WINDOW OPEN UNTIL USER CLOSES IT

main()