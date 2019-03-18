import openpyxl
import sys
from datetime import datetime

import csv


wb = openpyxl.load_workbook('expense_analysis.xlsx')


# change this to select the only sheet, which I have previously renamed as 'data' but it is usually some time stamp, think there is a notion of 'active' to select
sheet = wb['Data']

print('Opening workbook')

pilotRRCs = ['ATHLX', 'AUXSV', 'AVPFN', 'CPPMX', 'FMXXX', 'OHRXX', 'OITXX', 'PSRXX', 'PUBSF', 'SUFIN', 'SVPFO', 'UHLSF', 'UMDXX', 'USERV']
nonPilotRRCs = ['GPSTR', 'MNEXT', 'UMCXX', 'UMMXX', 'CCAPS', 'NURSG', 'OGCXX', 'UMRXX', 'PRESD', 'AUDIT', 'CSOMX', 'UEDUC', 'EQDIV', 'URELX', 'AESXX', 'CEHDX', 'RSRCH', 'GRADX', 'AHCSH', 'AHSCI', 'HLSCI', 'CLAXX', 'CSENG', 'DESGN', 'LAWXX', 'LIBRX', 'STDAF', 'PUBHL', 'DENTX', 'HHHXX', 'PHARM', 'CFANS', 'AAPRV', 'MEDXX', 'VETMD', 'CBSXX', 'RGNTS']

def RRClist():
    if len(sys.argv) > 1:
        includeList = []
        if sys.argv[1] == 'non-pilot':
            includeList = nonPilotRRCs
            return(includeList)
        else:
            for x in sys.argv[1:]:
                if x in pilotRRCs or x in nonPilotRRCs:

                    includeList.append(x)

            return(includeList)

    else:
        includeList = nonPilotRRCs + pilotRRCs
        return(includeList)

def approvalTime(includeList):

    i = includeList
    approvalTimes = []
    exportStatusList = ['Exported/Not Paid', 'Exported/Paid', 'Exported/Partially Paid']
    ERsAccountedFor = []
    lessThanThreeDaysCounter = 0


    for row in range(4, sheet.max_row + 1):


        ERID = (sheet['B' + str(row)].value)
        submittedDate = (sheet['C' + str(row)].value)
        exportedDate = (sheet['W' + str(row)].value)
        RRC = (sheet['Y' + str(row)].value)
        exportStatus = (sheet['V' + str(row)].value)


        if RRC in i:
            if exportStatus in exportStatusList:
                if  ERID not in ERsAccountedFor:
                    ERsAccountedFor.append(ERID)
                    try:
                        daysPassed = (exportedDate - submittedDate).days

                        approvalTimes.append(daysPassed)
    
                    except:
                        print('failed %s' % ERID)
                    if daysPassed <= 3:
                                lessThanThreeDaysCounter += 1

    average = sum(approvalTimes)/len(approvalTimes)

    lessThanThreePercent = lessThanThreeDaysCounter/len(approvalTimes)
    approvalTimeResults = {'Average': average,
                           'LessThan3Days': lessThanThreePercent}

    return(approvalTimeResults)

def spendAnalysis(includeList):

    i = includeList
            

    spendAnalysis = {'OOP': 0, 'Tcard': 0, "PD&M": 0}
    OOPalways = ['Per Diem Meals', 'Mileage']

    for row in range(4, sheet.max_row + 1):

        appAmount = (sheet['K' + str(row)].value)
        firmPaid = (sheet['X' + str(row)].value)
        RRC = (sheet['Y' + str(row)].value)
        appStatus = (sheet['U' + str(row)].value)
        expenseType = (sheet['G' + str(row)].value)

        if appStatus == 'Approved':
            if RRC in i:
            
                if firmPaid == 'Y':
                    spendAnalysis['Tcard'] += float(appAmount)
                else:
                    if expenseType in OOPalways:
                        spendAnalysis['PD&M'] += float(appAmount)
                    else:
                        spendAnalysis['OOP'] += float(appAmount)

    return(spendAnalysis)    
 

# ERsApprovedByRRC and ERsAffiliation should be combined, speed things up if only need to create the ERs Accounted For once
def ERsApprovedByRRC(includeList):

    i = includeList
    ERsAccountedFor = []
    RRCdata = {}
    TotalApproved = 0
    affiliationData = {}


    for row in range(4, sheet.max_row + 1):
        

        RRC = (sheet['Y' + str(row)].value)
        ERID = (sheet['B' + str(row)].value)
        appStatus = (sheet['U' + str(row)].value)

        if RRC in i:
        
            if appStatus == 'Approved':
                
                if ERID not in ERsAccountedFor:
                    TotalApproved += 1
                    ERsAccountedFor.append(ERID)

                    RRCdata.setdefault(RRC, 0)
                

                    RRCdata[RRC] += 1 
               

    RRCdata.setdefault('Total', TotalApproved)

    return(RRCdata)

def ERsAffiliation(includeList):
    i = includeList
    ERsAccountedFor = []
    affiliationData = {}
    
    TotalApproved = 0
    for row in range(4, sheet.max_row + 1):
        ERID = (sheet['B' + str(row)].value)
        affl = (sheet['Z' + str(row)].value)
        RRC = (sheet['Y' + str(row)].value)
        if RRC in i:
            if ERID not in ERsAccountedFor:
                affiliationData.setdefault(affl, 0 )
                TotalApproved += 1
                ERsAccountedFor.append(ERID)
                affiliationData[affl] +=1

    affiliationData.setdefault('Total', TotalApproved)
    return(affiliationData)




def main():

    includeList = RRClist()

    a = ERsAffiliation(includeList)
    r = ERsApprovedByRRC(includeList)
    s = spendAnalysis(includeList)
    d = approvalTime(includeList)

    name = str(datetime.now().date()) + ".csv"

    with open(name, 'w') as f:  # Just use 'w' mode in 3.x
        
    
        w = csv.writer(f)

        for row in a.items():
            w.writerow(row)
        
        for row in r.items():
            w.writerow(row)

        for row in s.items():
            w.writerow(row)

        for row in d.items():
            w.writerow(row)

main()

