import openpyxl, sys

wb = openpyxl.load_workbook('expense_analysis.xlsx')

sheet = wb['Data']

print('Opening workbook')

pilotRRCs = ['ATHLX', 'AUXSV', 'AVPFN', 'CPPMX', 'FMXXX', 'OHRXX', 'OITXX', 'PSRXX', 'PUBSF', 'SUFIN', 'SVPFO', 'UHLSF', 'UMDXX', 'USERV']
nonPilotRRCs = ['GPSTR', 'MNEXT', 'UMCXX', 'UMMXX', 'CCAPS', 'NURSG', 'OGCXX', 'UMRXX', 'PRESD', 'AUDIT', 'CSOMX', 'UEDUC', 'EQDIV', 'URELX', 'AESXX', 'CEHDX', 'RSRCH', 'GRADX', 'AHCSH', 'AHSCI', 'HLSCI', 'CLAXX', 'CSENG', 'DESGN', 'LAWXX', 'LIBRX', 'STDAF', 'PUBHL', 'DENTX', 'HHHXX', 'PHARM', 'CFANS', 'AAPRV', 'MEDXX', 'VETMD', 'CBSXX', 'RGNTS']

def RRClist():
    if len(sys.argv) > 1:
    
        if sys.argv[1] == 'non-pilot':
            includeList = nonPilotRRCs
            return(includeList)

    else:
        includeList = nonPilotRRCs + pilotRRCs
        return(includeList)
        
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

    print(a)
    print('------------')
    print(r)
    print('------------')
    print(s)

main()
