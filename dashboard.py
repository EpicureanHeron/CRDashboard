import openpyxl

wb = openpyxl.load_workbook('expense_analysis.xlsx')

sheet = wb['Data']

print('Opening workbook')


def spendAnalysis():

    spendAnalysis = {'OOP': 0, 'Tcard': 0, "PD&M": 0}
    OOPalways = ['Per Diem Meals', 'Mileage']

    for row in range(4, sheet.max_row + 1):

        appAmount = (sheet['K' + str(row)].value)
        firmPaid = (sheet['X' + str(row)].value)
        appStatus = (sheet['U' + str(row)].value)
        expenseType = (sheet['G' + str(row)].value)

        if appStatus == 'Approved':
            if firmPaid == 'Y':
                spendAnalysis['Tcard'] += float(appAmount)
            else:
                if expenseType in OOPalways:
                    spendAnalysis['PD&M'] += float(appAmount)
                else:
                    spendAnalysis['OOP'] += float(appAmount)

    return(spendAnalysis)    
 

# ERsApprovedByRRC and ERsAffiliation should be combined, speed things up if only need to create the ERs Accounted For once
def ERsApprovedByRRC():
    
    ERsAccountedFor = []
    RRCdata = {}
    TotalApproved = 0
    affiliationData = {}


    for row in range(4, sheet.max_row + 1):

        RRC = (sheet['Y' + str(row)].value)
        ERID = (sheet['B' + str(row)].value)
        appStatus = (sheet['U' + str(row)].value)
        
        if appStatus == 'Approved':
            
            if ERID not in ERsAccountedFor:
                TotalApproved += 1
                ERsAccountedFor.append(ERID)

                RRCdata.setdefault(RRC, 0)
               

                RRCdata[RRC] += 1 
               

    RRCdata.setdefault('Total', TotalApproved)

    return(RRCdata)

def ERsAffiliation():
    ERsAccountedFor = []
    affiliationData = {}
    TotalApproved = 0
    for row in range(4, sheet.max_row + 1):
        ERID = (sheet['B' + str(row)].value)
        affl = (sheet['Z' + str(row)].value)
        if ERID not in ERsAccountedFor:
            affiliationData.setdefault(affl, 0 )
            TotalApproved += 1
            ERsAccountedFor.append(ERID)
            affiliationData[affl] +=1

    affiliationData.setdefault('Total', TotalApproved)
    return(affiliationData)


def main():
    a = ERsAffiliation()
    r = ERsApprovedByRRC()
    s = spendAnalysis()

    print(a)
    print('------------')
    print(r)
    print('------------')
    print(s)

main()
