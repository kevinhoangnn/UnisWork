import time
from pandas import read_excel
import BNPauto

'''
Billing Items Here:
'''


try:
    start_time = time.time()

    billingAcc = ''
    accName = ''
    facility = ''
    startPeriod = input('Input Start Date (MM/DD/YY): ')
    endPeriod = input('Input End Date (MM/DD/YY): ')
    user = input('Input Username for Computer: ')
    billCycle = 'Bimonthly'

    #Input file path for activity report
    activityLoc = input('Input File Path for Wise Activity Report (Remove ""): ')
    activityLoc.replace('/', '//')

    facility, invoicePath = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, user)
    facility = facility.lower()
    facility = facility.replace(' ', '')
    facility = facility.capitalize()

    reportLoc = BNPauto.invoiceToReport(user, accName, facility, billCycle, invoicePath)

    billingItemDict = {'Create Dictionary for Comparing Item Name and Description. '}
    itemStepsList = ['Put int the function Names']

    report = read_excel(reportLoc)
    description = report['Description'].tolist()
    itemName = report['ItemName'].tolist()
    itemList = []

    for count, name in enumerate(description):
        tempList = [itemName[count], name]
        itemList.append(tempList)

    for index, item in enumerate(itemList):
        
        try:
            itemStep = itemStepsList[billingItemDict[item[0]]]
        except (KeyError):
            try:
                itemStep = itemStepsList[billingItemDict[item[1]]]
            except (KeyError):
                continue

        qty = itemStep(activityLoc)

        print(f'Item: {item}, Qty: {qty}')

        report['WISE Qty'][index] = qty

    report.to_excel(reportLoc, index=False)
    print('DONE!')
    print("--- %s seconds ---" % (time.time() - start_time))

    x = input('Press Enter to Exit')

    exit()

except Exception as e:
    if hasattr(e, 'message'):
        print(e.message)

        x = input('Press Enter to Exit')
        
        exit()
    else:
        print('An error occured at ',e.args,e.__doc__)
        print('An error occured at ')

        x = input('Press Enter to Exit')

        exit()
