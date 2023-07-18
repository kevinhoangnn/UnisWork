import BNPauto
import WISEauto
from datetime import date, timedelta, datetime
import os.path
from logging import info, error, basicConfig, DEBUG
from os import rename
from pandas import read_excel
from calendar import monthrange
from shutil import copy, move
from glob import glob
from time import time

start_time = time()

def getInvoice(acc, facility, startP, endP, accName, cycle, wise = False):
    try:
        downloaddir = 'C:\\Users\\' + os.getlogin() + '\\Downloads'
        accsdir = 'C:\\Users\\' + os.getlogin() +'\\Desktop\\Discrepancy Reports\\' + cycle

        if wise:
            WISEauto.exportReport(acc, facility, startP, endP)
            newName = accName + '-' + facility + '-' + cycle + '-Activity_Report.xlsx'
            copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\' +  cycle + '\\03 - Current Activity reports\\' + newName
        else:
            if BNPauto.exportHandle(acc, facility, startP, endP, accName) == False: 
                return False

            newName = accName + '-' + facility + '-' + cycle + '-Invoice.xlsx'
            copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\' + cycle + '\\02 - Current Invoices\\' + newName
        
        fileList = list(filter(os.path.isfile, glob(downloaddir + '\\*.xlsx')))
        fileList.sort(key=lambda x: os.path.getmtime(x))
        file = fileList[len(fileList) - 1]
        fileEnd = os.path.basename(file)
        
        if wise:
            move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, '01 - Historical Activity reports', fileEnd))
            copy(os.path.join(accsdir, '01 - Historical Activity reports', fileEnd), copyPath)
            info('-------------------------' + '\n\n\nDownloaded Wise ACtivity Report ' + str(datetime.now()) + '\n\n\n-------------------------------------------------')
        else:
            move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, '00 - Historical Invoices', fileEnd))
            copy(os.path.join(accsdir, '00 - Historical Invoices', fileEnd), copyPath)
            info('-------------------------' + '\n\n\nDownloaded BNP Invoice ' + str(datetime.now()) + '\n\n\n-------------------------------------------------')

            if fileEnd == 'Invoice[Multi].xlsx':
                rename(os.path.join(accsdir, '00 - Historical Invoices', fileEnd), os.path.join(accsdir, '00 - Historical Invoices', accName + "-" + facility + "-" + fileEnd))

        print(f'Downloaded and Copied {newName} \n')

        return True

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)
            invoiceAccs['Downloaded'][index] = e.message
            invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
        else:
            print('An error occured at ', e.args, e.__doc__)
            invoiceAccs['Downloaded'][index] = e.args, e.__doc__
            invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
        info('-------------------------\n\n\n' + 'Downloaded Error ' + str(datetime.now()))
        error(e)
        info('\n\n\n-------------------------------------------------')


def excelOutput(bnpAcc, wiseAcc, facility, startP, endP, accName, cycle):
    bnpStart = startP.strftime("%m/%d/%y")
    bnpEnd = endP.strftime("%m/%d/%y")
    
    wiseStart = startP.strftime("%y-%m-%d")
    wiseEnd = endP.strftime("%y-%m-%d")

    bnpInvoice = getInvoice(bnpAcc, facility, bnpStart, bnpEnd, accName, cycle)
    wiseInvoice = False

    if (bnpInvoice):
        wiseInvoice = getInvoice(wiseAcc, facility, wiseStart, wiseEnd, accName, cycle, True)
        if (wiseInvoice):
            invoiceAccs['Downloaded'][index] = True
            invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
        else:
            invoiceAccs['Downloaded'][index] = 'No Wise Report'
            invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
            return False
    else:
        invoiceAccs['Downloaded'][index] = 'No BNP Invoice'
        invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
        return False
    
    return True

def getBimonthlyDate(todayDate):
    if todayDate.day < 16:
        previousMonth = todayDate.month - 1 if todayDate.month != 1 else 12
        previousYear = todayDate.year - 1 if todayDate.month == 1 else todayDate.year
        start = date(previousYear, previousMonth, 16)
        end = date(previousYear, previousMonth, monthrange(previousYear, previousMonth)[1])
    elif today.day >= 16:
        start = date(todayDate.year, todayDate.month, 1)
        end = date(todayDate.year, todayDate.month, 15)
    
    return start, end

def getWeeklyDate(todayDate):
    idx = (todayDate.weekday() + 1) % 7 # MON = 0, SUN = 6 -> SUN = 0 .. SAT = 6

    sun = todayDate - timedelta(7+idx)
    sat = todayDate - timedelta(7+idx-6)
    
    return sun, sat

if __name__ == '__main__':

    basicConfig(filename = "logs.txt", level = DEBUG, format = "%(asctime)s %(message)s")

    invoiceAccs = read_excel('BNP Excel Sheet.xlsx', sheet_name='Account_Fac_Freq')
    invoiceAccs['Downloaded'] = ''
    bimonthly = False
    weekly = False

    today = date.today()
    today = date(2023, 1, 29)
    dayName = today.strftime("%A")

    if dayName == 'Sunday' and (today.day == 16 or today.day == 1):
        bimonthly = True
        weekly = True
    elif dayName == 'Sunday':
        invoiceAccs = invoiceAccs[invoiceAccs['BillingFreq']== 'Weekly']
        weekly = True
    elif today.day == 16 or today.day == 1:
        invoiceAccs = invoiceAccs[invoiceAccs['BillingFreq'] == 'Bimonthly']
        bimonthly = True
    else:
        raise Exception("Not a valid day to run program!")


    for index in invoiceAccs.index:
        fac = invoiceAccs['Facility Name'][index]
        bnpName = invoiceAccs['BNP Account Name'][index]
        wiseName = invoiceAccs['Wise Account Name'][index]
        accName = invoiceAccs['AccountName'][index]

        
        info('-------------------------' + f'\n\n\nGetting Invoice for {accName}-{fac} ' + str(datetime.now()) + '\n\n\n-------------------------------------------------')

        if bimonthly and weekly:
            if invoiceAccs['BillingFreq'][index] == 'Bimonthly':
                #Assuming is ran on 1st and 16th
                startPeriod, endPeriod = getBimonthlyDate(today)

                downloadBool = excelOutput(bnpName, wiseName, fac, startPeriod, endPeriod, accName, 'Bimonthly')

            elif invoiceAccs['BillingFreq'][index] == 'Weekly':
                startPeriod, endPeriod = getWeeklyDate(today)

                downloadBool = excelOutput(bnpName, wiseName, fac, startPeriod, endPeriod, accName, 'Weekly')

        elif bimonthly:
            startPeriod, endPeriod = getBimonthlyDate(today)

            downloadBool = excelOutput(bnpName, wiseName, fac, startPeriod, endPeriod, accName, 'Bimonthly')

        elif weekly:
            startPeriod, endPeriod = getWeeklyDate(today)

            downloadBool = excelOutput(bnpName, wiseName, fac, startPeriod, endPeriod, accName, 'Weekly')
            

    info('-------------------------' + '\n\n\nInvoices Downloaded ' + str(datetime.now()))
    info(('Completion Time - %s seconds' % (time() - start_time)) + '\n\n\n-------------------------------------------------')
    print("--- %s seconds ---" % (time() - start_time))
    x = input("Press Enter to finish. ")