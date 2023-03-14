# -*- coding: utf-8 -*-
"""
Created on Tue Oct 19 07:36:23 2021

@author: dpatel
"""


import pandas as pd
import os as os
import time as time
import shutil
import datetime as datetime
import glob as glob
#import xlrd
#import re
#import numpy as np
import openpyxl

#Who is running the file?

username = 'dpatel'
#username = ''

# Set folder and get basic files to use
os.getcwd()
os.chdir("C:/Users/" + username + "/OneDrive - Weber-Stephen Products LLC/US Analytics/POS/Weekly POS")
path = "C:/Users/" + username + "/OneDrive - Weber-Stephen Products LLC/US Analytics/POS/Weekly POS/"
os.getcwd()

supportpath = "C:/Users/" + username + "/OneDrive - Weber-Stephen Products LLC/US Analytics/Support Files/"

dfFiscalCal = pd.read_excel(supportpath + 'Fiscal Calendar v3.xlsx', sheet_name='Fiscal Calendar')
dfProdHier = pd.read_excel(supportpath + 'Weber Prod Hierarchy.xlsx', sheet_name='Weber Prod Hierarchy')
dfWeeklyLookup = pd.read_excel(supportpath +'Fiscal Calendar v3.xlsx', sheet_name='Weekly Lookup')
dfWMTFiscalCal = pd.read_excel(supportpath + "Walmart Fiscal Calendar.xlsx",dtype={'2020ReportFiscalWeek':str})
dfWMTFiscalCal['2020DateLookupReportDate'] = pd.to_datetime(dfWMTFiscalCal['2020DateLookupReportDate']).dt.date
dfMsrp = pd.read_excel(supportpath + 'MSRP lookup.xlsx', sheet_name='msrp', dtype={'Item No.':str})




dfPCR2021 = pd.DataFrame()
for PCRfiles in os.listdir(path + "PC Richard/Upated/"):
    for f in glob.glob(path + "PC Richard/Upated/"+ PCRfiles):
        sheets_dict = pd.read_excel(f, sheet_name=None, header=1, dtype={'Model':str,'Store#':str}, usecols='A:E')
        for PCRfiles, frame in sheets_dict.items():
            frame['Sheet'] = PCRfiles
            dfPCR2021 = dfPCR2021.append(frame, ignore_index=True)
       
dfPCR2021['Month'] = dfPCR2021['Sheet'].str[:2]
dfPCR2021['Day'] = dfPCR2021['Sheet'].str[2:4]
dfPCR2021['Year'] = dfPCR2021['Sheet'].str[4:8]

dfPCR2021['Calendar Day'] = pd.to_datetime(dfPCR2021[['Year', 'Month', 'Day']])

dfPCR2021.rename(columns = {'Model': 'Model #'}, inplace=True)

dfPCR2021 = dfPCR2021[['Store#','Model #','Week Start','Calendar Day']]

dfPCR2021.rename(columns = {'Week Start': 'Units Sold'}, inplace=True)


dfPCR2021 = pd.merge(dfPCR2021, dfFiscalCal, how='left', left_on='Calendar Day',right_on='Calendar Date')


dfPCR2021 = pd.merge(dfPCR2021, dfProdHier, how='left', left_on='Model #', right_on='Material')

dfPCR2021 = pd.merge(dfPCR2021, dfMsrp, how='left', left_on='Model #', right_on='Item No.' )



dfPCR2021['Retail Dollars'] = dfPCR2021['Units Sold'] * dfPCR2021['MSRP']

dfPCR2021['Week Start'] = dfPCR2021['SundayWeekStartDate']
dfPCR2021['Week Start'] = pd.to_datetime(dfPCR2021['Week Start']).dt.date

dfPCR2021['Retailer'] = 'PC Richard'

#dfPCR2021.loc[dfPCR2021['Store#'] == '12', 'Channel'] = 'eCommerce'
#dfPCR2021.loc[dfPCR2021['Store#'] != '12', 'Channel'] = 'B&M'

dfPCR2021['Channel'] = dfPCR2021['Store#'].apply(lambda x: 'eCommerce' if x =='12' else 'B&M')


dfPCR2021 = dfPCR2021.drop(columns=['Material'])

dfPCR2021.rename(columns = {'Model #': 'Material'}, inplace=True)

dfPCR2021 = dfPCR2021[['Retailer','Channel','Product Category','Product Family','Material','Material and Desc','Model','Week Start','Retail Dollars','Units Sold']]

dfPCR2021 = dfPCR2021[dfPCR2021['Units Sold'].notnull()]


dfPCR2021.to_csv('C:/Users/' + username + '/OneDrive - Weber-Stephen Products LLC/US Analytics/POS/Weekly POS/ABT PCR BBQ working folder/PC Richards Working Filev2.csv')



#ABT

            
dfABT = pd.DataFrame()
for ABTfiles in os.listdir(path + "ABT/ABT all/"):
    temp=pd.read_csv(path + "ABT/ABT all/" + ABTfiles, dtype={'VSN':str})
    temp['filename'] = ABTfiles
    dfABT = pd.concat([dfABT,temp],ignore_index=True)
    
    

dfABT = dfABT.drop(columns=['VE_CD'])
dfABT = dfABT.drop(columns=['MNR_CD'])
dfABT = dfABT.drop(columns=['DES'])
dfABT = dfABT.drop(columns=['WEEK_END_DT'])

dfABTbm = dfABT.iloc[:, : 7]
dfABTbm = dfABTbm.drop(dfABTbm.columns[3:5], axis=1)


dfABTbm.rename(columns = {'VSN': 'Material'}, inplace=True)
dfABTbm.rename(columns = {'SAL_01': 'Units Sold'}, inplace=True)
dfABTbm.rename(columns = {'WEEK_START_DT': 'Week Start'}, inplace=True)
dfABTbm.rename(columns = {'RET_01': 'Units Return'}, inplace=True)

dfABTbm['Channel'] = 'B&M'

dfABTecom = dfABT.iloc[:, : 7]
dfABTecom = dfABTecom.drop(dfABTecom.columns[1:3], axis=1)

dfABTecom.rename(columns = {'VSN': 'Material'}, inplace=True)
dfABTecom.rename(columns = {'SAL_03': 'Units Sold'}, inplace=True)
dfABTecom.rename(columns = {'WEEK_START_DT': 'Week Start'}, inplace=True)
dfABTecom.rename(columns = {'RET_03': 'Units Return'}, inplace=True)


dfABTecom['Channel'] = 'eCommerce'

dfABT2021 = pd.concat([dfABTbm,dfABTecom], ignore_index=True)
dfABT2021['Units Sold'] = dfABT2021['Units Sold'] - dfABT2021['Units Return']
dfABT2021 = dfABT2021.drop(columns=['Units Return'])
dfABT2021 = dfABT2021.drop(columns=['QTYOH'])


dfABT2021 = pd.merge(dfABT2021, dfMsrp, how='left', left_on='Material', right_on='Item No.' )

dfABT2021['Retail Dollars'] = dfABT2021['Units Sold'] * dfABT2021['MSRP']

dfABT2021 = dfABT2021.drop(columns=['Item No.'])
dfABT2021 = dfABT2021.drop(columns=['MSRP'])

dfABT2021 = pd.merge(dfABT2021, dfProdHier, how='left', left_on='Material', right_on='Material' )


dfABT2021['Week Start'] = pd.to_datetime(dfABT2021['Week Start']).dt.normalize()
dfABT2021['Retailer'] = 'ABT'


dfABT2021 = dfABT2021[['Retailer','Channel','Product Category','Product Family','Material','Material and Desc','Model','Week Start','Retail Dollars','Units Sold']]


dfABT2021.to_csv('C:/Users/' + username + '/OneDrive - Weber-Stephen Products LLC/US Analytics/POS/Weekly POS/ABT PCR BBQ working folder/ABT Working Filev2.csv')



#BBQ 


dfBBQ2021 = pd.DataFrame()
for BBQfiles in os.listdir(path + "BBQ historical/BBQ current data/"):
    for f in glob.glob(path + "BBQ historical/BBQ current data/"+ BBQfiles):
        sheets_dict = pd.read_excel(f, sheet_name=None, header=0, dtype={'Model #':str}, usecols='A:E')
        for BBQfiles, frame in sheets_dict.items():
            frame['Sheet'] = BBQfiles
            dfBBQ2021 = dfBBQ2021.append(frame, ignore_index=True)
        

dfBBQ2021['Month'] = dfBBQ2021['Sheet'].str[:2]
dfBBQ2021['Day'] = dfBBQ2021['Sheet'].str[2:4]
dfBBQ2021['Year'] = dfBBQ2021['Sheet'].str[4:8]



dfBBQ2021['Calendar Day'] = pd.to_datetime(dfBBQ2021[['Year', 'Month', 'Day']])



dfBBQ2021 = dfBBQ2021.drop(columns=['Name'])
dfBBQ2021 = dfBBQ2021.drop(columns=['PPID'])
dfBBQ2021 = dfBBQ2021.drop(columns=['Units Sold 2020'])


dfBBQ2021 = pd.merge(dfBBQ2021, dfFiscalCal, how='left', left_on='Calendar Day',right_on='Calendar Date')


dfBBQ2021 = pd.merge(dfBBQ2021, dfProdHier, how='left', left_on='Model #', right_on='Material')

dfBBQ2021 = pd.merge(dfBBQ2021, dfMsrp, how='left', left_on='Model #', right_on='Item No.' )


dfBBQ2021['Retail Dollars'] = dfBBQ2021['Units Sold 2021'] * dfBBQ2021['MSRP']



dfBBQ2021['Week Start'] = dfBBQ2021['SundayWeekStartDate']
dfBBQ2021['Week Start'] = pd.to_datetime(dfBBQ2021['Week Start']).dt.date

dfBBQ2021.rename(columns = {'Units Sold 2021': 'Units Sold'}, inplace=True)
dfBBQ2021['Retailer'] = 'BBQ Guys'

dfBBQ2021['Channel'] = 'eCommerce'


dfBBQ2021 = dfBBQ2021.drop(columns=['Material'])

dfBBQ2021.rename(columns = {'Model #': 'Material'}, inplace=True)


dfBBQ2021 = dfBBQ2021[['Retailer','Channel','Product Category','Product Family','Material','Material and Desc','Model','Week Start','Retail Dollars','Units Sold']]


dfBBQ2021.to_csv('C:/Users/' + username + '/OneDrive - Weber-Stephen Products LLC/US Analytics/POS/Weekly POS/ABT PCR BBQ working folder/BBQ Guys Working Filev2.csv')





folder_path = path + "/ABT PCR BBQ working folder/"
dfAgg = pd.DataFrame()
for file_ in os.listdir(folder_path):
    temp=pd.read_csv(folder_path + file_)
    dfAgg=pd.concat([dfAgg,temp],ignore_index=True)

dfAgg['Week Start'] = pd.to_datetime(dfAgg['Week Start']).dt.normalize()

#dfAgg = pd.merge(dfAgg, dfWeeklyLookup, how='left', left_on='Week Start', right_on='SundayWeekStartDate')

timestamp = time.strftime('%m-%d-%Y-%H%M', time.localtime())
aggFileName = "Weekly POS Export " + str(timestamp) + ".csv"

dfAgg.to_csv(path + "ABT BBQ PCR tableau file/" + aggFileName,index=False)

print("Blended data completed and exported")



