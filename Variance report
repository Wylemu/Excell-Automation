from datetime import date
from gettext import npgettext
from operator import index
from platform import win32_edition
import numpy
import pandas as pd
import numpy as np
import glob
from csv import writer
from datetime import datetime
import smtplib
import win32com.client as win32

#---------------------------PROGRAMMMMMME-------------------------------------

# ----IMPORTING AND SORTING-------------------
var_report = "C:\\Users\\user\\Documents\\Stock take Variances\\StockTakeVarianceDetail.csv"
df1 = pd.read_csv(var_report)
df1['StockTakeDate'] = pd.to_datetime(df1['StockTakeDate'])
df = df1[df1['StockTakeStatus'] == 'Approved']
df_colNames = pd.DataFrame(columns= df1.columns)
df1['IMEI/Serial No'] = df1['IMEI/Serial No'].astype(str)
df = df1.groupby('Store')
dt_string = datetime.strftime(datetime.now(),'%Y-%m-%d')
Tday =pd.to_datetime('today',format='%Y-%m-%d').normalize()
Today_ = str(Tday)
Tday_data = datetime.strftime(df1.at[0,'StockTakeDate'],'%Y-%m-%d')
print(df1['IMEI/Serial No'])

# ---VARIABLES -----

NSC = 'SABC0844 MTN Store - Nicolway Shopping CentreL47 - Bryanston'
HVM = 'SABC1367 MTN Store - Highveld Mall Shop 222'
CGM = 'SABC1726 MTN Store - Castle Gate Retail Centre 35 - Pretoria'   
MM = 'SABC0471 MTN Store - Middleburg Mall 11 - Middleburg'
SM = 'SABC0958 MTN Store - Secunda Mall LG28 - Secunda'
GSM = 'SABC2299 MTN Store - Greenstone Shopping Centre L073 - Edenvale'
PM = 'SABC2300 MTN Lite - Phola Mall 62A - KwaMhlanga'
KM = 'SABC2301  MTN Lite - Kwagga Mall 100 - Kwaggafontein'
HM = 'SABC2302 MTN Store - Highland Mews 24A -  Witbank'
JM = 'SABC0870 MTN Store - Jubilee Mall 16 - Hammanskraal West'

#-----------------------------DATA FRAMES---------------------------------------------

#NICOLWAY MALL
if NSC in df1['Store'].values:
        df_NSC = df.get_group(NSC)
        Over_item = df_NSC.loc[df_NSC['VarianceType']=='OVER']
        Under_item = df_NSC.loc[df_NSC['VarianceType']=='UNDER']     
        totalRisk_NSC = Under_item['Total Cost Price for the Variance OMS/Actual'].sum()
        totalUnderItems_NSC = Under_item['OMS2/Actual Variance Qty'].sum()
        T_O_NSC = Over_item['OMS2/Actual Variance Qty'].sum()
       
else:
        
        df_NSC = df_colNames
        totalRisk_NSC = 0
        T_O_NSC:0
        
        
#HIGHVELD MALL       
if HVM in df1['Store'].values:
    df_HVM = df.get_group(HVM)
    Over_item_hvm = df_HVM.loc[df_HVM['VarianceType']=='OVER']
    Under_item_hvm = df_HVM.loc[df_HVM['VarianceType']=='UNDER']
    totalUnderItems_Hvm = Under_item_hvm['OMS2/Actual Variance Qty'].sum()
    totalRisk_Hvm = df_HVM['Total Cost Price for the Variance OMS/Actual'].sum()
    T_O_HVM = Over_item_hvm['OMS2/Actual Variance Qty'].sum()
else:
    df_HVM = df_colNames
    
    totalRisk_Hvm =0
    totalUnderItems_Hvm = 0
    T_O_HVM = 0
    
    
#GREEN STONE MALL
    
if GSM in df1['Store'].values:
    df_GSM = df.get_group(GSM)
    Over_item_GSM = df_GSM.loc[df_GSM['VarianceType']=='OVER']
    Under_item_GSM = df_GSM.loc[df_GSM['VarianceType']=='UNDER']
    totalUnderItems_GSM = Under_item_GSM['OMS2/Actual Variance Qty'].sum()
    totalRisk_GSM = df_GSM['Total Cost Price for the Variance OMS/Actual'].sum()
    T_O_GSM = Over_item_GSM['OMS2/Actual Variance Qty'].sum()
    
else:
     df_GSM = df_colNames
     T_O_GSM = 0     
     totalUnderItems_GSM =0 
     totalRisk_GSM =0
     
     
#MIDDELBURG MALL
    
if MM in df1['Store'].values:
    df_MM = df.get_group(MM)
    Over_item_MM = df_MM.loc[df_MM['VarianceType']=='OVER']
    Under_item_MM = df_MM.loc[df_MM['VarianceType']=='UNDER']
    totalUnderItems_MM = Under_item_MM['OMS2/Actual Variance Qty'].sum()
    T_O_MM = Over_item_MM['OMS2/Actual Variance Qty'].sum()
    totalRisk_MM = df_MM['Total Cost Price for the Variance OMS/Actual'].sum()
    
else:
    df_MM = df_colNames
    totalRisk_MM =0
    T_O_MM =0
    totalUnderItems_MM =0
    
     
#SECUNDA MALL  
 
if SM in df1['Store'].values:
    df_SM = df.get_group(SM)
    Over_item_SM = df_SM.loc[df_SM['VarianceType']=='OVER']
    Under_item_SM = df_SM.loc[df_SM['VarianceType']=='UNDER']
    totalUnderItems_SM = Under_item_SM['OMS2/Actual Variance Qty'].sum()
    T_O_SM = Over_item_SM['OMS2/Actual Variance Qty'].sum()
    totalRisk_SM = df_SM['Total Cost Price for the Variance OMS/Actual'].sum()
    
else:
    df_SM = df_colNames
    totalRisk_SM =0
    T_O_SM =0
    totalUnderItems_SM =0
    
    
       
# Castle Gate MALL

if CGM in df1['Store'].values:
        df_CGM = df.get_group(CGM)
        Over_item_CGM = df_CGM.loc[df_CGM['VarianceType']=='OVER']
        Under_item_CGM = df_CGM.loc[df_CGM['VarianceType']=='UNDER']
        totalUnderItems_CGM = Under_item_CGM['OMS2/Actual Variance Qty'].sum()
        T_O_CGM = Over_item_CGM['OMS2/Actual Variance Qty'].sum()
        totalRisk_CGM = df_CGM['Total Cost Price for the Variance OMS/Actual'].sum()
        
else:
        df_CGM = df_colNames
        totalRisk_CGM =0
        totalUnderItems_CGM= 0
        T_O_CGM = 0
        
              
#Phola MAll 
       
if PM in df1['Store'].values:
        df_PM = df.get_group(PM)
        Over_item_PM = df_PM.loc[df_PM['VarianceType']=='OVER']
        Under_item_PM = df_PM.loc[df_PM['VarianceType']=='UNDER']
        totalUnderItems_PM = Under_item_PM['OMS2/Actual Variance Qty'].sum()
        T_O_PM = Over_item_PM['OMS2/Actual Variance Qty'].sum()
        totalRisk_PM = df_PM['Total Cost Price for the Variance OMS/Actual'].sum() 
          
else:
    df_PM = df_colNames
    totalUnderItems_PM= 0
    T_O_PM = 0
    totalItems_PM = 0
       

#Kwagga Mall

if KM in df1['Store'].values:
    df_KM = df.get_group(KM)
    Over_item_KM = df_KM.loc[df_KM['VarianceType']=='OVER']
    Under_item_KM = df_KM.loc[df_KM['VarianceType']=='UNDER']
    totalUnderItems_KM = Under_item_KM['OMS2/Actual Variance Qty'].sum()
    T_O_KM = Over_item_KM['OMS2/Actual Variance Qty'].sum()
    totalRisk_KM = df_KM['Total Cost Price for the Variance OMS/Actual'].sum()
  
else:
    df_KM =df_colNames
    totalRisk_KM = 0
    totalUnderItems_KM= 0
    T_O_KM = 0
    
# HIGHLAND MEWS
    
if HM in df1['Store'].values:
    df_HM = df.get_group(HM)
    Over_item_HM = df_HM.loc[df_HM['VarianceType']=='OVER']
    Under_item_HM = df_HM.loc[df_HM['VarianceType']=='UNDER']
    totalUnderItems_HM = Under_item_HM['OMS2/Actual Variance Qty'].sum()
    T_O_HM = Over_item_HM['OMS2/Actual Variance Qty'].sum()
    totalRisk_HM = df_HM['Total Cost Price for the Variance OMS/Actual'].sum()
   
else:
    df_HM =df_colNames
    totalRisk_HM = 0
    totalUnderItems_HM= 0
    T_O_HM = 0
        
# JUBILLIE MALL
    
if JM in df1['Store'].values:    
    df_JM = df.get_group(JM)
    Over_item_JM = df_JM.loc[df_JM['VarianceType']=='OVER']
    Under_item_JM = df_JM.loc[df_JM['VarianceType']=='UNDER']
    totalUnderItems_JM = Under_item_JM['OMS2/Actual Variance Qty'].sum()
    T_O_JM = Over_item_JM['OMS2/Actual Variance Qty'].sum()    
    totalRisk_JM = df_JM['Total Cost Price for the Variance OMS/Actual'].sum()   
else:
    df_JM =df_colNames
    totalRisk_JM = 0
    totalUnderItems_JM= 0
    T_O_JM = 0
  
    
    
#-------------------PRINT TO EXCELL-----------------------------------


with pd.ExcelWriter("C:\\Users\\user\\Documents\\Stock take Variances\\StockTakeVarianceDetail" +Tday_data+'.xlsx',engine= 'xlsxwriter') as writer_1:
   
#write the sheets with different groups

    df_NSC.to_excel(writer_1,index= False,sheet_name='NSC')
    df_HVM.to_excel(writer_1,index= False,sheet_name='HVM')
    df_CGM.to_excel(writer_1,index= False,sheet_name='CGM')
    df_MM.to_excel(writer_1,index= False,sheet_name='MM')
    df_SM.to_excel(writer_1,index= False,sheet_name='SM')
    df_GSM.to_excel(writer_1,index= False,sheet_name='GSM')
    df_PM.to_excel(writer_1,index= False,sheet_name='PM')
    df_KM.to_excel(writer_1,index= False,sheet_name='KM')
    df_HM.to_excel(writer_1,index= False,sheet_name='HM') 
    df_JM.to_excel(writer_1,index= False,sheet_name='JM')       

#-----------------SAVE THE FILE-----------------------------
writer_1.save()

#----- SEND EMAIL USING OUTLOOK-------------------------------------
 
olApp = win32.Dispatch('Outlook.Application')
olns = olApp.GetNameSpace('MAPI')
Atta_ment = "C:\\Users\\user\\Documents\\Stock take Variances\\StockTakeVarianceDetail" +Tday_data+'.xlsx'
mailItem = olApp.CreateItem(0)
mailItem.Subject = 'Variance Report'+Tday_data 
mailItem.BodyFormat = 1
mailItem.Body = 'Please see attached Your Variance report for '+Tday_data+ ' ,May you Please respond on All Variances Before 12PM today'
mailItem.To = 'Stock Take 1'
mailItem.CC = 'Thokozane.Ngwenya@mtn.com;vincent@onetouchmobility.co.za;blessing@onetouchmobility;Vincent.Nomtshongwane@mtn.com'
mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olns.Accounts.Item('walter@onetouchmobility.co.za')))
mailItem.Display()
mailItem.Attachments.Add(Atta_ment)
mailItem.BodyFormat = 2
mailItem.HTMLBody = """
    <html>
    <body>
    <div>
    <b> 
    <P>    
    Good Morning Team     
    </P>
    Please see attached Your Variance report for """ +Tday_data +"""     
    May you Please respond on All Variances Before 12PM today 
    </b>
    </div>
    <br></br>
    <table width = "800" border = "1" cellpadding ="0" cellspacing ="1" color = "green">
        <thead >
            <tr border = "0" cellpadding ="0" cellspacing ="1">
                <th font size= "25">Store </th>
                <th>Under Items </th>
                <th>Over Items </th>
                <th>Total Risk </th>
            </tr>
        </thead>
    <tbody border = "0" cellpadding ="0" cellspacing ="1">
        <tr>
            <td width="100" align ="center">HVM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_Hvm)+"""</td>
            <td width="100" align ="center">"""+str(T_O_HVM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_Hvm.round(2)) +"""</td>
           
        </tr>
        <tr>
            <td width="100" align ="center">NSC</td>
            <td width="100" align ="center">"""+str(totalUnderItems_NSC)+"""</td>
            <td width="100" align ="center">"""+str(T_O_NSC) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_NSC.round(2)) +"""</td>
           
        </tr>
        <tr>
            <td width="100" align ="center">GSM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_GSM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_GSM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_GSM.round(2)) +"""</td>
           
        </tr>   
        <tr>
            <td width="100" align ="center">MM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_MM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_MM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_MM.round(2)) +"""</td>
           
        </tr> 
        <tr>
            <td width="100" align ="center">SM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_SM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_SM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_SM.round(2)) +"""</td>
           
        </tr>
         <tr>
            <td width="100" align ="center">PM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_PM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_PM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_PM.round(2)) +"""</td>
           
        </tr>
         <tr>
            <td width="100" align ="center">KM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_KM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_KM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_KM.round(2)) +"""</td>           
        </tr>
         <tr>
            <td width="100" align ="center">JM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_JM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_JM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_JM.round(2)) +"""</td>
           
        </tr> 
         <tr>
            <td width="100" align ="center">HM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_HM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_HM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_HM.round(2)) +"""</td>
           
        </tr>
         <tr>
            <td width="100" align ="center">CGM</td>
            <td width="100" align ="center">"""+str(totalUnderItems_CGM)+"""</td>
            <td width="100" align ="center">"""+str(T_O_CGM) +"""</td>
            <td width="100" align ="center">"""+str(totalRisk_CGM.round(2)) +"""</td>
           
        </tr>                                               
    </tbody>
    </table>
    
     </body>
    </html> 
"""
