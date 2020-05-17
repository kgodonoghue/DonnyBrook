import pandas as pd 
from PIL import Image
import matplotlib.pyplot as plt
import os

payment_status='free'
payment_status='paid'
payment_status='all'

court_count='all'
#court_count='single'

date_range='18th May 2020 to 22nd May 2020'
i=27 # This is the link to the work week file

path='C:\main_folder\DTLC\Weekly_Utilization\ww' + str(i) 

imagelist=[]
airhall_list=[1,2,5,6,7]

os.chdir(path)

for airhall_number in airhall_list:    
    data = pd.read_csv("court_usage_ww"+str(i)+".csv") 
    no_days=data['bookingdate'].value_counts()
    no_days=no_days.count()
    data=data.fillna(0)
    data=data[data['booked_by']!='CLOSED']
    
    if airhall_number==0:
        court_count='all'
    else:
        court_count='single'        
    
    ''' single or all courts'''
    if court_count=='all':
        data=data
        airhall_number='all'
    if court_count=='single':
        data=data[data["courtname"]=="Outdoor Court "+str(airhall_number)]
    
    ''' choose payment status '''
    if payment_status=='all':
        data=data
    if payment_status=='paid':
        data_paid=data[(data['cost']==1) | (data['cost']==2)]
    if payment_status=='free':
        data_free=data[data['cost']==0]
    
    time_counts=data['starttime'].value_counts()
    
    if court_count=='all':
        utilization_time=(time_counts/5)/5
    else:
        utilization_time=(time_counts/5)*100
        
    utilization_time=pd.DataFrame(utilization_time)
    utilization_time['Time'] = utilization_time.index
    utilization_time.rename(columns={"starttime": "% Utilization"})
    utilization_time=utilization_time.sort_values(by='Time', ascending=True)
    utilization_time=utilization_time.set_index('Time')
    utilization_time.to_csv('utilization_time_ww'+str(i)+'.csv',index=True,header=True)
    
    fig, ax = plt.subplots(figsize=(20,15))
    ax.grid(linestyle='-', linewidth='1', color='grey') 
    plt.xticks(rotation = 90)
    plt.rc('xtick', labelsize=20) 
    plt.rc('ytick', labelsize=20)
    plt.plot(utilization_time,label='Utilization % point at time interval')
    plt.legend(loc='upper left', prop={'size': 20})    
    fig.suptitle('Court ' +str(airhall_number) + ' Utilization Summary 9:0am to 9.00pm_' + str(date_range), fontsize=25)
    plt.ylabel('Court Utilization', fontsize=25)
    plt.xlabel('Hour of Day', fontsize=25)
    plt.savefig('court' + str(airhall_number) + '.png')
    
    image1 = Image.open('court' + str(airhall_number) + '.png')     
    im = image1.convert('RGB')
    imagelist.append(im)  


im.save('ww' + str(i) +' Usage_Report.pdf',save_all=True, append_images=imagelist)

from PyPDF2 import PdfFileWriter, PdfFileReader
pages_to_keep = [1, 2, 3, 4, 5] # page numbering starts from 0
infile = PdfFileReader('ww' + str(i) +' Usage_Report.pdf', 'rb')
output = PdfFileWriter()

for j in pages_to_keep:
    p = infile.getPage(j)
    output.addPage(p)

with open('ww' + str(i) + '_Court_Usage_Report_' + str(date_range) + '_.pdf', 'wb') as f:
    output.write(f)
    
    
df_access=pd.read_csv('court_usage_ww27.csv',encoding = "ISO-8859-1")
df_access=df_access.join(df_access['opponents'].str.split(',', 2, expand=True).rename(columns={0:'Opponent 1', 1:'Opponent 2',2:'Opponent 3'}))
access_counts_1=df_access['booked_by'].value_counts()
access_counts_2=df_access['Opponent 1'].value_counts()
access_counts_3=df_access['Opponent 2'].value_counts()
access_counts_4=df_access['Opponent 3'].value_counts()
access_counts_1.to_csv('primary_booking_counts1.csv',index=True,header=True)
access_counts_2.to_csv('primary_booking_counts2.csv',index=True,header=True)
access_counts_3.to_csv('primary_booking_counts3.csv',index=True,header=True)
access_counts_4.to_csv('primary_booking_counts4.csv',index=True,header=True)
access1=pd.read_csv('primary_booking_counts1.csv',encoding = "ISO-8859-1")
access2=pd.read_csv('primary_booking_counts2.csv',encoding = "ISO-8859-1")
access3=pd.read_csv('primary_booking_counts3.csv',encoding = "ISO-8859-1")
access4=pd.read_csv('primary_booking_counts4.csv',encoding = "ISO-8859-1")     
df_new1=df_access['booked_by']
df_new2=df_access['Opponent 1']
df_new3=df_access['Opponent 2']
df_new4=df_access['Opponent 3']
df_new=df_new1.append(df_new2)
df_new=df_new.append(df_new3)
df_new=df_new.append(df_new4)
df_new.to_csv('all_counts.csv',index=True,header=True)
df_booking_all=pd.read_csv('all_counts.csv',encoding = "ISO-8859-1")
df_booking_all=df_booking_all.rename(columns={"Unnamed: 0": "Count", "0": "Name"})
booking_counts=df_booking_all['Name'].value_counts()
booking_counts.to_csv('all_counts_final.csv',index=True,header=True)  
df_count_final=pd.read_csv('all_counts_final.csv',encoding = "ISO-8859-1")
df_count_final=df_count_final.rename(columns={"Unnamed: 0": "Name", "Name": "Count"})
df_count_final=df_count_final[df_count_final['Name'] != 'CLOSED']
df_count_final=df_count_final[df_count_final['Name']!='CLUB CLOSED']
df_count_final.to_csv('all_counts_final.csv',index=True,header=True)  
   
    
