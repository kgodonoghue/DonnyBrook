import pandas as pd 
from PIL import Image
import matplotlib.pyplot as plt

payment_status='free'
payment_status='paid'
payment_status='all'

court_count='all'
court_count='single'

date_range='xx day xx month 2020 to xx day xx month 2020'

i=1
imagelist=[]

for airhall_number in range(0,9):    
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
        data=data[data["courtname"]=="Airhall Court "+str(airhall_number)]
    
    ''' choose payment status '''
    if payment_status=='all':
        data=data
    if payment_status=='paid':
        data_paid=data[(data['cost']==1) | (data['cost']==2)]
    if payment_status=='free':
        data_free=data[data['cost']==0]
    
    time_counts=data['starttime'].value_counts()
    
    if court_count=='all':
        utilization_time=(time_counts*(15/24))/8
    else:
        utilization_time=(time_counts*15/24)
        
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
    fig.suptitle('Court ' +str(airhall_number) + ' Utilization Summary 7:30am to 10.30pm_' + str(date_range), fontsize=25)
    plt.ylabel('Court Utilization', fontsize=25)
    plt.xlabel('Hour of Day', fontsize=25)
    plt.savefig('court' + str(airhall_number) + '.png')
    
    image1 = Image.open(r'C:\main_folder\DTLC\covid_court_use_report\court'+str(airhall_number)+'.png')     
    im = image1.convert('RGB')
    imagelist.append(im)  

im.save(r'C:\main_folder\DTLC\covid_court_use_report\Court_Usage_Report.pdf',save_all=True, append_images=imagelist)

from PyPDF2 import PdfFileWriter, PdfFileReader
pages_to_keep = [1, 2, 3, 4, 5, 6, 7, 8, 9] # page numbering starts from 0
infile = PdfFileReader('Court_Usage_Report.pdf', 'rb')
output = PdfFileWriter()

for i in pages_to_keep:
    p = infile.getPage(i)
    output.addPage(p)

with open('newfile.pdf', 'wb') as f:
    output.write(f)   
    