from fileinput import filename
import json
from django.shortcuts import render,redirect
from django.http import HttpResponse
from django.contrib import messages
import openpyxl
import xlsxwriter     
import numpy as np
import pandas as pd
from datetime import datetime
import math
from pathlib import Path


# Create your views here.
def main(request):
    return render(request,'output.html')

def analysis(request):
    context = {}
    if request.method == 'POST':
        excel_file_ = request.FILES["data_file"]
        
        if not excel_file_.name.endswith('.xlsx'):
            messages.error(request, 'data file is not a valid excel file')
    
        try:
            dataset = pd.read_excel(excel_file_)

            #Renaming column name to a valid column name format as per python
            dataset.rename(columns={'IN/OUT': 'IN_OUT'}, inplace=True)
            dataset.rename(columns={'E.Code': 'E_Code'}, inplace=True)

            #Droping rows that conatain 1 in 'IN_OUT' column
            dataset=dataset[dataset.IN_OUT!=1]

            #Converting Date column values into year-month-day format
            j,k=0,0
            for i in dataset.iloc[:,3]:
                dataset.Date[j]=datetime.strptime(str(i),'%Y%m%d').date()
                j+=1

            dataset['Date'] = pd.to_datetime(dataset['Date'], errors='coerce')

            #Calculating and appending Punch time difference for each employee with the corresponding E.Code into dataframe (l)
            l=[]
            for i in range(len(dataset)-1):
                j=i+1
                l.append([dataset.iloc[i]['E_Code'], datetime.combine(dataset.Date[0], datetime.strptime(str(dataset.iloc[j]['Time']),'%H%M%S').time()) - datetime.combine(dataset.Date[0], datetime.strptime(str(dataset.iloc[i]['Time']),'%H%M%S').time())])  
                i=i+2

            #Inserting date in datetime format after E.Code column
            for i in range(len(l)):
                l[i].insert(1,dataset.iloc[i]['Date'].date())

            #Replacing ":" with "." and getting upto 2 decimal places in Time to make mathematical operations feasible on i
            for i in range(len(l)):
                l[i][2] = l[i][2].__str__().replace(":",".")
                s=l[i][2]
                l[i][2]=s[0:-3]

            #Removing redundant rows and storing in new data-structure
            l2=[]
            for i in range(len(l)):
                if 'd' not in str(l[i][2]):
                    l2.append(l[i])

            #Converting string type to float for Time column
            for i in range(len(l2)):
                l2[i][2]=float("{:.2f}".format(float(l2[i][2])))

            """Converting time from float to mins, applied logic to check if it is greater than 9hrs & 30mins. Subtacting the 
            result to get over time (minutes) converting and storing it into hrs.minutes format. """

            l3=[]
            for i in range(len(l2)):
                frac, whole = math.modf(l2[i][2])
                mins=int(whole*60+frac*100)
                if mins>570:
                    mins=mins-570
                    hours=mins//60
                    minutes=mins%60
                    if minutes<10:
                        ot_time = float("{}.0{}".format(hours, minutes))
                    else:
                        ot_time = float("{}.{}".format(hours, minutes))
                    l2[i].append(ot_time)
                    l3.append(l2[i])

            # Converting 2d list to dataframe to apply CRUD opertations
            df = pd.DataFrame(l3, columns=['E_Code','Date','Total Time','OT'])

            #Droping an unnecesary column
            df=df.drop(['Total Time'], axis=1)

              

            """ Applying Groupby opr. to group dataframe on basis of E.Code and apply sum and count opr. to 
            compute Total_OT & Total_OT_Days """

            groupby_ecode_Total_OT = df.groupby(['E_Code']).sum()
            groupby_ecode_Total_OT_Days = df.groupby(['E_Code']).count()

            #Converting Groupby object into DataFrame to apply CRUD opr.
            df2=pd.DataFrame(groupby_ecode_Total_OT)
            df3=pd.DataFrame(groupby_ecode_Total_OT_Days)

            #Droping an unnecesary column
            df3=df3.drop(['Date'], axis=1)

            #Renaming Columns
            df2.rename(columns={'OT': 'Total_OT'}, inplace=True)
            df3.rename(columns={'OT': 'Total_OT_Days'}, inplace=True)

            #Creating a final dataframe (E.Code, Total OT_Days, Total OT)
            result = pd.concat([df2,df3], axis=1)

            """ Printing dataframe (E.Code, Total OT (after decimal place refers to mins, before refers to hrs; 
            e.g 0.27--> 27 minutes, 
                1.08--> 1hr & 8 minutes), 
                Total OT_Days) 
            """
            
            # Sorting dataframe in descending order on basis of Total OT_Days (User can get top x employees on basis of Total OT_Days)
            context2 = result.sort_values(by=['Total_OT_Days'], ascending=False)

            # Sorting dataframe in descending order on basis of Total OT (User can get top x employees on basis of Total OT)
            context3= result.sort_values(by=['Total_OT'], ascending=False)
            # text_file.write(html)
            # text_file.write(context2.to_html())
            # text_file.close()

            html = df.reset_index().to_json(orient='records', date_format='iso')
            data=[]
            data=json.loads(html)
            context['d']=data

            html2 = result.reset_index().to_json(orient='records', date_format='iso')
            data2=[]
            data2=json.loads(html2)
            context['d2']=data2

            html3 = context2.reset_index().to_json(orient='records', date_format='iso')
            data3=[]
            data3=json.loads(html3)
            context['d3']=data3

            html4 = context3.reset_index().to_json(orient='records', date_format='iso')
            data4=[]
            data4=json.loads(html4)
            context['d4']=data4  
            
            downloads_path = str(Path.home() / "Downloads")
            with pd.ExcelWriter(downloads_path+"/xyz3.xlsx") as writer:
                #writer = r"C:\Users\Kartikey\Desktop\shorttt.xlsx"
                # write dataframe to excel
                df.to_excel(writer,sheet_name='OT')
                result.to_excel(writer,sheet_name='Total_OT')
                context2.to_excel(writer,sheet_name='Descending_Total_OT_Days')
                context3.to_excel(writer,sheet_name='Descending_Total_OT')
            
        except Exception as e:
            return HttpResponse("Error Occured , Reason : " + str(e))

        

    return render(request,'analysis.html',context)


 