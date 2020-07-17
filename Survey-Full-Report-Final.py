#!/usr/bin/env python
# coding: utf-8

# In[5]:


import pandas as pd
import re
import six
import os.path,sys
from pathlib import Path


print("This program is designed to create Environmental Discovery Reports for all \ncustomers who have taken the survey")
print("The output will be files named as SurveyReport_<customer name>.pdf for each customer")
file = input("Please enter the full path and name of survey results xls file:\n")

#df = pd.read_excel('Common DNAC Land and Ad opt - Environment Discovery - Responses -  2020-7-8 13-48 26996.xlsx')

my_file = Path(file)  
if os.path.isfile(my_file) == False:
    print("The file does not exist. Easiest way is to copy the xlsx file into the same directory as the exe")
    quit()
else:
    df = pd.read_excel(file)


# In[3]:


from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, inch, A4

a = df.columns
for i in range(0,len(df.index)): 
    print(df[a[5]][i])


# In[4]:


def insert_newline(ss):
    if not isinstance(ss,six.string_types):
        #print("Not string: {}".format(ss))
        return(ss)
    if len(ss) > 180:
        #print(ss)
        return(ss[:90] + "\n" + ss[90:180] + "\n" + ss[180:])
    elif len(ss) > 90:
        return(ss[:90] + "\n" + ss[90:])
    else:
        return(ss)

a = df.columns


for i in range(0,len(df.index)): 
    result = pd.DataFrame([['','']], columns = ['Heading','Details'])             

    for a in df: 
        column_name = a
        x = re.match('.*Topic.*',a)

        if x:
            s = re.search(r'.*?\[(.*)].*', a)
            ques = s.group(1)
            ques=ques.replace(';',':')

            #print (ques)
            lis = ques.split(":")
            #print("{}*****{}**** {} = {}".format(lis[5],lis[3], lis[1],df[column_name][0]))
            result = result.append({'Heading':'Topic','Details':insert_newline(lis[5])},True)
            if (len(lis[3]) != 1):
                result = result.append({'Heading':'Category','Details':lis[3]},True)
            result = result.append({'Heading':'Question','Details':insert_newline(lis[1])},True)
            result = result.append({'Heading':'Answer','Details':insert_newline(df[column_name][i])},True)
            result = result.append({'Heading':'','Details':''},True)
        else:
            x = re.match('^.Q.*_',a)
            if x:
                b = a
                lis = b.split(".",1)
                result = result.append({'Heading':'Question','Details':insert_newline(lis[1])},True)
                result = result.append({'Heading':'Answer','Details':insert_newline(df[column_name][i])},True)
                result = result.append({'Heading':'','Details':''},True)




    
    elements = []
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate("SurveyReport_{}_{}.pdf".format(i,df[df.columns[5]][i]), id='landscape', pagesize=landscape(letter), rightMargin=100,leftMargin=100, topMargin=30,bottomMargin=18)
    elements.append(Paragraph("Survey Report", styles['Title']))
    elements.append(Paragraph("Organization Name - {}".format(df[df.columns[5]][i]), styles['Title']))

    lista = [result.columns[:,].values.tolist()] + result.values.tolist()

    t=Table(lista, colWidths=[1*inch,7*inch])
    t.setStyle(TableStyle([('ALIGN',(1,1),(-2,-2),'RIGHT'),
                            ('TEXTCOLOR',(1,1),(-2,-2),colors.red),
                            ('VALIGN',(0,0),(0,-1),'TOP'),
                            ('TEXTCOLOR',(0,0),(0,-1),colors.blue),
                            ('ALIGN',(0,-1),(-1,-1),'CENTER'),
                            ('VALIGN',(0,-1),(-1,-1),'MIDDLE'),
                            ('TEXTCOLOR',(0,-1),(-1,-1),colors.green),
                            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                            ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                            ]))

                       
    elements.append(t)
    doc.build(elements)

print("Processing done.. Please look for files named SurveyReport_<customer name>.pdf in the same directory")


# In[ ]:





# In[ ]:




