#!/usr/bin/env python
# coding: utf-8

# In[5]:


import pandas as pd
import re
import os.path
from pathlib import Path

print("This program is designed to create Environmental Discovery answers as input to the SDA HLD tool")
print("The output will be files named as SDA_HLD_input_<customer name>.pdf for each customer")
file = input("Please enter the full path and name of survey results xls file:\n")
#df = pd.read_excel('Common DNAC Land and Ad opt - Environment Discovery - Responses -  2020-7-8 13-48 26996.xlsx')

my_file = Path(file)  
if not os.path.isfile(my_file):
    print("The file does not exist. Easiest way is to copy the xlsx file into the same directory as the exe")
    quit()
else:
    df = pd.read_excel(file)


# In[6]:


from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, inch, A4


# In[7]:


def isyes(b):
    if (b >= 0):
        return ("Yes")
    else:
        return ("No")

a = df.columns

for i in range(0,len(df.index)): 
    result = pd.DataFrame([['','']], columns = ['HLD', 'Answer']) 


    result = result.append({'HLD':"List the total number of concurrent wired endpoints (excluding guests)",
                            'Answer':df[a[28]][i]}, True)
    result = result.append({'HLD':"Are any wired endpoints connected to IP phones",
                            'Answer':isyes(df[a[32]][i])}, True)
    result = result.append({'HLD':"List the total number of such endpoints",
                            'Answer':df[a[32]][i]}, True)
    result = result.append({'HLD':"Total number of concurrent wireless endpoints (excluding guests)",
                            'Answer':df[a[29]][i]}, True)

    result = result.append({'HLD':"Will guest be utilized",
                            'Answer':isyes(df[a[31]][i])}, True)
    result = result.append({'HLD':"List the total guest wired endpoints",
                            'Answer':df[a[30]][i]}, True)
    result = result.append({'HLD':"list the total guest wireless endpoints",
                            'Answer':df[a[31]][i]}, True)
    result = result.append({'HLD':"Is there a requirement for the guest traffic to be completely isolated \nfrom corporate network",
                            'Answer':df[a[77]][i]}, True)

    result = result.append({'HLD':"Is there a requirement for industrial standard swithces that can \nwithstand high temp and altitutes",
                            'Answer':isyes(df[a[33]][i])}, True)
    result = result.append({'HLD':"Requirement for east-west policy enforcement for these switches",
                            'Answer':"yes"}, True)
    result = result.append({'HLD':"How many end hosts will connect to these switches",
                            'Answer':df[a[33]][i]}, True)
    result = result.append({'HLD':"Is there a required for small form factor fanless switches or UPOE \nswtiches deisgn for smart buildings",
                            'Answer':isyes(df[a[34]][i]>0)}, True)
    result = result.append({'HLD':"Total number of end points connecing behind the fanless/UPOE switches",
                            'Answer':df[a[34]][i]}, True)


    result = result.append({'HLD':"Add endpoint type and if it needs broadcast",
                            'Answer': "IP Phones-No, \nCamaras-No, \nEmp Laptops-{}, \nEmp BYODs-No, \nGuest Laptops/BYODs-No".format(df[a[56]][i])}, True)


    result = result.append({'HLD':"Do you know the total number of VRFs, segments and endpoint user groups \nwhich will be configured in the network",
                            'Answer':"RAT"}, True)
    result = result.append({'HLD':"Total number of VNs",
                            'Answer':"RAT"}, True)
    result = result.append({'HLD':"How many VNs will utilize multicast",
                            'Answer':"RAT"}, True)
                        

    result = result.append({'HLD':"Total number of IP pools needed for traffic separation",
                            'Answer':df[a[68]][i]}, True)
    result = result.append({'HLD':"Number of IP pools for broadcast",
                            'Answer':df[a[69]][i]}, True)
    result = result.append({'HLD':"Total number of SGTs to segment endpoint groups",
                            'Answer':df[a[71]][i]}, True)
    result = result.append({'HLD':"Do any of these SGTs have static IP address that cannot be changed \nwhen migrating to fabric",
                            'Answer':isyes(df[a[72]][i])}, True)

    result = result.append({'HLD':"Is the IP address scheme v4, v6 or dual",
                            'Answer':df[a[70]][i]}, True)
    result = result.append({'HLD':"Do network devices need IPv6 or dual stack",
                            'Answer':df[a[80]][i]}, True)

    result = result.append({'HLD':"Is there a requiremnt for multicast in the network",
                            'Answer':df[a[55]][i]}, True)
    result = result.append({'HLD':"How many multicast groups have senders or recievers in the fabric",
                            'Answer': "IP Phones-No, \nCamaras-No, \nEmp Laptops-{}, \nEmp BYODs-No, \nGuest Laptops/BYODs-No".format(df[a[55]][i])}, True)

                        
    result = result.append({'HLD':"Do fabric hosts have a requirment to share subnets with non fabric \nhosts how many such subnets exists ?",
                            'Answer': "IP Phones-{}, \nCamaras-{}, \nIoT-{}, \nEmp Laptops-{}, \nEmp BYODs-{}, \nGuest Laptops/BYODs-{}".format(df[a[49]][i],df[a[50]][i],df[a[51]][i],df[a[52]][i],df[a[53]][i],df[a[54]][i])}, True)

                        
    result = result.append({'HLD':"For each subnet sharing IP pool with non-fabric hosts provide \nthe total number of hosts that sits inside and outside fabric",
                            'Answer':df[a[35]][i]}, True)

    result = result.append({'HLD':"Is there a requirement for any of the end points to have their gateway \non a firewall of outside the fabric gateway on a firewall of outside the fabric",
                            'Answer': "Emp Laptops-{}, \nEmp BYODs-{}, \nGuest Laptops-{}, \nGuest BYODs-{}".format(df[a[61]][i],df[a[62]][i],df[a[63]][i],df[a[64]][i])}, True)



    result = result.append({'HLD':"Does the network need to be highly resilient",
                            'Answer':"Yes"}, True)
   
    result = result.append({'HLD':"Is there a requirment to monitor network traffic in real time \nfor cloud based threat detection",
                            'Answer':df[a[74]][i]}, True)
    result = result.append({'HLD':"Does the site have single or multiple exits for network traffic",
                        'Answer':df[a[65]][i]}, True)
    result = result.append({'HLD':"Is there a requiremnt for site to be able to handle communication \nbetween hosts in different VNs",
                        'Answer':df[a[82]][i]}, True)
    result = result.append({'HLD':"",'Answer':""}, True)
    
    elements = []
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate("SDA_HLD_input_{}.pdf".format(df[a[5]][i]), pagesize=landscape(letter), rightMargin=100,leftMargin=100, topMargin=30,bottomMargin=18)
    elements.append(Paragraph("SDA HLD Tool Input from Survey", styles['Title']))
    elements.append(Paragraph("Organization Name - {}".format(df[a[5]][i]), styles['Title']))

    lista = [result.columns[:,].values.astype(str).tolist()] + result.values.tolist()


    t=Table(lista, colWidths=[6*inch, 2*inch])
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

print("Processing done.. Please look for files named as SDA_HLD_input_<customer name>.pdf for each customer")


# In[ ]:




