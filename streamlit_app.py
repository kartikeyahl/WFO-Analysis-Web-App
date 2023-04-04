import streamlit as st
import altair as alt
from openpyxl import Workbook
import xlsxwriter
import os
# Importing the libraries
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime,timedelta
import math
import concurrent.futures as ft
import multiprocessing
import random
import matplotlib.patches as mpatches
import matplotlib.lines as mlines

st.set_page_config(page_title="WFO App", layout="wide", initial_sidebar_state="expanded")

# Use Streamlit's file uploader to allow the user to select an Excel file
col1, col2, col3 = st.columns(3)
with col1:
    uploaded_file = st.file_uploader("Choose a Punch file", type="xlsx")
with col2:
    uploaded_file2 = st.file_uploader("Choose a master data file", type="xlsx")
with col3:
    uploaded_file3 = st.file_uploader("Choose a calendar file", type="xlsx")

st.markdown("---")

# If a file was uploaded
@st.cache_data
def wfo():
    if uploaded_file is not None:
        dataset = pd.read_excel(uploaded_file,parse_dates=['Date'],date_parser=lambda x: pd.to_datetime(x, format='%Y%m%d'))
        dataset2=pd.read_excel(uploaded_file2)
        dataset3=pd.read_excel(uploaded_file3)
        
        #Renaming column name to a valid column name format as per python
        dataset.rename(columns={'IN/OUT': 'IN_OUT'}, inplace=True)
        dataset3['MONTHDATEYEAR'] = pd.to_datetime(dataset3['MONTHDATEYEAR'])

        #Droping rows that conatain 1  and 'Withdrawn' in 'IN_OUT' and 'Status' column
        dataset=dataset[dataset.IN_OUT!=1]
        dataset=dataset[dataset.IN_OUT!='P20']
        dataset2=dataset2[dataset2.Status!='Withdrawn']

        month = (dataset.iloc[0]['Date']).month
        groupby_month = dataset3.groupby(pd.Grouper(key='MONTHDATEYEAR', freq='M'),as_index=False).sum()

        daydiff = (dataset.iloc[-1]['Date']).day-(dataset.iloc[0]['Date']).day+1

        start_date = pd.to_datetime(dataset.iloc[0]['Date'])
        end_date = pd.to_datetime(dataset.iloc[-1]['Date'])
        working_days = dataset3.loc[(dataset3['MONTHDATEYEAR'] >=start_date) & (dataset3['MONTHDATEYEAR'] <= end_date), 'ISWORKINGDAY'].sum()

        dataset.dropna(inplace=True)

        dataset = dataset.groupby(['E.Code']).size().reset_index(name='counts')

        ll1=[]
        def binarySearch(arr, l, r, x):

            while l <= r:

                mid = l + (r - l) // 2

                # Check if x is present at mid
                if arr[mid] == x:
                    return mid

                # If x is greater, ignore left half
                elif arr[mid] < x:
                    l = mid + 1

                # If x is smaller, ignore right half
                else:
                    r = mid - 1

            # If we reach here, then the element
            # was not present
            return -1


        # Driver Code
        arr = [int(i) for i in dataset2['E Code']]

        for i in range(len(dataset)):
        # Function call
            result = binarySearch(arr, 0, len(arr)-1, dataset.iloc[i]['E.Code'])

            if result != -1:
                ll1.append([dataset2.iloc[result]['E Code'],dataset2.iloc[result]['Full Name'], dataset2.iloc[result]['Designation'],dataset2.iloc[result]['Location'],dataset2.iloc[result]['Operation'],dataset2.iloc[result]['Division'],dataset2.iloc[result]['Department']])
            else:
                ll1.append(['NaN','NaN','NaN','NaN','NaN','NaN','NaN'])

        result_who = pd.DataFrame(ll1, columns=['E.Code','Full Name','Designation','Location','Operation','Division','Department'])
        result_wfo =  pd.concat([result_who[['E.Code','Full Name','Designation','Location','Operation','Division','Department']],dataset['counts']], axis=1)

        result_wfo_opr = result_wfo.groupby('Operation',as_index=False).sum()
        result_wfo_opr_manpower = dataset2.groupby(['Operation']).size().reset_index(name='opr Manpower')
        result_wfo_opr_manpower=result_wfo_opr_manpower.set_index('Operation').T.to_dict('list')

        result_wfo_div = result_wfo.groupby('Division',as_index=False).sum()
        result_wfo_div_manpower = dataset2.groupby(['Division']).size().reset_index(name='div Manpower')
        result_wfo_div_manpower=result_wfo_div_manpower.set_index('Division').T.to_dict('list')

        result_wfo_loc = result_wfo.groupby('Location',as_index=False).sum()
        result_wfo_loc_manpower = dataset2.groupby(['Location']).size().reset_index(name='loc Manpower')
        result_wfo_loc_manpower=result_wfo_loc_manpower.set_index('Location').T.to_dict('list')

        result_wfo_des = result_wfo.groupby('Designation',as_index=False).sum()
        result_wfo_des_manpower = dataset2.groupby(['Designation']).size().reset_index(name='des Manpower')
        result_wfo_des_manpower=result_wfo_des_manpower.set_index('Designation').T.to_dict('list')

        result_wfo_dept = result_wfo.groupby('Department',as_index=False).sum()
        result_wfo_dept_manpower = dataset2.groupby(['Department']).size().reset_index(name='dept Manpower')
        result_wfo_dept_manpower=result_wfo_dept_manpower.set_index('Department').T.to_dict('list')

        l1 = [i for i in range(len(result_wfo_opr)) if result_wfo_opr.iloc[i]['Operation']=='NaN' or result_wfo_opr.iloc[i]['Operation']=='nan']
        result_wfo_opr.drop(l1,axis=0,inplace=True)
        result_wfo_opr.reset_index(inplace=True)

        l1 = [i for i in range(len(result_wfo_div)) if result_wfo_div.iloc[i]['Division']=='NaN' or result_wfo_div.iloc[i]['Division']=='nan']
        result_wfo_div.drop(l1,axis=0,inplace=True)
        result_wfo_div.reset_index(inplace=True)

        l1 = [i for i in range(len(result_wfo_loc)) if result_wfo_loc.iloc[i]['Location']=='NaN' or result_wfo_loc.iloc[i]['Location']=='nan']
        result_wfo_loc.drop(l1,axis=0,inplace=True)
        result_wfo_loc.reset_index(inplace=True)

        l1 = [i for i in range(len(result_wfo_des)) if result_wfo_des.iloc[i]['Designation']=='NaN' or result_wfo_des.iloc[i]['Designation']=='nan']
        result_wfo_des.drop(l1,axis=0,inplace=True)
        result_wfo_des.reset_index(inplace=True)

        l1 = [i for i in range(len(result_wfo_dept)) if result_wfo_dept.iloc[i]['Department']=='NaN' or result_wfo_dept.iloc[i]['Department']=='nan']
        result_wfo_dept.drop(l1,axis=0,inplace=True)
        result_wfo_dept.reset_index(inplace=True)

        percent_wfo_div=[]
        for i in range(len(result_wfo_div)):
            if result_wfo_div.iloc[i]['Division']!='NaN':
                x=(result_wfo_div.iloc[i]['counts']/(result_wfo_div_manpower[result_wfo_div.iloc[i]['Division']][0]*working_days))*100
                percent_wfo_div.append(x)

        percent_wfo_opr=[]
        for i in range(len(result_wfo_opr)):
            if result_wfo_opr.iloc[i]['Operation']!='NaN':
                x=(result_wfo_opr.iloc[i]['counts']/(result_wfo_opr_manpower[result_wfo_opr.iloc[i]['Operation']][0]*working_days))*100
                percent_wfo_opr.append(x)    

        percent_wfo_loc=[]
        for i in range(len(result_wfo_loc)):
            if result_wfo_loc.iloc[i]['Location']!='NaN':
                x=(result_wfo_loc.iloc[i]['counts']/(result_wfo_loc_manpower[result_wfo_loc.iloc[i]['Location']][0]*working_days))*100
                percent_wfo_loc.append(x)

        percent_wfo_des=[]
        for i in range(len(result_wfo_des)):
            if result_wfo_des.iloc[i]['Designation']!='NaN':
                x=(result_wfo_des.iloc[i]['counts']/(result_wfo_des_manpower[result_wfo_des.iloc[i]['Designation']][0]*working_days))*100
                percent_wfo_des.append(x)  

        percent_wfo_dept=[]
        for i in range(len(result_wfo_dept)):
            if result_wfo_dept.iloc[i]['Department']!='NaN':
                x=(result_wfo_dept.iloc[i]['counts']/(result_wfo_dept_manpower[result_wfo_dept.iloc[i]['Department']][0]*working_days))*100
                percent_wfo_dept.append(x)  

        percent_wfo=[]
        for i in range(len(result_wfo)):
            percent_wfo.append((result_wfo.iloc[i]['counts']/working_days)*100)

        percent_wfo2=[round(i,1) for i in percent_wfo]
        percent_wfo_opr2 = [round(i,1) for i in percent_wfo_opr]
        percent_wfo_div2 = [round(i,1) for i in percent_wfo_div]
        percent_wfo_loc2 = [round(i,1) for i in percent_wfo_loc]
        percent_wfo_des2 = [round(i,1) for i in percent_wfo_des]
        percent_wfo_dept2 = [round(i,1) for i in percent_wfo_dept]

        result_who_final = pd.DataFrame(percent_wfo2, columns=['Percent'])
        result_wfo_associate =  pd.concat([result_who[['E.Code','Full Name','Designation','Location','Operation','Division','Department']],result_who_final['Percent']], axis=1)

        result_who_final2 = pd.DataFrame(percent_wfo_opr2, columns=['Percent'])
        result_wfo_operation =  pd.concat([result_wfo_opr['Operation'],result_who_final2['Percent']], axis=1)

        result_who_final3 = pd.DataFrame(percent_wfo_div2, columns=['Percent'])
        result_wfo_division =  pd.concat([result_wfo_div['Division'],result_who_final3['Percent']], axis=1)

        result_who_final4 = pd.DataFrame(percent_wfo_loc2, columns=['Percent'])
        result_wfo_location =  pd.concat([result_wfo_loc['Location'],result_who_final4['Percent']], axis=1)

        result_who_final4 = pd.DataFrame(percent_wfo_des2, columns=['Percent'])
        result_wfo_designation =  pd.concat([result_wfo_des['Designation'],result_who_final4['Percent']], axis=1)

        result_who_final5 = pd.DataFrame(percent_wfo_dept2, columns=['Percent'])
        result_wfo_department =  pd.concat([result_wfo_dept['Department'],result_who_final5['Percent']], axis=1)

        return result_wfo_associate,result_wfo_operation,result_wfo_division,result_wfo_location,result_wfo_designation,result_wfo_department
    
result_wfo_associate,result_wfo_operation,result_wfo_division,result_wfo_location,result_wfo_designation,result_wfo_department=wfo()


# clear_all = st.sidebar.checkbox("Clear filter")
# if clear_all:
#     operation = []  # clear operation multiselect
#     division = []  # clear division multiselect
#     department = []  # clear department multiselect
#     designation = []  # clear designation multiselect
#     location = []  # clear location multiselect


st.sidebar.header("Please filter here:")
opr = [
            "Brand & Communication",
            "Customer Service",
            "Finance & Accounts",
            "General and Corporate Affairs",
            "Honda India Foundation",
            "Internal Audit",
            "Logistics Planning & Control",
            "Overseas Business",
            "Premium Motorcycle Business",
            "Sales & Marketing",
            "Strategic Information System"
        ]

div = {"Brand & Communication":["Corporate Communications","Motor Sports","Safety Riding Promotion & Training"],
        "Customer Service":["After Sales Business","CS Field Service","CS Technical","CS Technology & Customer Relations", "Regional Business Central CS","Regional Business East CS","Regional Business South CS","Regional Business North CS","Regional Business West CS"],
        "Finance & Accounts":["Accounts","Finance","Taxation"],
        "General and Corporate Affairs":["Administration & Security","External Affairs","Health & Wellness","Human Resource","Industrial Relations","Legal & Secretarial"],
        "Honda India Foundation":["CSR Foundation"],
        "Internal Audit":["Internal Audit","Specialist IA"],
        "Logistics Planning & Control":["Commercial","Logistics"],
        "Overseas Business":["Export Packaging & Logistics","Export/Import Sales & After Sales","Export Planning & Control Function","Export Quality Control"],
        "Premium Motorcycle Business":["Customer Service","Sales"],
        "Sales & Marketing":["Marketing","Product Planning","Regional Business Central","Regional Business East","Regional Business South","Regional Business West","Sales & Business Planning","Sales Resource Quality & Training"],
        "Strategic Information System":["Enterprise IT Application","Enterprise IT Infrastructure","IT Compliance"]
}

dept={
    "Corporate Communications":["Corporate PR"],
    "Motor Sports":["Motorsports Planning & Promotion"],
    "Safety Riding Promotion & Training":["Safety Education & Centre Mgmt-S,W,C,E","Safety Education & Centre Mgmt-N & IDTR"],
    "After Sales Business":["Parts Planning","Parts Sales","Accessories & New Development"],
    "CS Field Service":["Network Development","Service Marketing","CS Information Management"],
    "CS Technical":["New Model Management","Market Quality Information","Warranty & FSC Management"],
    "CS Technology & Customer Relations":["Skill Enhancement","Complaint Handling","Contact Management"],
    "Regional Business Central CS":["After Sales Business Central 1","After Sales Business Central 2","MP East CS","U.P. East CS","U.P. Central CS","U.P. West CS","Chattisgarh CS","MP West CS"],
    "Regional Business East CS":["After Sales Business East 1","After Sales Business East 2","Orissa CS","North East CS","West Bengal CS","Bihar CS","Jharkhand CS"],
    "Regional Business North CS":["After Sales Business North 1","After Sales Business North 2","Rajasthan 1 CS","Rajasthan 2 CS","Haryana CS","Delhi CS","Punjab CS"],
    "Regional Business South CS":["After Sales Business South 1","After Sales Business South 2","TN North CS","TN South CS","Karnataka North CS","Karnataka South CS","Andhra Pradesh CS","Kerala CS","Telangana CS"],
    "Regional Business West CS":["After Sales Business West 1","After Sales Business West 2","Nagpur CS","Mumbai/Goa CS","Gujrat 1 CS","Gujrat 2 CS","Pune CS"],
    "Accounts":["Business Planning","Financial Accounting","Import & Export Banking","Sale Accounting"],
    "Finance":["Banking & Insurance","Payable Management & System Automation"],
    "Taxation":["Corporate Tax","GST & Tax Litigation","Tax Compliance Assurance"],
    "Administration & Security":["Administration & Vigilance","Expat & Fleet Management"],
    "External Affairs":["Planning","Road Safety"],
    "Health & Wellness":[],
    "Human Resource":["Associate Lifecycle","Strategy & Planning"],
    "Industrial Relations":["Employee Benefit","IR & Legal Compliance","IR Establishment","Manpower Management"],
    "Legal & Secretarial":["Compliance & Governance","Legal"],
    "CSR Foundation":[],
    "Internal Audit":["Internal Audit"],
    "Specialist IA":[],
    "Commercial":["Corporate Commercial Planning","Plant Commercial 1F & 3F","Plant Commercial 2F & 4F"],
    "Logistics":["Budget, Audits & Projects","Domestic Logistics 1F","Domestic Logistics 2F","Domestic Logistics 3F","Domestic Logistics 4F","HO Corporate Logistics"],
    "Export Packaging & Logistics":["Export Logistics","Export Packing"],
    "Export/Import Sales & After Sales":["Export After Sales","Export Sales","Export/Import Management"],
    "Export Planning & Control Function":["Export Business Planning","Overseas Market Quality"],
    "Export Quality Control":["CBU Quality Control & CKD 3F/4F","CKD Quality Control 1F & 2F"],
    "Customer Service":["Planning & Field Service","Training & Technical Support"],
    "Sales":["Business Planning & Development","Customer Experience","Field Sales","Marketing"],
    "Enterprise IT Application":["Application Development","Application Support","Business Transformation & New Projects","Dealer & Ext System Dev & Support","Technical & Database Mgt"],
    "Enterprise IT Infrastructure":["Central Desk","Network & Security","System Infrastructure"],
    "IT Compliance":["ISMS & IT GRC","SOX/IFCR Compliance"]
}

operation=st.sidebar.multiselect(
"Select Operation:",
options=result_wfo_operation["Operation"].unique(),
default=result_wfo_operation["Operation"].unique().tolist()
)


divi = []
for o in operation:
    divi += div.get(o, [])
divi = list(set(divi))
division=st.sidebar.multiselect(
"Select Division:",
options=result_wfo_division["Division"].unique(),
default=divi
)

depart = []
for o in division:
    depart += dept.get(o, [])
depart = list(set(depart))
department=st.sidebar.multiselect(
"Select Department:",
options=result_wfo_department["Department"].unique(),
default=depart
)

designation=st.sidebar.multiselect(
"Select Designation:",
options=result_wfo_designation["Designation"].unique(),
default=result_wfo_designation["Designation"].unique().tolist()
)

location = st.sidebar.multiselect(
    "Select Location:",
    options=result_wfo_location["Location"].unique(),
    default=result_wfo_location["Location"].unique().tolist()
)

df_selection = result_wfo_operation.query("Operation == @operation")
df_selection2 = result_wfo_division.query("Division == @division")
df_selection3 = result_wfo_department.query("Department == @department")
df_selection4 = result_wfo_designation.query("Designation == @designation")
df_selection5 = result_wfo_location.query("Location == @location")


# Define the color scheme
color_scale = alt.Scale(domain=['<75%','>=75%'], range=['green','red'])

# create an Altair bar chart
bars_loc = alt.Chart(df_selection5).mark_bar().encode(
    x=alt.X('Location', axis=alt.Axis(labelAngle=-90)),
    y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
    color=alt.condition(
        alt.datum.Percent >= 75,
        alt.value('red'),  # bars above 75% will be red
        alt.value('green')  # bars below 75% will be green
    )
)

# create an Altair bar chart
bars_opr = alt.Chart(df_selection).mark_bar().encode(
    x=alt.X('Operation', axis=alt.Axis(labelAngle=-90)),
    y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
    color=alt.condition(
        alt.datum.Percent >= 75,
        alt.value('red'),  # bars above 75% will be red
        alt.value('green')  # bars below 75% will be green
    )
)

# create an Altair bar chart
bars_des = alt.Chart(df_selection4).mark_bar().encode(
    x=alt.X('Designation', axis=alt.Axis(labelAngle=-90)),
    y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
    color=alt.condition(
        alt.datum.Percent >= 75,
        alt.value('red'),  # bars above 75% will be red
        alt.value('green')  # bars below 75% will be green
    )
)

# create an Altair bar chart
bars_div = alt.Chart(df_selection2).mark_bar().encode(
    x=alt.X('Division', axis=alt.Axis(labelAngle=-90)),
    y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
    color=alt.condition(
        alt.datum.Percent >= 75,
        alt.value('red'),  # bars above 75% will be red
        alt.value('green')  # bars below 75% will be green
    )
)

# add a red dotted line at 75%
rule = alt.Chart(pd.DataFrame({'y': [75]})).mark_rule(color='red', strokeDash=[3,3]).encode(y='y')

# combine the bar chart and the rule
chart_opr = (bars_opr + rule).properties(width=alt.Step(40))
chart_div = (bars_div + rule).properties(width=alt.Step(40))
chart_des = (bars_des + rule).properties(width=alt.Step(40))
chart_loc = (bars_loc + rule).properties(width=alt.Step(40))



# Define the legend
legend = alt.Chart(pd.DataFrame({'labels': ['<75%', '>=75%']})).mark_rect().encode(
    y=alt.Y('labels:N', axis=None),
    color=alt.Color('labels:N', scale=color_scale)
)


# Show the chart and the legend
st.altair_chart(chart_opr | legend)
st.altair_chart(chart_div | legend)
st.altair_chart(chart_des | legend)
st.altair_chart(chart_loc | legend)


# st.altair_chart(chart_des | legend)
# st.altair_chart(chart_loc | legend)
