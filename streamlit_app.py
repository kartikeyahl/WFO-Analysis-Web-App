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
global dataset 
dataset = pd.read_excel(uploaded_file,parse_dates=['Date'],date_parser=lambda x: pd.to_datetime(x, format='%Y%m%d'))
# Create sidebar widgets for "From" and "To" dates
st.sidebar.header("Please filter here:")
from_date = st.sidebar.date_input("From Date", value=dataset['Date'].min())
to_date = st.sidebar.date_input("To Date", value=dataset['Date'].max())
# Convert the date input values to datetime objects
from_date = pd.to_datetime(from_date)
to_date = pd.to_datetime(to_date)
dataset = dataset[(dataset['Date'] >= from_date) & (dataset['Date'] <= to_date)]
st.markdown("---")

# If a file was uploaded
def wfo(dataset):
    if uploaded_file is not None:
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

        result_who = pd.DataFrame(ll1, columns=['E_Code','Full Name','Designation','Location','Operation','Division','Department'])
        result_wfo =  pd.concat([result_who[['E_Code','Full Name','Designation','Location','Operation','Division','Department']],dataset['counts']], axis=1)

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
        result_wfo_associate =  pd.concat([result_who[['E_Code','Full Name','Designation','Location','Operation','Division','Department']],result_who_final['Percent']], axis=1)

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
      
        return result_wfo_associate,result_wfo_operation,result_wfo_division,result_wfo_location,result_wfo_designation,result_wfo_department,working_days,dataset2
   

if uploaded_file is not None and uploaded_file2 is not None and uploaded_file3 is not None:
    wfo1 = st.cache_data(wfo)
    result_wfo_associate,result_wfo_operation,result_wfo_division,result_wfo_location,result_wfo_designation,result_wfo_department,working_days,dataset2 = wfo1(dataset)

else:
    result_wfo_associate,result_wfo_operation,result_wfo_division,result_wfo_location,result_wfo_designation,result_wfo_department,working_days,dataset2 = wfo(dataset)
    

# Define the table styles
table_styles = [
    {
        'selector': 'table',
        'props': [
            ('border-collapse', 'collapse'),
            ('border', '2px solid #ddd')
        ]
    },
    {
        'selector': 'th, td',
        'props': [
            ('padding', '8px'),
            ('border', '1px solid #ddd')
        ]
    },
    {
        'selector': 'th',
        'props': [
            ('background-color', '#eee')
        ]
    }
]

search_value = st.sidebar.number_input("Search by E.Code:", step=1)

option = st.selectbox(
    'Sort Division on basis of:',
    ('Operation', 'Location')
)

if option == 'Operation' or option==None :
# r3=result_wfo_associate.groupby('Operation')['Division'].unique().apply(list).to_dict()
    r3 = (
        result_wfo_associate.groupby('Operation')
        .apply(lambda x: x.groupby('Division')['Department'].unique().apply(list).to_dict())
        .to_dict()
    )

    operation=st.sidebar.multiselect(
    "Select Operation:",
    options=result_wfo_operation["Operation"].unique(),
    default=result_wfo_operation["Operation"].unique().tolist()
    )

    divi = []
    for o in operation:
        for i in r3[o]:
            divi.append(i)
    divi=[x for x in divi if not isinstance(x, float) or not np.isnan(x)]
    division=st.sidebar.multiselect(
    "Select Division:",
    options= result_wfo_division["Division"].unique(),
    default=divi
    )

    depart = []
    for o in operation:
        for d in r3[o]:
            for i in r3[o][d]:
                depart.append(i)
    depart=[x for x in depart if not isinstance(x, float) or not np.isnan(x)]
    department=st.sidebar.multiselect(
    "Select Department:",
    options=result_wfo_department["Department"].unique(),
    default=depart
    )

    mask = result_wfo_associate['Department'].isin(depart)
    r4 = result_wfo_associate[mask].groupby('Department')['Designation'].unique().apply(list).to_dict()

    des = []
    for i in r4.values():
        for j in i:
            des.append(j)   
    des=list(set([x for x in des if not isinstance(x, float) or not np.isnan(x)]))
    designation=st.sidebar.multiselect(
    "Select Designation:",
    options=result_wfo_designation["Designation"].unique(),
    default=des
    )

    location = st.sidebar.multiselect(
        "Select Location:",
        options=result_wfo_location["Location"].unique(),
        default=result_wfo_location["Location"].unique().tolist()
    )


if option == 'Location':
    r3 = (
        result_wfo_associate.groupby('Location')
        .apply(lambda x: x.groupby('Division')['Department'].unique().apply(list).to_dict())
        .to_dict()
    )
    
    location = st.sidebar.multiselect(
        "Select Location:",
        options=result_wfo_location["Location"].unique(),
        default=result_wfo_location["Location"].unique().tolist()
    )

    operation=st.sidebar.multiselect(
    "Select Operation:",
    options=result_wfo_operation["Operation"].unique(),
    default=result_wfo_operation["Operation"].unique().tolist()
    )

    divi = []
    for o in location:
        for i in r3[o]:
            divi.append(i)
    divi=[x for x in divi if not isinstance(x, float) or not np.isnan(x)]
    division=st.sidebar.multiselect(
    "Select Division:",
    options= result_wfo_division["Division"].unique(),
    default=divi
    )

    depart = []
    for o in location:
        for d in r3[o]:
            for i in r3[o][d]:
                depart.append(i)
    depart=[x for x in depart if not isinstance(x, float) or not np.isnan(x)]
    department=st.sidebar.multiselect(
    "Select Department:",
    options=result_wfo_department["Department"].unique(),
    default=depart
    )

    mask = result_wfo_associate['Department'].isin(depart)
    r4 = result_wfo_associate[mask].groupby('Department')['Designation'].unique().apply(list).to_dict()

    des = []
    for i in r4.values():
        for j in i:
            des.append(j)   
    des=list(set([x for x in des if not isinstance(x, float) or not np.isnan(x)]))
    designation=st.sidebar.multiselect(
    "Select Designation:",
    options=result_wfo_designation["Designation"].unique(),
    default=des
    )

df_selection = result_wfo_operation.query("Operation == @operation")
df_selection2 = result_wfo_division.query("Division == @division")
df_selection3 = result_wfo_department.query("Department == @department")
df_selection4 = result_wfo_designation.query("Designation == @designation")
df_selection5 = result_wfo_location.query("Location == @location")
df_selection6 = result_wfo_associate.query("E_Code == @search_value")
df_selection7 = result_wfo_associate.query("Department == @department")

if result_wfo_operation is not None:
    result_wfo_dess = df_selection7.groupby('Designation',as_index=False).count()
    result_wfo_dess=result_wfo_dess.drop(['E_Code','Full Name','Location','Operation','Division','Department'], axis=1)
    result_wfo_dess.rename(columns={'Percent': 'Counts'}, inplace=True)
    
    department_list = df_selection3['Department'].tolist()
    filtered_data = dataset2[dataset2['Department'].isin(department_list)]
    result_wfo_dess_manpower = filtered_data.groupby(['Designation']).size().reset_index(name='des Manpower')
    result_wfo_dess_manpower=result_wfo_dess_manpower.set_index('Designation').T.to_dict('list')

    l1 = [i for i in range(len(result_wfo_dess)) if result_wfo_dess.iloc[i]['Designation']=='NaN' or result_wfo_dess.iloc[i]['Designation']=='nan']
    result_wfo_dess.drop(l1,axis=0,inplace=True)
    result_wfo_dess.reset_index(inplace=True)
    percent_wfo_dess=[]
    for i in range(len(result_wfo_dess)):
        if result_wfo_dess.iloc[i]['Designation']!='NaN':
            xx=(result_wfo_dess.iloc[i]['Counts']/(result_wfo_dess_manpower[result_wfo_dess.iloc[i]['Designation']][0]*working_days))*100
            percent_wfo_dess.append(xx)  

    percent_wfo_dess2 = [round(i,1) for i in percent_wfo_dess]

    result_who_finall4 = pd.DataFrame(percent_wfo_dess2, columns=['Percent'])
    result_wfo_designationn =  pd.concat([result_wfo_dess['Designation'],result_who_finall4['Percent']], axis=1)


if df_selection6.empty:
    pass
else:
    st.table(df_selection6.style.set_table_styles(table_styles))
    st.markdown("---")

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
# Filter the data to include only rows with Percent above 75
filtered_data1 = df_selection[df_selection['Percent'] > 75]
filtered_data2 = result_wfo_designationn[result_wfo_designationn['Percent'] > 75]
filtered_data3 = df_selection2[df_selection2['Percent'] > 75]
filtered_data4 = df_selection3[df_selection3['Percent'] > 75]

toggle_button = st.checkbox("Greater than 75%")
if toggle_button:
    # create an Altair bar chart
    bars_opr = alt.Chart(filtered_data1).mark_bar().encode(
        x=alt.X('Operation', axis=alt.Axis(labelAngle=-90)),
        y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
        color=alt.condition(
            alt.datum.Percent >= 75,
            alt.value('red'),  # bars above 75% will be red
            alt.value('green')  # bars below 75% will be green
        )
    )

    # create an Altair bar chart
    bars_des = alt.Chart(filtered_data2).mark_bar().encode(
        x=alt.X('Designation', axis=alt.Axis(labelAngle=-90)),
        y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
        color=alt.condition(
            alt.datum.Percent >= 75,
            alt.value('red'),  # bars above 75% will be red
            alt.value('green')  # bars below 75% will be green
        )
    )

    # create an Altair bar chart
    bars_div = alt.Chart(filtered_data3).mark_bar().encode(
        x=alt.X('Division', axis=alt.Axis(labelAngle=-90)),
        y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
        color=alt.condition(
            alt.datum.Percent >= 75,
            alt.value('red'),  # bars above 75% will be red
            alt.value('green')  # bars below 75% will be green
        )
    )

    # create an Altair bar chart
    bars_dept = alt.Chart(filtered_data4).mark_bar().encode(
        x=alt.X('Department', axis=alt.Axis(labelAngle=-90)),
        y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
        color=alt.condition(
            alt.datum.Percent >= 75,
            alt.value('red'),  # bars above 75% will be red
            alt.value('green')  # bars below 75% will be green
        )
    )

else:
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
    bars_des = alt.Chart(result_wfo_designationn).mark_bar().encode(
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

    # create an Altair bar chart
    bars_dept = alt.Chart(df_selection3).mark_bar().encode(
        x=alt.X('Department', axis=alt.Axis(labelAngle=-90)),
        y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
        color=alt.condition(
            alt.datum.Percent >= 75,
            alt.value('red'),  # bars above 75% will be red
            alt.value('green')  # bars below 75% will be green
        )
    )

# Create an Altair text chart for value labels
def label_fun(df,x):
    labels_dept = alt.Chart(df).mark_text(
        align='center',
        baseline='middle',
        dy=-5,  # Adjust vertical position of the labels
    ).encode(
        x=alt.X(x, axis=alt.Axis(labelAngle=-90)),
        y=alt.Y('Percent', axis=alt.Axis(title='Percentage')),
        text=alt.Text('Percent:Q', format='.1f'),  # Display bar values with 2 decimal places
    )
    return labels_dept

# add a red dotted line at 75%
rule = alt.Chart(pd.DataFrame({'y': [75]})).mark_rule(color='red', strokeDash=[3,3]).encode(y='y')

# combine the bar chart and the rule
if toggle_button:
    chart_opr = (bars_opr +label_fun(filtered_data1,'Operation')+ rule).properties(width=alt.Step(40))
    chart_div = (bars_div +label_fun(filtered_data3,'Division')+ rule).properties(width=alt.Step(40))
    chart_des = (bars_des +label_fun(filtered_data2,'Designation')+ rule).properties(width=alt.Step(40))
    chart_loc = (bars_loc +label_fun(df_selection5,'Location')+ rule).properties(width=alt.Step(40))
    chart_dept = (bars_dept + label_fun(filtered_data4,'Department')+ rule).properties(width=alt.Step(40))

else:
    chart_opr = (bars_opr +label_fun(df_selection,'Operation')+ rule).properties(width=alt.Step(40))
    chart_div = (bars_div +label_fun(df_selection2,'Division')+ rule).properties(width=alt.Step(40))
    chart_des = (bars_des +label_fun(result_wfo_designationn,'Designation')+ rule).properties(width=alt.Step(40))
    chart_loc = (bars_loc +label_fun(df_selection5,'Location')+ rule).properties(width=alt.Step(40))
    chart_dept = (bars_dept + label_fun(df_selection3,'Department')+ rule).properties(width=alt.Step(40))




# Define the legend
legend = alt.Chart(pd.DataFrame({'labels': ['<75%', '>=75%']})).mark_rect().encode(
    y=alt.Y('labels:N', axis=None),
    color=alt.Color('labels:N', scale=color_scale)
)


# Show the chart and the legend
st.altair_chart(chart_opr | legend)
st.altair_chart(chart_div | legend)
st.altair_chart(chart_dept | legend)
st.altair_chart(chart_des | legend)
st.altair_chart(chart_loc | legend)
