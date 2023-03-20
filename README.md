# WFO-Analysis-Web-App

At Honda, divisions are required to maintain a 3:1 WFO:WFH ratio for their associates/employees, meaning 75% WFO and 25% WFH. 
WFO Analysis project will be a web application that takes three Excel files as input and provides visualizations and tabular reports on the WFO percentage completed within each operation, division, department, and designation. 
This will help the organization keep track of who has complied with the policy, and alert those who are not complying.


Input Files: 
1.	Master data: Contains employee’s information
Full Name	Father Name	Sex	Location	Status	Designation	Operation	Division	Department	Section	Section Code	DOJ

2.	Punch file: Punch in/out details
E.Code	IN/OUT	MID	Date	Time	Location

3.	HMSI Calendar2023: 
MONTHDATEYEAR	DATEDAY	ISWORKINGDAY


Steps and logic:
1.	Importing files:
a.	Month data (compiled punch file)
b.	Master data (containing associates’ details)
c.	HMSI calendar location wise. 
2.	Dropping rows with 
a.	IN/OUT as 1 and P20.
b.	Status as ‘withdrawn’ in master data file
3.	Data cleaning (removing missing values, redundant columns/rows, etc.)
4.	Mapping compiled punch file with master data to get associates details like name, designation, department, division, operation, and location as data-frame.
5.	Group latest data-frame by designation, department, division, operation, and location.
6.	Group latest data-frame by designation, department, division, operation, and location to get count of operations, designation, etc., occurring during that period.
7.	Group master data by designation, department, division, operation, and location to get count of operations, designation, etc., existing.
8.	Calculating WFO% (refer below screen shots):

 


 
9.	Auto generate two files:
a.	WFO_Analysis: Analysis in Tabular form.
b.	WFO_Visualization: representation of analysis in Bar graphs.

10.	In addition, to generating the above files, redirect to analytics dashboard with bar graphs. 



Technology used: python, Django and Streamlit
Platform: VS Code, GitHub, and AWS.
