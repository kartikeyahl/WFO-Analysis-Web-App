# WFO-Analysis-Web-App

At Honda, divisions are required to maintain a 3:1 WFO:WFH ratio for their associates/employees, meaning 75% WFO and 25% WFH. 
WFO Analysis project will be a web application that takes three Excel files as input and provides visualizations and tabular reports on the WFO percentage completed within each operation, division, department, and designation. 
This will help the organization keep track of who has complied with the policy, and alert those who are not complying.


## Input Files: 
### 1.	Master data: Contains employee’s information
![image](https://user-images.githubusercontent.com/43701324/226509394-e1b4e863-ca17-4c29-ba7e-26a821877671.png)

### 2.	Punch file: Punch in/out details
![image](https://user-images.githubusercontent.com/43701324/226509447-ab0fcede-13a6-408c-8bce-e61adadc10fd.png)

### 3.	HMSI Calendar2023: 
![image](https://user-images.githubusercontent.com/43701324/226509497-f680dd8a-cb03-41cd-86cd-19f0155191f6.png)



## Steps and logic:
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
![image](https://user-images.githubusercontent.com/43701324/226509522-879dc50d-bb81-48df-af3b-b8657c62910e.png)
![image](https://user-images.githubusercontent.com/43701324/226509539-9dc0721e-e241-4faa-8c7b-de7314427785.png)
9.	Auto generate two files:<br />
&nbsp;&nbsp; a.	WFO_Analysis: Analysis in Tabular form.<br />
&nbsp;&nbsp; b.	WFO_Visualization: representation of analysis in Bar graphs.

10.	In addition, to generating the above files, redirect to analytics dashboard with bar graphs. 



## Technology used: 
python, Django and Streamlit

## Platform: 
VS Code, GitHub, and AWS.
