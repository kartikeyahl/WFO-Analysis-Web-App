# WFO-Analysis-Web-App

At Honda, divisions are required to maintain a 3:1 WFO:WFH ratio for their associates/employees, meaning 75% WFO and 25% WFH. 
The WFO Analysis dashboard is a web application that inputs three Excel files from the users and provides visualizations on the WFO percentage completed within each operation, division, department, and designation. 
This helps the organization keep track of who has complied with the policy, and alert those who are not complying.


## Input Files: 
### 1.	Master data: Contains employee’s information
![image](https://user-images.githubusercontent.com/43701324/226509394-e1b4e863-ca17-4c29-ba7e-26a821877671.png)

### 2.	Punch file: Punch in/out details
![image](https://user-images.githubusercontent.com/43701324/226511134-c97e22ee-9312-442e-8396-bc8a430df913.png)

### 3.	HMSI Calendar2023: 
![image](https://user-images.githubusercontent.com/43701324/226511180-b65552ee-c14c-4e84-9f57-50e0b3233582.png)



## Steps and logic:
1.	Importing files:<br />
&nbsp;&nbsp; a.	Month data (compiled punch file)<br />
&nbsp;&nbsp; b.	Master data (containing associates’ details)<br />
&nbsp;&nbsp; c.	HMSI calendar location wise. 
2.	Dropping rows with <br />
&nbsp;&nbsp; a.	IN/OUT as 1 and P20.<br />
&nbsp;&nbsp; b.	Status as ‘withdrawn’ in master data file
3.	Data cleaning (removing missing values, redundant columns/rows, etc.)
4.	Mapping compiled punch file with master data to get associates details like name, designation, department, division, operation, and location as data-frame.
5.	Group latest data-frame by designation, department, division, operation, and location.
6.	Group latest data-frame by designation, department, division, operation, and location to get count of operations, designation, etc., occurring during that period.
7.	Group master data by designation, department, division, operation, and location to get count of operations, designation, etc., existing.
8.	Calculating WFO% (refer below screen shots):
![image](https://user-images.githubusercontent.com/43701324/226509522-879dc50d-bb81-48df-af3b-b8657c62910e.png)
![image](https://user-images.githubusercontent.com/43701324/226509539-9dc0721e-e241-4faa-8c7b-de7314427785.png)
9.	Analytics dashboard with bar graphs.



## Technology used: 
python -> Backend logic development
Streamlit -> Frontend, packing it with backend code to create a single file (streamlit_app.py)
Streamlit Cloud -> Hosting
GitHub -> File storage and version control

## Platform: 
VS Code, GitHub, and Streamlit Cloud.
