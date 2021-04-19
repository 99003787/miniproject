## Python_miniproject

## Introduction-

In this python project, we have to read multiple excel sheet data having 40 rows x 10 column that contain 3 similar columns in each sheet i.e. PS number, Name & Email address. Next, we have to take the user input for printing the desired student data in a master sheet. Here we are reading the multiple sheet and storing the similar student data into master sheet.


## Folder Structure
Folder             | Description
-------------------| -----------------------------------------
`1_Requirements`   | Documents detailing requirements and research
`2_Design`         | Documents specifying design details
`3_Implementation` | All code and documentation
`4_Test_plan`      | Documents with test plans and procedures

## Library required for running this project:

SLNo |	Library name	| Operation	Install Library Code
-----------------------------------------------------------------
1 |	Openpyxl	| Reading and writing excel sheet	pip install openpyxl
2 |	Pandas | To automate excel sheet	pip install pandas

## About the project
The aim of the project is to extract the data present in different spreadsheets in one excel file as required by the user by different paths given by him. The excel sheet scrolls through all the spreadsheets with the following data common in all the sheets:

* Name :
* Ps Number :
* Email id :

The user defines the data that needs to be searched on the basis of the common data. The python program then reads the data corresponding to the particular data from different spreadsheets of excel. It then creates a mastersheet and adds the data from all the sheets to it. In the end, the data to be provided to the user is printed to the console.



 ## Features that are integrated in this project are
 
* Reading multiple excel sheets each having 40 rows x 10 column
* Searching methods to search details for user input values
* Combining all the similar data in news master sheet

## Challenges faced and how were they overcome
* Implementing the library
* Writing data in master sheet
* Creating the database
* Reading the file path from multiple directory



