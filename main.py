# Working with a private spreadsheet in Google Sheets
# You need a Google Workbook in your google drive
# Login to console.developers.google.com
# Create Project and name it
#  Then Enable APIs and Services
# Search for and enable Google Drive API
#  Create Credentials for Google Drive API
#   Check Application Data
#   Fill in a Service Account Name
#   Grant the Project role
#   Create Keys in JSON file and download
# Upload to the repl (or wherever you are working with python)
# rename the file to secrets.json
# In that file copy the Client_email string
# 
# Back in the Google console, click on the three lines menu
# Select APIs and Services
# Click on Enable APIs and Services
# Search for, and enable, google sheets api
#
# Go to the workbook, and share with edit permissions
# For the email address, paste the client_email string from the JSON file, and send
# 

import gspread
import re
import statistics
import time

gc=gspread.service_account('secrets.json')

spreadsheet=gc.open('Weather')
print(spreadsheet)

# Get a worksheet by index
worksheet1=spreadsheet.get_worksheet(0)
print(worksheet1)
data=worksheet1.get_all_records()
print(data[5])

# Get a worksheet by name
worksheet2=spreadsheet.worksheet('2014')
print(worksheet2)
data=worksheet2.get_all_records()
print(data[3:5])

# Get a list of values in a range of cells
data=worksheet2.get_values('A5:E6')
print(data)

# Get a list of values in a column by colimn index
column=worksheet1.col_values(5)  #Col index starts at 1
column=worksheet1.col_values(5)[1:] # without header
print(column)

# Get row values by index
row=worksheet1.row_values(7)   #Row index starts at 1
print(row)

# Get the value in a cell
cell=worksheet1.acell('E7').value
print(type(cell),cell)

# Find a cell value
cell=worksheet1.find('-10')
print(cell)
print(cell.col, cell.row, cell.value)

# Find all cells with a value
cells=worksheet1.findall('-9')
print(cells)

# Find all cells with a partial value
reg=re.compile(r'996')
cells=worksheet1.findall(reg)
print(cells)
for cell in cells:
  print(cell.row, cell.col)

# Update a cell
worksheet1.update('E5', -29)

# Update a cell by cell co-ordinates
worksheet1.update_cell(6,5,-31)
print(worksheet1.get_values('E5:E6'))

# Add and update an entire column
# Convert temperature values to Farenheit
existingCol=worksheet1.col_values(5)[1:]
newCol=[[round((float(i)*9/5+32),1)] for i in existingCol]
worksheet1.add_cols(1)
worksheet1.update('G1:G25',[['Farenheit']]+newCol)

# Calculate the mean average for the Farenheit column
# Find the Farenheit column
fcol=worksheet1.find('Farenheit').col

# Get the column values
FarenheitCol=[float(i) for i in worksheet1.col_values(fcol)[1:]]

# Get the last row
lastrow=len(FarenheitCol)+1  #+1 because we stripped the header

# Calculate the mean average
meanAv=statistics.mean(FarenheitCol)

# Add the mean to the column
worksheet1.update_cell(lastrow+1,fcol,meanAv)

# Listen for a cell to change
value1=worksheet1.acell('G26').value
value2=value1

while value1==value2:
  time.sleep(5)
  value2=worksheet1.acell('G26').value

print("Changed")

