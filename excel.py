#Created by Eric Buchanan
# Objective: Open Excel file, take data from two specific sheets and compare that data. Once a match is found take address for matching query and assign it to persons address.
from openpyxl import Workbook, load_workbook

#Load Workbook into wb
wb = load_workbook('document.xlsx')

matchFirst = []
matchLast = []
practiceName = []
practiceAddress = []
foundAddress = []
tempFirst = ''
tempLast = ''

#declares and assigns variables for indivdual sheets
matchData = wb['Notification']
practiceWB = wb['Practice']

#For loops to take data from each sheet and row and append array for every new variable. 
for row  in matchData.iter_rows(min_row=2,min_col=3,max_row=12997, max_col=3):
        for cell in row:
            matchFirst.append(cell.value)

#debug statement to check specific entry after assigning
#print(matchFirst[0])

for row  in matchData.iter_rows(min_row=2,min_col=5,max_row=12997, max_col=5):
        for cell in row:
            matchLast.append(cell.value)

for row  in practiceWB.iter_rows(min_row=5,min_col=4,max_row=12736, max_col=4):
        for cell in row:
            if cell.value == None:
                practiceName.append("fhueabfghebahgi")
            else:
                practiceName.append(cell.value)      


for row  in practiceWB.iter_rows(min_row=5,min_col=20,max_row=12736, max_col=20):
        for cell in row:
            if cell.value == None:
                practiceAddress.append("not found")  
            else:
                practiceAddress.append(cell.value)            
#debug statement to check specific entry after assigning
#print(practiceName[0])

#Ensure lengths stay constant
matchLength = len(matchFirst)
practiceLength = len(practiceName)

# For loop increment i till it equals matchLength set tempFirst and tempLast
for i in range(matchLength):
    tempFirst = matchFirst[i].lower()
    tempLast = matchLast[i].lower()
    tempAddress = 'notfound'
    #print (tempFirst,tempLast) debug statement to check getting assigned properly
    #  for loop increment J until it equals practiceLength set tempName = praciceName.lower() to ensure no issues with matching names. 
    for j in range(practiceLength):
        tempName = practiceName[j].lower()
        if (tempFirst in tempName and tempLast in tempName):        #if tempFirst and tempLast are within tempName then take the address from practiceAddress[j] (current row through scan for matching name) and append it to foundAddress
            tempAddress = practiceAddress[j]
            foundAddress.append(practiceAddress[j])     
            break 
        else:
            continue
    if tempAddress == 'notfound':   # if tempAddress is equal to not found then put in foundAddress list as Not Found
        foundAddress.append("Not found")

#Sets matchData as active sheet
wb.active = matchData
ws = wb.active

#for Loop to go down the foundAddress list and writes it to column24
for i, address in enumerate(foundAddress, start=2):
    ws.cell(row=i, column=24, value=address)

#debug statements
#print (len(foundAddress),matchLength,practiceLength)    
#print(foundAddress)


#save file whenever we get to that       
wb.save('balances.xlsx')