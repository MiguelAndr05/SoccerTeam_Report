# Miguel Andrade
# This code checks wether a players parent hasn't paid and how much they they need to pay
# and put it into a file to so the team manager is able to email the parents easier
# aswell as some quality of life features for the manager
# Excel file is chatGpt data generated
import subprocess
import openpyxl
import os
import pyinputplus as pyip
import time


# variables for removing hard coding elements
playerNameCol =  'a' #columns for players name
parentNameCol = 'b' # column for parents name
balRemainingCol = 'f'  #column where balance remaining will be displayed
balAlreadyPaidCol = 'e'
creditDueCol = 'h'   #column where amount already paid will be stored
totalPriceDue = 'g' #total cost of the sport
parentsEmailsCol = 'd' #parents email list
dateBalDueCol = 'i' #date

def checkForExcelSheet():
    try:
        # attempt to load the Excel workbook named "PlayerWorkBook.xlsx"
        loadsPlayerWorkBook = openpyxl.load_workbook("PlayerWorkBook.xlsx")
        loadsPlayerWorkBook.active
        # if successful return the loaded workbook
        return loadsPlayerWorkBook
    except FileNotFoundError:
        # if the file is not found handle the exception
        print("The players WorkBook file was not found")
        print("Make sure the file is named correctly and/or it's in the same directory")
        # Return None to show failure in loading the workbook
        return None

# Call the function to check for the Excel workbook give off the workbook to playerWorkBook
playerWorkBook = checkForExcelSheet()


def checkPlayersWithBalance(playerWorkBook): 
    try:
        #checks if the function return a None to handle that error
        if playerWorkBook is not None:
            loadedWorkBook = playerWorkBook.active
            #gets the min and max of each row
            maxRow = loadedWorkBook.max_row
            minRow = loadedWorkBook.min_row

            #instantiated list for email function
            parentListEmail = []
            parentWithBalDue = []
            playerListName = []
            remainingTotalBal = []
            dateBalDue = []
            
            # For loop iterates between both column getting the values then getting the data needed
            for row in range(minRow + 1, maxRow + 1):
                #gets and credit and type cast it into an int and handles an incoming None's
                creditDue = loadedWorkBook[creditDueCol + str(row)]
                creditUse = int(creditDue.value) if creditDue.value is not None else 0
                #gets what the parent has already paid and type cast it into an int and handles an incoming None's
                pricePaid = loadedWorkBook[balAlreadyPaidCol + str(row)]
                pricePaidVal = int(pricePaid.value) if pricePaid.value is not None else 0
                #gets total cost and type cast it into an Int and handles an incoming None's
                totalRemaining = loadedWorkBook[totalPriceDue + str(row)]
                totalRemainingBal = int(totalRemaining.value) if totalRemaining.value is not None else 0
                # Payment logic to see who has and hasn't paid
                balDue = pricePaidVal + creditUse
                finalBal = totalRemainingBal == balDue
                
                # Says what parent exactly has finished their payment and who hasn't payed logic
                if not finalBal >= True:
                    # Append parent name
                    parentWithBal = loadedWorkBook[parentNameCol + str(row)]  
                    parentWithBalDue.append(parentWithBal.value)  
                    # Append parent email
                    parentExcelEmail = loadedWorkBook[parentsEmailsCol + str(row)]  
                    parentListEmail.append(parentExcelEmail.value) 
                    #appends parent name and child name
                    playerExcelName = loadedWorkBook[playerNameCol + str(row)]
                    playerListName.append(playerExcelName.value)
                    # Add remaining balance of
                    remainingTotal = loadedWorkBook[balRemainingCol + str(row)]
                    remainingTotalBal.append(remainingTotal.value) 
                    #adds date
                    dateForBal = loadedWorkBook[dateBalDueCol + str(row)]
                    dateBalDue.append(dateForBal.value)
            # Return both lists after the loop
            return parentWithBalDue, parentListEmail, playerListName, remainingTotalBal, dateBalDue
        # Return the needed lists after the loop
        else:
            print("Data Transfer Error")
            print("Error has occurred with worksheet. Make sure the file was named correctly")
            #looks out for common errors that could occur when updating the excel sheet
    except KeyError: 
        print("Doesn't exist in scope - wrong column letter most likely issue")
    except TypeError:
        print("Incorrect data type entered - a number in a string field or vice versa ")

 # takes all of needed data for the next function and takes it in lists to cycle through 
parentWithBalDue, parentListEmail, playerListName, remainingTotalBal, dateBalDue = checkPlayersWithBalance(playerWorkBook)

def parentsUncompletedPayment(parentWithBalDue, parentListEmail, playerListName,  remainingTotalBal, dateBalDue):
    try:
        # Optional file naming to make it easier to find and input validation on file
        emailParentFile = pyip.inputStr(prompt="Enter file name (type 'default' for default file)\nNo need to type .txt: ")
        #handling the file extension so the user doesn't need too
        if emailParentFile.lower() == 'default':
            emailParentFile = "emailParent.txt"
        elif not emailParentFile.endswith(".txt"):
            emailParentFile += ".txt"
        #opens whatever the user names the file
        emailParents = open(emailParentFile, 'w') 
        #makes sure the lengths are all the same making sure there an issue with None's
        if (len(parentWithBalDue) == len(parentListEmail) == len(playerListName) == len(remainingTotalBal) == len(dateBalDue)):    
            for i in range(len(parentWithBalDue)): 
                #context for the emails sends all the info the parents need for the soccer team to easily send an email
                emailParents.write(str(parentWithBalDue[i]) + ' - '  + str(parentListEmail[i]) + "\n")
                emailParents.write("Hello is an automated email from Peniche C.C Toronto informing you that your child\n"
                                    + str(playerListName[i]) + " still owes a balance of $" + str(remainingTotalBal[i]) + " for this season\n")
                #tell the parent when the payment is due date and time wise
                emailParents.write("Payment is due on " +  str(dateBalDue[i]))
                emailParents.write("\n-------------------------------------------------------------------------------------\n")
        print("Opening....")
        #give the program a little time to actually work so it doesn't lag when opening the txt file
        time.sleep(5)
        # Opens the certain notepad file making it easier on the user 
        subprocess.Popen(["notepad.exe",  emailParentFile])
    #checks and handles the errors to the file while also giving the user extra instructions to fix
    except FileNotFoundError:
        print("File not found -  please ensure you have Notepad installed.")
    except TypeError:
        print("Not a valid input -  please enter a filename without any special characters")
    

parentsUncompletedPayment(parentWithBalDue, parentListEmail, playerListName, remainingTotalBal, dateBalDue)
