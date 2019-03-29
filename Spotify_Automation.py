import openpyxl as pyxl
import pandas as pd
import numpy as np
import datetime as dt

"""
Things to do: 
Figure out how to incorporate FX CCY Checks to make ingestion easier. Also - ask if Intl vs. Domestic really matters? 
Modify Code to focus on copying 3 Situations - Clean, USD and Intl under a single Function. 
Modify Code to handle Currency Checks when copying the data - Needs Yes - USD, Yes - Intl, and No
Modify Code to handle passing through Sheet Number to simplify code to handle copying data more effectively
"""


# Sets Workbook Name and Loads workbook
wb_Spotify = pyxl.Workbook()
wb_Spotify = pyxl.load_workbook('Spotify_Test_XLS.xlsx') #Cannot do CSV
wb_Spotify_Sheets = []  # Allows for collection of Sheet Names
ws_Active = wb_Spotify.active # Defines Active Sheet Function
# Sets Header Row Data
HeaderRow = ['Source', 'Geocuntry','Product','URI','UPC','EAN','ISRC','Type','Title','Artist','Composer','Album Name',
             'Streams','Label','Payable invoice','Invoice currency','Payable EUR','Revenue']

# Creates all sheets needed for processing the original file
ws_Active.title = "Spotify Raw"
ws_Clean = wb_Spotify.create_sheet('Spotify Clean')
ws_USD = wb_Spotify.create_sheet('Spotify USD')
ws_Intl = wb_Spotify.create_sheet('Spotify Intl')

# Sets Sheet Name List for Active Sheet Management
for i in range (len(wb_Spotify.sheetnames)):
    wb_Spotify_Sheets.append(wb_Spotify.sheetnames[i])

# Sets Starting and Ending Rows and Columns
Starting_Row = 1
Ending_Row = ws_Active.max_row
Starting_Column = 1
Ending_Column = len(HeaderRow)
print(Ending_Column)

# Fx used to copy a range of data
def Copy_Range(Start_Col, Start_Row, End_Col, End_Row, sheet):
    # Holds a full Array of Columns and Rows to be copied
    rangeSelected = []


    # Loops through selected Rows
    for i in range(Start_Row, End_Row + 1, 1):

        # Holds an Array of all the data for a given row
        rowSelected = []

        for j in range(Start_Col, End_Col + 1, 1):

            # Handles appending the array to house row data
            rowSelected.append(sheet.cell(row=i, column=j).value)


        # Appends the Range Array to store individual rows as part of a bigger array
        rangeSelected.append(rowSelected)

    return rangeSelected

# Takes the data that was copied and pastes it into a specific location
def Paste_Range(Start_Col, Start_Row, End_Col, End_Row, Sheet_Transfered, Copied_Data):

    #
    Count_Row = 0

    for i in range(Start_Row, End_Row + 1, 1):
        Count_Col = 0

        for j in range(Start_Col, End_Col + 1, 1):

            Sheet_Transfered.cell(row=i, column=j).value = Copied_Data[Count_Row][Count_Col]
            Count_Col += 1

        Count_Row += 1
    return


print(ws_Active.max_row)

# Fx for handling the copy over
def Copy_Data(Start_Col, Start_Row, End_Col, End_Row, Input_Sheet, Output_Sheet):
    Selected_Range = Copy_Range(Start_Col, Start_Row, End_Col, End_Row, Input_Sheet)
    Pasted_Range = Paste_Range(Start_Col, Start_Row, End_Col, End_Row, Output_Sheet, Selected_Range)
    return

#Handles Amending the Clean Sheet to read the way it should
def Amend_Clean_Data(Sheet_Name, HeaderRow):
    Sheet_Name.delete_rows(1,2)
    Sheet_Name.insert_cols(1,1)
    Sheet_Name.insert_cols(8,1)

    print(dt.time())
    for i in range (0, len(HeaderRow)):
        print(HeaderRow[i])

    Count_Row = 1
    Count_Col = 0


    for col in Sheet_Name.iter_cols(min_row=2, max_col=1, max_row=Sheet_Name.max_row):
        for cell in col:
            cell.value = "Spotify"

    for col in Sheet_Name.iter_cols(min_row=2, min_col=8, max_row=Sheet_Name.max_row, max_col=8):
        for cell in col:
                cell.value = "sound_recording"


    while Count_Col < len(HeaderRow):
        Sheet_Name.cell(row = 1, column = Count_Col+1).value = HeaderRow[Count_Col]
        Count_Col +=1

    return
'''
    while Count_Row < (Sheet_Name.max_row+1):

        if Count_Row == 1:

            while Count_Col < (len(HeaderRow)):
                Sheet_Name.cell(row = 1, column = (Count_Col+1)).value = HeaderRow[Count_Col]
                Count_Col += 1
        else:
            Sheet_Name.cell(row=Count_Row, column=1).value = "Source"
            Sheet_Name.cell(row=Count_Row, column=8).value = "sound_recording"

        Count_Row +=1
    
    return
'''

def FX_Cleanup(CCY, sheet):

    if CCY == "USD":

        for i in range(2, sheet.max_row +1, 1):
            if sheet.cell(row = i+1, column = 16).value == "USD":
                continue
            else:
                sheet.delete_rows(i,1)
    else:
        for i in range(2, sheet.max_row +1, 1):
            if sheet.cell(row = i+1, column = 16).value != "USD":
                continue
            else:
                sheet.delete_rows(i,1)

    return




print(dt.time())
Copy_Data(Starting_Column, Starting_Row, Ending_Column, Ending_Row, ws_Active, ws_Clean)
print(dt.time())
Amend_Clean_Data(ws_Clean, HeaderRow)
print(dt.time())
Copy_Data(Starting_Column, Starting_Row, Ending_Column, Ending_Row, ws_Clean, ws_USD)
print(dt.time())
Copy_Data(Starting_Column, Starting_Row, Ending_Column, Ending_Row, ws_Clean, ws_Intl)
print(dt.time())
#FX_Cleanup("USD", ws_USD)
print(dt.time())
#FX_Cleanup("EUR", ws_Intl)
print(dt.time())
wb_Spotify.save('Spotify_Output_File.xlsx')
print(ws_Clean['A1'].value)
print(ws_Clean.cell(row=1,column=16).value)