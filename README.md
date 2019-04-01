# Spotify_Automation

The purpose of this code is to automate the process of taking the raw Spotify report and edit it to have the proper headers and columns for ingestion as dividing the report into two smaller portions - one for Spotify Direct (USD) and one for Spotify International (Non-USD). 

To run this report, the code needs to be directed to the appropriate file to ingest. The file must be in .xlsx format due to the use of OpenPyxle for Python. After entering the file directory and file name, the code take the file, create three additional worksheets (one that is cleaned up from the original raw, one that is for USD only and one for International), add the appropriate columns and headers, copies the data across three sheets, then handles any cleanup of unnecessary information and saves the data to an output workbook. 

Required Libraries:
  OpenPyxl - Handles all functions related to editing and modifying an XLSX
  CSV - To handle output of data into CSV format for ingestion. 

To Do List: 
  Timer Functions to track how long each task takes
      Purpose - Figure out which areas need further optimization and which ones don't. 
  Implement CSV Output
  Merge FX_Cleanup with Copy/Paste functions
  Delete any columns not needed for ingestion (ex. FX_Toggle Column) upon completion of code
  Unit_Testing Code
    File Extention Matching - must be .XLSX
    Header Matching on the original file to confirm that it's an actual report
    Currency Checking on the outputs to ensure that no incorrect currencies are copied to the wrong report.
  Testing for Size - Currently it runs smoothly with 100 Row and 5000 Row test files. Need to examine 10K and 50K Table Sizes.     
  

Last Edit: 4/1/19
MJF
