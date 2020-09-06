# Text-to-Excel-Converter
Text to Excel Converter Using Openpyxl

This is a practice script to convert text files to excel files. 

The general process is as follows:
1. Import the necessary modules, the main one being openpyxl for this exercise.
2. Open a new excel workbook and define a sheet for data from the text files to go to.
3. Set the current working directory to one containing your text files.
4. Define the directory where your text files are. 
5. Define a list to later insert the line of text to be contained in it, which will be used to copy to the excel sheet.
6. Using os.walk, walk through the directories needed in order to find all the text files, have readlines() append the text to list from step 5, and iterate until it finds all the textfiles in the directories being walked through.
7. Use the proper logic to loop through the list of appended text and insert it into the proper sections of the excel sheet.
8. Save the excel workbook.
