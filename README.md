# BlueBeam-Script-Engine-Markup-Update
This is a VBA script built for Excel and BlueBeam. The script will update custom column data for markups in BlueBeam.
The intention of this program is to dramatically speed up the markup process in BlueBeam using the BlueBeam Script Engine. Markups are often expoted into CSV files and uploaded to Excel files in order to manipulate or make calculations based off of the markup data. However, this data normally is either updated manually in BlueBeam after the fact or not updated at all. Utilizing this VBA code along with Excel's formulas, one can quickly update markups in mass.

This process works by writing all of the new markup data to a BCI file, which is a BlueBeam Script file. The program will then send a shell command to the BlueBeam Script Engine to execute the contents of that BCI file, updating all of the target markups.

To run this program, the user must export a markup list CSV file with custom columns showing in the Markup List. Then copy the contents of the CVS file into a sheet named "IMPORT". Update the custom column data with the new data you want to upload to BlueBeam. Then, on a sheet names "Sheet1", enter in Cell A1 the full folder path of the script engine (C:\Users\etc\Revu\ScriptEngine.exe). In Cell A2 enter the full folder path of the BCI file (C:\Users\etc\myScript.bci). In Cell A3 enter the full folder path of the target pdf (C:\Users\etc\target.pdf). Now run the VBA Subroutine UpdateMarkups to update the pdf with the new custom column data. 

Make sure to Close the PDF before running this code! Otherwise the process will not work. 
