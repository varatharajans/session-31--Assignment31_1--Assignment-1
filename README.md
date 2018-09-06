# session-31--Assignment31_1--Assignment-1
DATA ANALYTICS WITH R, EXCEL AND TABLEAU SESSION 31 ASSIGNMENT 31_1
                                                                                                                                                  SESSION 8: EXCEL ANALYTICS (CONTD.) 
                                                                                                                                                                   Assignment 1

4. Associated Data Files Use “Sales_Dataset.xlsx” file 
5. Problem Statement 
• Create and execute a macro (show a button), to sum of “Sales” where the “attractiveness” rating is more than 5. 
• Do it by recording a Macro. 
• Try it using writing a VBA code in Macro.
• Use http://www.excel-easy.com/vba.html for reference to write VBA script.
 • Protect the Workbook and save version and upload as a first part of assignment 
• Share the protected workbook, by giving edit access and implement track change. Upload this workbook as second part of your assignment.  
Note:- Both upload should contain macros (two types as mentioned above).


Answers
Details are submitted as excel output downloadable files
A macro is created with a command button, on clicking which, sum of sales for "attractiveness>5 " is displayed as a message.

The workbook is protected with password "password@123456" , saved and uploaded.




The protected workbook is shared by giving edit access and implementing  track change.

The macro code is copied below for reference: ( common for both files)

Sub Macro1()
Sum = 0
For Row = 2 To 201
If Cells(Row, 4) > 5 Then
Sum = Sum + Cells(Row, 2)
End If
Next Row
MsgBox ("Sum of sales where attractiveness greater than 5 => " & Sum)
End Sub


