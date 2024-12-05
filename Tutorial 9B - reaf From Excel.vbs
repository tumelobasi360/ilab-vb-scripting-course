'********************************************
'       Tutorial 9B - Read From Excel
'********************************************
REM This tutorial teaches you how to read from an Excel file

' Declare
DIM objExcel
DIM objWorkbook
DIM objWorksheet
DIM strFilePath
DIM intRow
DIM strOutput

' Assign
strFilePath = "C:\Users\Tumelo Maphalla\Documents\VBScript\Interns\Files\Book1.xlsx"
SET objExcel = CreateObject("Excel.Application")
SET objWorkbook = objExcel.Workbooks.Open(strFilePath)
SET objWorksheet = objWorkbook.Worksheets(1) ' Access the first worksheet

' Initialize variables
intRow = 2 ' Start reading from row 2
strOutput = objWorksheet.Cells(1, 1).Value & "          " & objWorksheet.Cells(1, 2).Value & vbNewLine _
            & "======================================" & vbNewLine

' Read data until an empty cell is found in column 1
DO UNTIL objWorksheet.Cells(intRow, 1).Value = ""
    strOutput = strOutput & objWorksheet.Cells(intRow, 1).Value & "             " & objWorksheet.Cells(intRow, 2).Value & vbNewLine
    intRow = intRow + 1
LOOP

' Display the output
MsgBox strOutput

' Cleanup
objWorkbook.Close False ' Close without saving changes
objExcel.Quit
SET objWorksheet = Nothing
SET objWorkbook = Nothing
SET objExcel = Nothing
