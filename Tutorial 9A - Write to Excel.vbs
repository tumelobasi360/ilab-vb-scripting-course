'********************************************
'       Tutorial 9 - Write to Excel
'********************************************
REM This script will write to an Excel Workbook

Option Explicit

' Declare
DIM objExcel ' Object used to open Excel Workbook
DIM strFilePath ' Location of the Workbook
DIM objFileSystemObject

' Assign
SET objExcel = CreateObject("Excel.Application")
SET objFileSystemObject = CreateObject("Scripting.FileSystemObject")
strFilePath = "C:\Users\Tumelo Maphalla\Documents\VBScript\Interns\Files\excel_write.xls"

' Check if file exists
IF objFileSystemObject.FileExists(strFilePath) THEN
    objExcel.Workbooks.Open(strFilePath) ' Open the Workbook
    objExcel.ActiveWorkbook.Save
ELSE
    ' Create a new Workbook and write data
    objExcel.Workbooks.Add
    objExcel.Worksheets(1).Cells(1,1) = "City"
    objExcel.Worksheets(1).Cells(1,2) = "Population"
    objExcel.Worksheets(1).Cells(2,1) = "Pretoria"
    objExcel.Worksheets(1).Cells(2,2) = 900000
    objExcel.Worksheets(1).Cells(3,1) = "Midrand"
    objExcel.Worksheets(1).Cells(3,2) = 1000000
    objExcel.Worksheets(1).Cells(4,1) = "Joburg"
    objExcel.Worksheets(1).Cells(4,2) = 1200000
    objExcel.Worksheets(1).Cells(5,1) = "Vaal"
    objExcel.Worksheets(1).Cells(5,2) = 36036500
    objExcel.ActiveWorkbook.SaveAs(strFilePath)
END IF

' Close the Excel application
objExcel.Quit

' Cleanup objects
SET objExcel = Nothing
SET objFileSystemObject = Nothing
