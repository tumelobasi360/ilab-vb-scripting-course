Option Explicit

' Declare
DIM objExcel ' Object used to open Excel Workbook
DIM strFilePath ' Location of the Workbook
DIM objFileSystemObject

' Assign
strFilePath = "C:\Users\Tumelo Maphalla\Documents\VBScript\Interns\Files\Excecise6.xls"
SET objExcel = CreateObject("Excel.Application")
SET objFileSystemObject = CreateObject("Scripting.FileSystemObject")
'SET obj = objExcel.Workbooks.open(strFilePath)


' Check if file exists
IF objFileSystemObject.FileExists(strFilePath) THEN
    objExcel.Workbooks.Open(strFilePath) ' Open the Workbook
    objExcel.ActiveWorkbook.Save
ELSE
    ' Create a new Workbook and write data
    objExcel.Workbooks.Add
    objExcel.Worksheets(1).Cells(1,1) = "Student name"
    objExcel.Worksheets(1).Cells(1,2) = "test mark"
    objExcel.Worksheets(1).Cells(1,3) = "Assignment mark"
    objExcel.Worksheets(1).Cells(1,4) = "Final mark"
    objExcel.Worksheets(1).Cells(2,1) = "Tumelo Maphalla"
    objExcel.Worksheets(1).Cells(2,2) = 50
    objExcel.Worksheets(1).Cells(2,3) = 52
    objExcel.Worksheets(1).Cells(2,4) = (objExcel.Worksheets(1).Cells(2,2).Value + objExcel.Worksheets(1).Cells(2,3).Value)/2

    objExcel.Worksheets(1).Cells(3,1) = "Tumelo West"
    objExcel.Worksheets(1).Cells(3,2) = 65
    objExcel.Worksheets(1).Cells(3,3) = 65
    objExcel.Worksheets(1).Cells(3,4) = (objExcel.Worksheets(1).Cells(3,2).Value + objExcel.Worksheets(1).Cells(3,3).Value)/2

    objExcel.Worksheets(1).Cells(4,1) = "Tumelo Basi"
    objExcel.Worksheets(1).Cells(4,2) = 70
    objExcel.Worksheets(1).Cells(4,3) = 72
    objExcel.Worksheets(1).Cells(4,4) = (objExcel.Worksheets(1).Cells(4,2).Value + objExcel.Worksheets(1).Cells(4,3).Value)/2
    objExcel.ActiveWorkbook.SaveAs(strFilePath)
END IF

' Close the Excel application
objExcel.Quit