'********************************************
'       Tutorial 7B - Functions
'********************************************
REM This script teaches you about Functions

Function getFullname()
    DIM strFirstName, strLastName, strFullName
    'Assign
    strFirstName = Trim(inputBox("Please enter FirstName: "))
    strLastName = Trim(inputBox("Please enter LastName: "))
    strFullName = strFirstName & " " & strLastName

    getFullname = strFullName
end Function

'============================================
'Declare
DIM strFullName
'Assign
strFullName = getFullname()
'Use
MsgBox("Fullname: " & strFullName)