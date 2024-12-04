'********************************************
'       Tutorial 7 - Subrotines
'********************************************
REM This script teaches you about subrotines
'create a subrotine that will display a welcome message
Sub showMessage(strName)
    MsgBox("Hi, " & strName & " Welcome to VB Scripting Course")
end sub

sub showFullName(strName, strSurname)
    DIM strFullName

    strFullName = strName & " " & strSurname
    MsgBox("Fullname: " & strFullName)
end sub
'=============================================
DIM strName, strSurname

strName = inputBox("Please enter your name: ")
strSurname = inputBox("Please enter your surname: ")
Call showMessage(strName)
call showFullName(strName, strSurname)