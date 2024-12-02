'********************************************
'       Tutorial 1
'********************************************
REM This tutorial teaches you about diffrent data types

Option Explicit

'Declare
DIM strName
DIM intAge
DIM dblHeight
DIM isMarried
'Assign
'Get input from the user
strName = inputBox("Please enter your name")
intAge = CInt(inputBox("Please enter your age"))
dblHeight = CDbl(inputBox("Please enter your height"))
isMarried = CBool(CInt(inputBox("Is Married? " & vbNewLine _
                                         & "1 - Yes" & vbNewLine _
                                         & "0 - No")))

'Use
MsgBox("Name: " & strName & " - Age: " & intAge & " - Height: " & dblHeight)
MsgBox("Is Married: " & isMarried)
MsgBox("Name: " & strName & vbNewLine _
                & "Age: " & intAge & vbNewLine _
                & "Height: " & dblHeight & vbNewLine _
                & "Is Married? " & isMarried)

