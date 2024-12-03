'********************************************
'       Tutorial 5 - Working with arrays
'********************************************
REM This tutorial teaches you the different selection structures

'Declare
DIM intAge
DIM strResult

'Assign
intAge = 25

'Unary IF
IF intAge >= 18 THEN
    strResult = "Adult"
END IF

'Binary IF
IF intAge >= 18 THEN
    strResult = "Adult"
ELSE strResult = "Not an Adult"
END IF

MsgBox("Age: " & intAge & vbNewLine _
               & "Results: " & strResult)

IF intAge >= 18 THEN
    strResult = "Adult"
ELSEIF intAge >= 13 THEN
    strResult = "Teenager"
ELSEIF intAge >= 0 THEN
    strResult = "Kid"
ELSE strResult = "Age cannot be less than 0"
END IF

'Use
MsgBox("Age: " & intAge & vbNewLine _
               & "Results: " & strResult)