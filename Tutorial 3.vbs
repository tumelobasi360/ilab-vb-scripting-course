REM This tutorial teaches you about string functions
'********************************************
'       Tutorial 3 - String fucntions
'********************************************

'Declare
    DIM strFirstName, strLastName
    DIM strFullname
    DIM intStrPosition
    DIM strSearchStr
    DIM strCourseName, strCourseInput
    DIM strStringCompare
'Assign
strFirstName = Trim(inputBox("Please enter first name"))
strLastName = Trim(inputBox("Please enter last name"))

strFullname = UCase(strFirstName & " " & strLastName)
strCourseName = "VB Scripting"
'Use
MsgBox("Firstname: " & strFirstName & vbNewLine _
                     & "Lastname: " & strLastName & vbNewLine _
                     & "============================" & vbNewLine _
                     & "Fullname: " & strFullname & vbNewLine _
                     & strFirstName & " has " & Len(strFirstName) & " characters " & vbNewLine _
                     & "The first 3 characters of " & strFirstName & " are: " & Left(strFirstName, 3))
strSearchStr = inputBox("Type the letter to search from the name (" & strFirstName & ")" )
intStrPosition = inStr(1, strFirstName, strSearchStr)

MsgBox(strSearchStr & " is found at position: " & intStrPosition)

strCourseInput = inputBox("Guess the name of the course! ")
strStringCompare = strComp(strCourseName, strCourseInput)
MsgBox("Comaprison between " & strCourseInput & " and " & strCourseName & " is: " & strStringCompare)