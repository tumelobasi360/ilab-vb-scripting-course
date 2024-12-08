'Create a program that will receive the fullname from the user, 
'split the fullname into firstname and lastname. Then display 
'firstname and lastname separately, the first letters of the 
'firstname and lastname should be in upper case

'input
'fullname: mike ross

'output
'firstname: Mike
'lastname: Ross


'Declare
DIM fullName, strFirstName, strLastName
DIM intStrPosition
'Assign
fullName = Trim(inputBox("Please enter full name"))
intStrPosition = inStr(1, fullName, " ")


strFirstName = Trim(Left(fullName, intStrPosition - 1))
strLastName = Trim(Right(fullName, Len(fullName) - intStrPosition))

firstLetterOfFirstName = Left(strFirstName, 1)
firstLetterOfLastName = Left(strLastName, 1)

lenFirstName = Len(strFirstName)
lenLastName = Len(strLastName)

newFirstName = Ucase(firstLetterOfFirstName) + LCase(Right(strFirstName, lenFirstName -1))

newLastName = UCase(firstLetterOfLastName) + Lcase(Right(strLastName, lenLastName-1))
'Assign

'Use
MsgBox("Firstname: " & newFirstName & vbNewLine _
                     & "Lastname: " & newLastName)

