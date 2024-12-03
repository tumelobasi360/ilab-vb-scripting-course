'Declare
DIM fullName, strFirstName, strLastName
DIM intStrPosition
'Assign
fullName = Trim(inputBox("Please enter full name"))

intStrPosition = inStr(1, fullName, " ")
strFirstName = Left(fullName, Len(fullName) - intStrPosition-1)
strLastName = Right(fullName, Len(fullName) - intStrPosition)

firstLetterOfFirstName = Left(strFirstName, 1)
firstLetterOfLastName = Left(strLastName, 1)

lenFirstName = Len(strFirstName)
lenLastName = Len(strLastName)

newFirstName = Ucase(firstLetterOfFirstName) + LCase(Right(strFirstName, lenFirstName -1))

newLastName = UCase(firstLetterOfLastName) + Lcase(Right(strLastName, lenLastName-1))
'Assign

'Use
MsgBox(newFirstName)
MsgBox(newLastName)
'MsgBox(strFirstName)
'MsgBox(strLastName)