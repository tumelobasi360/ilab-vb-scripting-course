

'Declare
DIM strFirstName, strLastName
DIM lenFirstName, lenLastName

'Assign
strFirstName = Trim(inputBox("Please enter first name"))
strLastName = Trim(inputBox("Please enter last name"))

firstLetterOfFirstName = Left(strFirstName, 1)
firstLetterOfLastName = Left(strLastName, 1)

lenFirstName = Len(strFirstName)
lenLastName = Len(strLastName)

newFirstName = Ucase(firstLetterOfFirstName) + LCase(Right(strFirstName, lenFirstName -1))

newLastName = UCase(firstLetterOfLastName) + Lcase(Right(strLastName, lenLastName-1))

'Use

MsgBox(newFirstName & " " & newLastName)