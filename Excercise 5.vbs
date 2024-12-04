

Sub showWelcomeMessage(strName)
    MsgBox("Hi, " & strName & " Welcome to VB Scripting Course")
end sub

strName = inputBox("Please enter your name: ")

call showWelcomeMessage(strName)

Function getSum(intNumber1, intNumber2)
    getSum = intNumber1 + intNumber2
end Function

Function getDifference(intNumber1, intNumber2)
    getDifference = intNumber1 - intNumber2
end Function

Function getProduct(intNumber1, intNumber2)
    getProduct = intNumber1 * intNumber2
end Function

Function getquotient(intNumber1, intNumber2)
    
    getquotient = intNumber1 / intNumber2
end Function

DIM sum, intNumber1, intNumber2



intNumber1 = CInt(Trim(inputBox("Enter the first number: ")))
intNumber2 = CInt(Trim(inputBox("Enter the second number: ")))


sum = getSum(intNumber1, intNumber2)
difference = getDifference(intNumber1, intNumber2)
product = getProduct(intNumber1, intNumber2)
quotient = getquotient(intNumber1, intNumber2)

MsgBox("The results of the sum: " & sum)
MsgBox("The results of the subtraction: " & difference)
MsgBox("The results of the multiplication: " & product)
MsgBox("The results of the division: " & quotient)

