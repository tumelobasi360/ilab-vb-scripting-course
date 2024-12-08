'create a program that will perform different calculations using 
'functions. The program should have the following functions:
'1 - a function to calculate the sum 
'2 - a function to calculate the product 
'3 - a function to calculate the difference 
'4 - a function to calculate the quotient
'All these functions should take 2 numbers as parameters 
'The program should have the following subroutine:
'1 - a sub to display a welcome message
'Lastly, the program should display the sum, difference, product
'and quotient of the 2 numbers

Sub showWelcomeMessage(strName)
    MsgBox("Hi, " & strName & " Welcome to our Calculator Universe")
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

MsgBox("The sum of " & intNumber1 & " and " & intNumber1 & " is: " & sum & vbNewLine _
                & "The product of " & intNumber1 & " and " & intNumber2 & " is: " & product & vbNewLine _
                & "The difference of " & intNumber1 & " and " & intNumber2 & " is: " & difference & vbNewLine _
                & "The quotient of " & intNumber1 & " and " & intNumber2 & " is: " & quotient)
