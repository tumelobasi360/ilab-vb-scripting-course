'********************************************
'       Tutorial 8 - Classes and Objects
'********************************************
REM This tutorial teaches you about Classes and Objects

CLASS Calculator

    public Function welcomeMessage()
        welcomeMessage = "Welcome to our calculator app"
    END Function
    public Function calcSum(intNumber1, intNumber2)
        calcSum = (intNumber1 + intNumber2)
    END Function

    Function getDifference(intNumber1, intNumber2)
        getDifference = intNumber1 - intNumber2
    end Function

    Function getProduct(intNumber1, intNumber2)
        getProduct = intNumber1 * intNumber2
    end Function

    Function getquotient(intNumber1, intNumber2)
        DIM dblQuotient
        IF num2 = 0 THEN
            dblQuotient = "Undefined"
        ELSE
            dblQuotient = intNumber1 / intNumber2
        END IF
        getquotient = dblQuotient
    end Function

    public Function display()
        display = "Sum: " & sum & vbNewLine _
               & "Product: " & product & vbNewLine _
               & "Difference: " & difference & vbNewLine _
               & "Quotient: " & quotient
    END Function
END CLASS
'=============================================
'Declare
    DIM num1, num2, sum, product, difference, quotient
    Const strTitle = "Calclator App"
    'declare / create an object of calculator class
    SET objCalculator = NEW Calculator

'Assign
    'call a sub from the class
    MsgBox objCalculator.welcomeMessage(), vbOkOnly, strTitle

    num1 = CInt(inputBox("Enter number 1"))
    num2 = CInt(inputBox("Enter number 2"))

    sum = objCalculator.calcSum(num1, num2)
    difference = objCalculator.getDifference(num1, num2)
    product = objCalculator.getProduct(num1, num2)
    quotient = objCalculator.getquotient(num1, num2)


'Use


MsgBox objCalculator.display(), vbOkOnly, strTitle