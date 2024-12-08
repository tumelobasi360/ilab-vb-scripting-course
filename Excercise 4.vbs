'display the following menu to the user:

 'Please enter any of these operators:
 '+ Addition
 '- Subtraction
 '* Multiply
 '/ Divide
 
 'the program should allow the user to use one operator from the menu. 
 'Display the menu as long as the user enters an incorrect operator. The program should
 'ask for 2 integers from the user and perform the neccessary computation based on
 'the selected operator, and then display the results.
 
 'The program should as the user if they want to continue using it, if they user says
 'yes then the program should display the menu again, otherwise the program should 
 'exit

DIM operator, intNumber1, intNumber2, repsone

DO
    intNumber1 = CInt(Trim(inputBox("Please enter the first number")))
    intNumber2 = CInt(Trim(inputBox("Please enter the second number")))


    operator = Trim(inputBox("Please enter any of these operators: " & vbNewLine _ 
            & "+ Addition" & vbNewLine _
            & " - Subrtaction" & vbNewLine _
            & " * Multiply" & vbNewLine _
            & "/ Devide"))

    DO WHILE operator <> "+" AND operator <> "-" AND operator <> "*" AND operator <> "/"
        operator = Trim(inputBox("Please enter any of these operators: " & vbNewLine _ 
            & "+ Addition" & vbNewLine _
            & " - Subrtaction" & vbNewLine _
            & " * Multiply" & vbNewLine _
            & "/ Devide"))
    LOOP

    IF operator = "+" THEN
            MsgBox("Results of: " & intNumber1 & " " & operator & " " & intNumber2 & " is: " & intNumber1 + intNumber2)
        ELSEIF operator = "-" THEN
            MsgBox("Results of: " & intNumber1 & " " & operator & " " & intNumber2 & " is: " & intNumber1 - intNumber2)
        ELSEIF operator = "*" THEN
            MsgBox("Results of: " & intNumber1 & " " & operator & " " & intNumber2 & " is: " & intNumber1 * intNumber2)
        ELSE
            DO WHILE intNumber2 = 0
            intNumber2 = CInt(Trim(inputBox("cannot devide by 0. please enter a different number")))
            LOOP
            MsgBox("Results of: " & intNumber1 & " " & operator & " " & intNumber2 & " is: " & intNumber1 / intNumber2)
    END IF
    repsone = LCase(inputBox("Do you want to continue? (Yes/No)"))

LOOP UNTIL repsone <> "yes"