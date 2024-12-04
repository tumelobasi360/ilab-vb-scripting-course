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
            MsgBox("Results: " & intNumber1 + intNumber2)
        ELSEIF operator = "-" THEN
            MsgBox("Results: " & intNumber1 - intNumber2)
        ELSEIF operator = "*" THEN
            MsgBox("Results: " & intNumber1 * intNumber2)
        ELSE
            DO WHILE intNumber2 = 0
            intNumber2 = CInt(Trim(inputBox("cannot devide by 0. please enter a different number")))
            LOOP
            MsgBox("Results: " & intNumber1 / intNumber2)
    END IF
    repsone = LCase(inputBox("Do you want to continue? (Yes/No)"))

LOOP UNTIL repsone <> "yes"