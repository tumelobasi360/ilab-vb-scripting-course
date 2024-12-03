DIM intNumber, sum

sum = 0
DO 
    intNumber = CInt(inputBox("Please enter a number : "))
    IF intNumber <> 0 THEN
        sum = sum +intNumber
    END IF
LOOP WHILE intNumber <> 0

MsgBox("sum: " & sum)