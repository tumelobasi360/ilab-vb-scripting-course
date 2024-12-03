
REM using nested conditional statements and logical operators

DIM intNumber
DIM strResult 

intNumber = 38

'25 OR 40
IF intNumber > 20 THEN
    'TRUE
    IF intNumber = 25 THEN
        strResult = "The number is: " & intNumber
    ELSEIF intNumber = 40 THEN
        strResult = "The number is: " & intNumber
    ELSE
        strResult = "Not the number we are looking for [" & intNumber & "]"
    END IF
ELSE
    'FALSE
    strResult = "Number [" & intNumber & "] is greater than 20"
END IF

MsgBox(strResult)