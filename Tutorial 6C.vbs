
'display the following menu to the user
'
DIM operator


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
        MsgBox("You selected a " & operator & " Addition operator.")
    ELSEIF operator = "-" THEN
        MsgBox("You selected a " & operator & " Subtraction operator.")
    ELSEIF operator = "*" THEN
        MsgBox("You selected a " & operator & " Multiplication operator.")
    ELSE
        MsgBox("You selected a " & operator & " Division operator.")
    END IF