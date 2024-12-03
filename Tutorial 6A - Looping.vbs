'********************************************
'       Tutorial 6 - Looping
'********************************************
REM This tutorial teaches you about different iteration structures

'display Tuesday 3times using a FOR loop

FOR x = 1 to 3
    MsgBox(x & " - Tuesday")
NEXT

'dispaly even numbers between 1 and 10 using FOR loop
FOR x = 2 to 10 STEP 2
    MsgBox(x)
NEXT

'dispaly Monday 3 times using a while loop
DIM intCount

intCount = 5

DO WHILE intCount <= 7
    MsgBox(intCount & " - Monday")
    intCount = intCount + 1
LOOP

'display Thursday 4 times using Do until
DIM intCountUntil

intCountUntil = 1

DO UNTIL intCountUntil = 4
    MsgBox("Thurday")
    intCountUntil = intCountUntil + 1
LOOP

'display Friday 3 times using a do loop
doCount = 1

DO 
    MsgBox("Friday")
    doCount  = doCount +1
LOOP UNTIL doCount = 4