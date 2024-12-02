'********************************************
'       Tutorial 4 - Working with arrays
'********************************************

REM This tutorial teaches you about Arrays
REM - variable that can store multiple values of the same data types

'Declare
DIM arrIntNumbers(5) 'emty array to store 5 items / numbers
DIM arrStrNames
'declare an array with intial values
arrStrNames = Array("Kelvin", "Thabo", "Kabelo", "Caroline", "David", "Jessica")

'Assign
arrIntNumbers(0) = 45
arrIntNumbers(1) = 38
arrIntNumbers(2) = 74
arrIntNumbers(3) = 25
arrIntNumbers(4) = 55

'Use
MsgBox(arrIntNumbers(2))
MsgBox(arrStrNames(5))