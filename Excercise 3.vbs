'Scenario: The program should determine if the student can submit assignment 1 
'and assignment 2. This will be based on whether the student scores above 50% 
'for their test. The student needs 55% and above on  assignment 1, 
'if they do, ask them to submit assignment 2. The student needs 60% and above 
'for assignment 2. If the student fails assignment 2, the program should 
'check if they have at least 75% class attendance.

'Declare
DIM intTestMark, assignment1percentage, assignment2percentage, classAttendancepercentage

'Assign
intTestMark = CInt(InputBox("Please enter test mark"))
IF intTestMark >= 50 THEN
    assignment1percentage = CInt(InputBox("Please enter assignment 1 mark"))
    IF assignment1percentage >= 55 THEN
        assignment2percentage = CInt(InputBox("Please enter assignment 2 mark"))
        IF assignment2percentage >= 60 THEN 
            MsgBox("You passed all the assignments")
        ELSE
            classAttendancepercentage = CInt(InputBox("Please enter class attendance"))
            IF classAttendancepercentage >= 75 THEN
            MsgBox("You have have at least 75% class attendance")
            ELSE 
            MsgBox("You failed to achieve at least 75% attendence")
            END IF
        END IF
    ELSE
        MsgBox("You failed Assignment 1, cannot submit assignment 2")
    END IF
ELSE
    MsgBox("You cannot submit the assignments since the test score is below 60%")
END IF