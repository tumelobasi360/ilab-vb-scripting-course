'Declare
DIM testpercentage, assignment1percentage, assignment2percentage, classAttendancepercentage

'Assign
testpercentage = 49
assignment1percentage = 50
assignment2percentage = 61
classAttendancepercentage = 60

IF testpercentage >= 50 THEN
    MsgBox("The student can submit the assignmets")
    IF assignment1percentage >= 55 THEN
        MsgBox("The student passed the first assignment")
        IF assignment2percentage >= 60 THEN 
            MsgBox("The student passed the second assignment")
        ELSEIF classAttendancepercentage >= 75 THEN
            MsgBox("The student has at least 75% class attendance")
        ELSE 
        MsgBox("The student fails the second assignment and has less than 70% class attendance")
        END IF
    ELSE
        MsgBox("The student fails the first assignment")
    END IF

ELSE
    MsgBox("The student cannot submit the assignments since the test score is below 60%")
END IF
