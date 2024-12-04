
CLASS Employee
    public strFirstname
    public strLastname
    public strEmployeeNo
    public intHorsWorked
    
    public Function welcomeMessage()
        welcomeMessage = "Hi " & strFirstname & " " & strLastname & " Welcome to Employee management system"
    END Function

    private Function calcSalary()
        DIM dblRate
        dblRate = 250
        calcSalary = intHorsWorked * dblRate
    END Function

    private Function getEmployeeNo()
        getEmployeeNo = "EMP-" & strEmployeeNo
    END Function
    public Function getDetails()
        getDetails = "Fullname: " & strFirstname & " " & strLastname & vbNewLine _
                     & "Employee No: " & getEmployeeNo() & vbNewLine _
                     & "Hours Worked: " & intHorsWorked & vbNewLine _
                     & "Salary: " & calcSalary()
    END Function
END CLASS
'=============================================
'Declare
    SET objEmployee = new Employee
    Const strTitle = "Employee Management System"
'Assign
    objEmployee.strEmployeeNo = inputBox("Please enter Employee No", strTitle)
    objEmployee.strFirstname = inputBox("Please enter firstname", strTitle)
    objEmployee.strLastname = inputBox("Please enter lasttname", strTitle)
    objEmployee.intHorsWorked = CInt(inputBox("Please enter hours worked", strTitle))

'Use
    MsgBox objEmployee.welcomeMessage(), vbOkOnly, strTitle
    MsgBox objEmployee.getDetails(), vbOkOnly, strTitle