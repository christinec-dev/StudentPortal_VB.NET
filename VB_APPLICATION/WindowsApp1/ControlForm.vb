'Import the system for the MS Access Database
Imports System.Data.OleDb

Public Class ControlForm

    'Global DB variables to be used between Private Subs.
    Dim dt As DataTable
    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim findDT As String
    Dim insertDT As String
    Dim modifyDT As String
    Dim deleteDT As String
    Dim dbFilePath As String = "C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb"

    'Secret passcode admins know to ensure verification, and authorization of events so students cannot access or submit unauthorized data.
    Dim passcode As String = "@dmin_all0w"

    Private Sub ControlForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'On form Load
    End Sub

    'About Developer Tab Start ---------------------------------------------------------------------------------------------------------------------------
    Private Sub AboutTab_Click(sender As Object, e As EventArgs) Handles AboutTab.Click, AboutTab.Enter

        'Executing the about form to show when the About Developer tab is clicked and entered..
        AboutForm.Show()

    End Sub

    'About Developer Tab End -------------------------------------------------------------------------------------------------------------------------------

    'Student Section Start --------------------------------------------------------------------------------------------------------------------------------
    Private Sub StudentFindBtn_Click(sender As Object, e As EventArgs) Handles StudentFindBtn.Click

        'Query to find the entered Student ID from the textbox in the MS DB
        findDT = ("SELECT * from Students_Table WHERE [Student ID] = '" & (StudentIDBox.Text) & "'")

        'Establishing and opening a connection to the DB where the table is
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        cmd = New OleDbCommand(findDT, con)
        con.Open()

        'Executing the Query above to the function below.
        Dim myRead As OleDb.OleDbDataReader = cmd.ExecuteReader

        'Function for finding a student and all their details via entering a Student ID, and returning the values in the textboxes.
        With myRead
            If myRead.Read() Then
                StudentFNameBx.Text = myRead("First Name")
                StudentMNameBx.Text = myRead("Middle Name")
                StudentLNameBx.Text = myRead("Last Name")
                StudentDadNameBx.Text = myRead("Father's Name")
                StudentMomNameBx.Text = myRead("Mother's Name")
                StudentBloodGBx.Text = myRead("Blood Group")
                StudentBDayBx.Text = myRead("Date of Birth")
                StudentEmailBx.Text = myRead("Email ID")
                StudentNumBx.Text = myRead("Mobile Number")
                StudentPAddressBx.Text = myRead("Permanent Address")
                StudentLAddressBx.Text = myRead("Local Address")
                StudentGenderBx.Text = myRead("Gender")
                StudentCategBx.Text = myRead("Category/Race")
            Else
                'Error handling for the scenario of a student not found/non existent.
                MessageBox.Show("Invalid user or record not found.", "Invalid Entry")
            End If
        End With

        'Closing the connection to the DB where the table is.
        con.Close()

    End Sub

    Private Sub StudentClearBtn_Click(sender As Object, e As EventArgs) Handles StudentClearBtn.Click

        'Function to clear all student details and values in the textboxes.
        StudentFNameBx.Clear()
        StudentMNameBx.Clear()
        StudentLNameBx.Clear()
        StudentDadNameBx.Clear()
        StudentMomNameBx.Clear()
        StudentBloodGBx.Clear()
        StudentBDayBx.ResetText()
        StudentEmailBx.Clear()
        StudentNumBx.Clear()
        StudentPAddressBx.Clear()
        StudentLAddressBx.Clear()
        StudentGenderBx.ResetText()
        StudentCategBx.ResetText()
        StudentIDBox.Clear()
    End Sub

    Private Sub StudentExitBtn_Click(sender As Object, e As EventArgs) Handles StudentExitBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If

    End Sub

    Private Sub StudentSubmitBtn_Click(sender As Object, e As EventArgs) Handles StudentSubmitBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to Submit/Enter a new student to the MS Database via new data entered in the textboxes.
        insertDT = "INSERT INTO Students_Table 
([Student ID], 
[First Name], 
[Middle Name],
[Last Name],
[Date of Birth],
[Mother's Name], 
[Father's Name], 
[Blood Group] , 
[Local Address], 
[Permanent Address], 
[Gender], 
[Email ID], 
[Mobile Number],  
[Category/Race])   
        VALUES 
('" & StudentIDBox.Text & "', 
'" & StudentFNameBx.Text & "', 
'" & StudentMNameBx.Text & "', 
'" & StudentLNameBx.Text & "',
'" & StudentBDayBx.Value & "',
'" & StudentMomNameBx.Text & "', 
'" & StudentDadNameBx.Text & "', 
'" & StudentBloodGBx.Text & "', 
'" & StudentLAddressBx.Text & "', 
'" & StudentPAddressBx.Text & "', 
'" & StudentGenderBx.Text & "', 
'" & StudentEmailBx.Text & "', 
'" & StudentNumBx.Text & "', 
'" & StudentCategBx.Text & "')"

        'Executing the Query above to the function below.
        Dim cmdInsert As New OleDbCommand(insertDT, con)

        'Function to execute the Query & check authorization, closing the connection, or handle the error if the student couldn't be added.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdInsert.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("New record added successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not save data. Try again later.", "Error")
        End Try

    End Sub

    Private Sub StudentModBtn_Click(sender As Object, e As EventArgs) Handles StudentModBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to update student details to the MS Database via overwriting data entered in the textboxes.
        modifyDT = "UPDATE Students_Table set [First Name] = '" & StudentFNameBx.Text & "', 
        [Middle Name] = '" & StudentMNameBx.Text & "', 
        [Last Name]= '" & StudentLNameBx.Text & "', 
        [Date of Birth]= '" & StudentBDayBx.Value & "', 
        [Mother's Name]= '" & StudentMomNameBx.Text & "', 
        [Father's Name] = '" & StudentDadNameBx.Text & "', 
        [Blood Group] = '" & StudentBloodGBx.Text & "', 
        [Local Address] = '" & StudentLAddressBx.Text & "',
        [Permanent Address]= '" & StudentPAddressBx.Text & "', 
        [Gender]= '" & StudentGenderBx.Text & "', 
        [Email ID]= '" & StudentEmailBx.Text & "', 
        [Mobile Number]= '" & StudentNumBx.Text & "', 
        [Category/Race]= '" & StudentCategBx.Text & "' 
        WHERE [Student ID] = '" & StudentIDBox.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdModify As New OleDbCommand(modifyDT, con)

        'Function to execute the Query, closing the connection, or handle the error if the student couldn't be updated.
        Try
            cmdModify.ExecuteNonQuery()
            con.Dispose()
            con.Close()
            MessageBox.Show("Record updated successfully.", "Success")
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not update record. Try again later.", "Error")
        End Try

    End Sub

    Private Sub StudentDeleteBtn_Click(sender As Object, e As EventArgs) Handles StudentDeleteBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to detlete student details to the MS Database via overwriting data entered in the textboxes.
        deleteDT = "DELETE * FROM Students_Table WHERE [Student ID] = '" & StudentIDBox.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdDelete As New OleDbCommand(deleteDT, con)

        'Function to execute the Query & check authorization to delete, closing the connection, or handle the error if the student couldn't be added.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                Dim result As DialogResult = MsgBox("Are you sure you want to delete this record?", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    cmdDelete.ExecuteNonQuery()
                    con.Dispose()
                    con.Close()
                    MessageBox.Show("Record deleted successfully.", "Success")
                End If
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not find or delete record. Try again later.", "Error")
        End Try

    End Sub

    'Student Section End ---------------------------------------------------------------------------------------------------------------------------------

    'Employee Section Start -------------------------------------------------------------------------------------------------------------------------------
    Private Sub EmployeePrintBtn_Click(sender As Object, e As EventArgs) Handles EmployeePrintBtn.Click

        'Executing the print dialogue to print the employee information.
        PrintDialog.Document = PrintDocument
        PrintDialog.PrinterSettings = PrintDocument.PrinterSettings
        PrintDialog.AllowSomePages = True

        If PrintDialog.ShowDialog = DialogResult.OK Then
            PrintDocument.PrinterSettings = PrintDialog.PrinterSettings
            PrintDocument.Print()
        End If
    End Sub

    Private Sub EmployeeSubmitBtn_Click(sender As Object, e As EventArgs) Handles EmployeeSubmitBtn.Click

        'Establishing And opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to Submit/Enter a new employee to the MS Database via new data entered in the textboxes.
        insertDT = "INSERT INTO Employee_Table 
([Employee ID], 
[Employee Title], 
[Employee First Name],
[Employee Last Name],
[Date of Birth],
[Mobile Number], 
[Phone Number],  
[Email],
[Permanent Address], 
[Local Address], 
[Employee Designation], 
[Employee Type], 
[Salary])   
        VALUES 
('" & EmployeeIDBx.Text & "', 
'" & EmployeeTitleBx.Text & "', 
'" & EmployeeFNameBx.Text & "', 
'" & EmployeeLNameBx.Text & "',
'" & EmployeeBDayBx.Value & "',
'" & EmployeeMobileBx.Text & "', 
'" & EmployeeTelBx.Text & "', 
'" & EmployeeEmailBx.Text & "', 
'" & EmployeePAddressBx.Text & "', 
'" & EmployeeLAddressBx.Text & "', 
'" & EmployeeDesigBx.Text & "', 
'" & EmployeeTypeBx.Text & "', 
'" & EmployeeSalaryBx.Text & "')"

        'Executing the Query above to the function below.
        Dim cmdInsert As New OleDbCommand(insertDT, con)

        'Function to execute the Query & check authorization, closing the connection, or handle the error if the student couldn't be added.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdInsert.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("New employee record added successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not save new employee record. Try again later.", "Error")
        End Try


    End Sub

    Private Sub EmployeeFindBtn_Click(sender As Object, e As EventArgs) Handles EmployeeFindBtn.Click

        'Query to find the entered Employee ID from the textbox in the MS DB
        findDT = ("SELECT * from Employee_Table WHERE [Employee ID] = '" & (EmployeeIDBx.Text) & "'")

        'Establishing and opening a connection to the DB where the table is
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        cmd = New OleDbCommand(findDT, con)
        con.Open()

        'Executing the Query above to the function below.
        Dim myRead As OleDb.OleDbDataReader = cmd.ExecuteReader

        'Function for finding an employee and all their details via entering a Employee ID, and returning the values in the textboxes.
        With myRead
            If myRead.Read() Then
                EmployeeTitleBx.Text = myRead("Employee Title")
                EmployeeFNameBx.Text = myRead("Employee First Name")
                EmployeeLNameBx.Text = myRead("Employee Last Name")
                EmployeeBDayBx.Text = myRead("Date of Birth")
                EmployeeMobileBx.Text = myRead("Mobile Number")
                EmployeeTelBx.Text = myRead("Phone Number")
                EmployeeEmailBx.Text = myRead("Email")
                EmployeePAddressBx.Text = myRead("Permanent Address")
                EmployeeLAddressBx.Text = myRead("Local Address")
                EmployeeDesigBx.Text = myRead("Employee Designation")
                EmployeeTypeBx.Text = myRead("Employee Type")
                EmployeeSalaryBx.Text = myRead("Salary")
            Else
                'Error handling for the scenario of a employee not found/non existent.
                MessageBox.Show("Invalid user or record not found.", "Invalid Entry")
            End If
        End With

        'Closing the connection to the DB where the table is.
        con.Close()

    End Sub

    Private Sub EmployeeModBtn_Click(sender As Object, e As EventArgs) Handles EmployeeModBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to update employee details to the MS Database via overwriting data entered in the textboxes.
        modifyDT = "UPDATE Employee_Table set [Employee Title] = '" & EmployeeTitleBx.Text & "', 
        [Employee First Name] = '" & EmployeeFNameBx.Text & "', 
        [Employee Last Name]= '" & EmployeeLNameBx.Text & "', 
        [Date of Birth]= '" & EmployeeBDayBx.Value & "', 
        [Mobile Number]= '" & EmployeeMobileBx.Text & "', 
        [Phone Number] = '" & EmployeeTelBx.Text & "', 
        [Email] = '" & EmployeeEmailBx.Text & "', 
        [Permanent Address]= '" & EmployeePAddressBx.Text & "', 
        [Local Address] = '" & EmployeeLAddressBx.Text & "', 
        [Employee Designation]= '" & EmployeeDesigBx.Text & "', 
        [Employee Type]= '" & EmployeeTypeBx.Text & "', 
        [Salary]= '" & EmployeeSalaryBx.Text & "'
        WHERE [Employee ID] = '" & EmployeeIDBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdModify As New OleDbCommand(modifyDT, con)

        'Function to execute the Query, closing the connection, or handle the error if the employee record couldn't be updated.
        Try
            cmdModify.ExecuteNonQuery()
            con.Dispose()
            con.Close()
            MessageBox.Show("Record updated successfully.", "Success")
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not update record. Try again later.", "Error")
        End Try

    End Sub

    Private Sub EmployeeDelBtn_Click(sender As Object, e As EventArgs) Handles EmployeeDelBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to delete employee details to the MS Database via overwriting data entered in the textboxes.
        deleteDT = "DELETE * FROM Employee_Table WHERE [Employee ID] = '" & EmployeeIDBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdDelete As New OleDbCommand(deleteDT, con)


        'Function to execute the Query & check authorization to delete, closing the connection, or handle the error if the employee couldn't be removed.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                Dim result As DialogResult = MsgBox("Are you sure you want to delete this record?", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    cmdDelete.ExecuteNonQuery()
                    con.Dispose()
                    con.Close()
                    MessageBox.Show("Record deleted successfully.", "Success")
                End If
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not find or delete record. Try again later.", "Error")
        End Try

    End Sub

    Private Sub EmployeeClearBtn_Click(sender As Object, e As EventArgs) Handles EmployeeClearBtn.Click

        'Function to clear all employee details and values in the textboxes.
        EmployeeTitleBx.ResetText()
        EmployeeFNameBx.Clear()
        EmployeeLNameBx.Clear()
        EmployeeMobileBx.Clear()
        EmployeeTelBx.Clear()
        EmployeeEmailBx.Clear()
        EmployeeBDayBx.ResetText()
        EmployeePAddressBx.Clear()
        EmployeeLAddressBx.Clear()
        EmployeeDesigBx.Clear()
        EmployeeTypeBx.Clear()
        EmployeeSalaryBx.Clear()
        EmployeeIDBx.Clear()
    End Sub

    Private Sub EmployeeExtBtn_Click(sender As Object, e As EventArgs) Handles EmployeeExtBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If
    End Sub

    'Employee Section End -------------------------------------------------------------------------------------------------------------------------------

    'Salary Section Start -------------------------------------------------------------------------------------------------------------------------------
    Private Sub SalaryPrintBtn_Click(sender As Object, e As EventArgs) Handles SalaryPrintBtn.Click

        'Executing the print dialogue to print the salary information.
        PrintDialog.Document = PrintDocument
        PrintDialog.PrinterSettings = PrintDocument.PrinterSettings
        PrintDialog.AllowSomePages = True

        If PrintDialog.ShowDialog = DialogResult.OK Then
            PrintDocument.PrinterSettings = PrintDialog.PrinterSettings
            PrintDocument.Print()
        End If

    End Sub

    Private Sub SalaryExitBtn_Click(sender As Object, e As EventArgs) Handles SalaryExitBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If

    End Sub


    Private Sub SalarySaveBtn_Click(sender As Object, e As EventArgs) Handles SalarySaveBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to save the entered Salary details from the textbox to the MS DB
        modifyDT = "UPDATE Salary_Table set [Salary ID] = '" & SalaryIDBx.Text & "', 
        [Employee ID] = '" & EmployeeIDSalBx.Text & "', 
        [Amount]= '" & SalaryAmountBx.Text & "', 
        [Salary Month]= '" & SalaryMonthBx.Text & "', 
        [Is Paid]= '" & SalaryPaidBx.Text & "', 
        [Due Amount] = '" & SalaryDueBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdModify As New OleDbCommand(modifyDT, con)

        'Function to execute the Query, closing the connection, or handle the error if the salary record couldn't be updated.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdModify.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("Salary record saved successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not save record. Try again later.", "Error")
        End Try

        'Closing the connection to the DB where the table is.
        con.Close()
    End Sub

    Private Sub SalarySearchBtn_Click(sender As Object, e As EventArgs) Handles SalarySearchBtn.Click

        'Query to find the Salary Details from Employee ID from the textbox in the MS DB
        findDT = ("SELECT * from Salary_Table WHERE [Employee ID] = '" & (EmployeeIDSalBx.Text) & "'")

        'Establishing and opening a connection to the DB where the table is
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        cmd = New OleDbCommand(findDT, con)
        con.Open()

        'Executing the Query above to the function below.
        Dim myRead As OleDb.OleDbDataReader = cmd.ExecuteReader


        'Function for finding an employee and all their salary details via entering a Employee ID, and returning the values in the textboxes & datagridview.
        With myRead
            If myRead.Read() Then
                SalaryIDBx.Text = myRead("Salary ID")
                EmployeeIDSalBx.Text = myRead("Employee ID")
                SalaryAmountBx.Text = myRead("Amount")
                SalaryMonthBx.Text = myRead("Salary Month")
                SalaryPaidBx.Text = myRead("Is Paid")
                SalaryDueBx.Text = myRead("Due Amount")
                SalaryDataGrid.Rows(0).Cells(0).Value = SalaryIDBx.Text
                SalaryDataGrid.Rows(0).Cells(1).Value = EmployeeIDSalBx.Text
                SalaryDataGrid.Rows(0).Cells(2).Value = SalaryAmountBx.Text
                SalaryDataGrid.Rows(0).Cells(3).Value = SalaryMonthBx.Text
                SalaryDataGrid.Rows(0).Cells(4).Value = SalaryPaidBx.Text
                SalaryDataGrid.Rows(0).Cells(5).Value = SalaryDueBx.Text
                SalaryDataGrid.Rows(0).Cells(6).Value = "-"
            Else
                'Error handling for the scenario of a employee record not found/non existent.
                MessageBox.Show("Invalid employee ID or record not found.", "Invalid Entry")
            End If
        End With

        'Closing the connection to the DB where the table is.
        con.Close()
    End Sub

    Private Sub SalaryDeleteBtn_Click(sender As Object, e As EventArgs) Handles SalaryDeleteBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to detlete salary details to the MS Database via overwriting data entered in the textboxes.
        deleteDT = "DELETE * FROM Salary_Table WHERE [Salary ID] = '" & SalaryIDBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdDelete As New OleDbCommand(deleteDT, con)
        'Function to execute the Query & check authorization to delete, closing the connection, or handle the error if the salary couldn't be removed.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                Dim result As DialogResult = MsgBox("Are you sure you want to delete this record?", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    cmdDelete.ExecuteNonQuery()
                    con.Dispose()
                    con.Close()
                    MessageBox.Show("Record deleted successfully.", "Success")
                End If
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not find or delete record. Try again later.", "Error")
        End Try


    End Sub

    'Salary Section End -------------------------------------------------------------------------------------------------------------------------------

    'Course Section Start -------------------------------------------------------------------------------------------------------------------------------

    Private Sub CourseExitBtn_Click(sender As Object, e As EventArgs) Handles CourseExitBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If

    End Sub

    Private Sub CourseClearBtn_Click(sender As Object, e As EventArgs) Handles CourseClearBtn.Click

        'Function to clear all course details and values in the textboxes.
        CourseIDBx.Clear()
        CourseTitleBx.Clear()
        CourseNameBx.Clear()
        CourseCodeBx.Clear()
        CourseFeeBx.Clear()
        CourseDurationBx.Clear()

    End Sub

    Private Sub CourseFindBtn_Click(sender As Object, e As EventArgs) Handles CourseFindBtn.Click

        'Query to find the entered Course ID from the textbox in the MS DB
        findDT = ("SELECT * from Courses_Table WHERE [Course ID] = '" & (CourseIDBx.Text) & "'")

        'Establishing and opening a connection to the DB where the table is
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        cmd = New OleDbCommand(findDT, con)
        con.Open()

        'Executing the Query above to the function below.
        Dim myRead As OleDb.OleDbDataReader = cmd.ExecuteReader

        'Function for finding a course and all their details via entering a Course ID, and returning the values in the textboxes.
        With myRead
            If myRead.Read() Then
                CourseIDBx.Text = myRead("Course ID")
                CourseTitleBx.Text = myRead("Course Title")
                CourseNameBx.Text = myRead("Course Name")
                CourseCodeBx.Text = myRead("Course Code")
                CourseFeeBx.Text = myRead("Course Fee")
                CourseDurationBx.Text = myRead("Course Duration")

            Else
                'Error handling for the scenario of a student not found/non existent.
                MessageBox.Show("Invalid course or record not found.", "Invalid Entry")
            End If
        End With

        'Closing the connection to the DB where the table is.
        con.Close()

    End Sub

    Private Sub CourseModBtn_Click(sender As Object, e As EventArgs) Handles CourseModBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to update course details to the MS Database via overwriting data entered in the textboxes.
        modifyDT = "UPDATE Courses_Table set [Course ID] = '" & CourseIDBx.Text & "', 
        [Course Title] = '" & CourseTitleBx.Text & "', 
        [Course Name]= '" & CourseNameBx.Text & "', 
        [Course Code]= '" & CourseCodeBx.Text & "', 
        [Course Fee]= '" & CourseFeeBx.Text & "', 
        [Course Duration] = '" & CourseDurationBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdModify As New OleDbCommand(modifyDT, con)

        'Function to execute the Query, closing the connection, or handle the error if the employee record couldn't be updated.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdModify.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("Record updated successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not update record. Try again later.", "Error")
        End Try

    End Sub

    Private Sub CourseSubmitBtn_Click(sender As Object, e As EventArgs) Handles CourseSubmitBtn.Click

        'Establishing And opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to Submit/Enter a new course to the MS Database via new data entered in the textboxes.
        insertDT = "INSERT INTO Courses_Table 
([Course ID], 
[Course Title], 
[Course Name],
[Course Code],
[Course Fee],
[Course Duration])   
        VALUES 
('" & CourseIDBx.Text & "', 
'" & CourseTitleBx.Text & "', 
'" & CourseNameBx.Text & "', 
'" & CourseCodeBx.Text & "',
'" & CourseFeeBx.Text & "',
'" & CourseDurationBx.Text & "')"

        'Executing the Query above to the function below.
        Dim cmdInsert As New OleDbCommand(insertDT, con)

        'Function to execute the Query, closing the connection, or handle the error if the course couldn't be added.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdInsert.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("New course record added successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not save new course record. Try again later.", "Error")
        End Try
    End Sub

    Private Sub CourseDeleteBtn_Click(sender As Object, e As EventArgs) Handles CourseDeleteBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to delete course details to the MS Database via overwriting data entered in the textboxes.
        deleteDT = "DELETE * FROM Courses_Table WHERE [Course ID] = '" & CourseIDBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdDelete As New OleDbCommand(deleteDT, con)

        'Function to execute the Query & check authorization to delete, closing the connection, or handle the error if the course couldn't be removed.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                Dim result As DialogResult = MsgBox("Are you sure you want to delete this record?", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    cmdDelete.ExecuteNonQuery()
                    con.Dispose()
                    con.Close()
                    MessageBox.Show("Record deleted successfully.", "Success")
                End If
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not find or delete record. Try again later.", "Error")
        End Try


    End Sub

    'Course Section End -------------------------------------------------------------------------------------------------------------------------------

    'Fee Section Start -------------------------------------------------------------------------------------------------------------------------------

    Private Sub FeesExitBtn_Click(sender As Object, e As EventArgs) Handles FeesExitBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If

    End Sub

    Private Sub FeesClearBtn_Click(sender As Object, e As EventArgs) Handles FeesClearBtn.Click

        'Function to clear all fee details and values in the textboxes.
        FeesIDBx.Clear()
        FeesStudentIDBx.Clear()
        PaymentTypeBx.ResetText()
        FeesDayBx.ResetText()
        FeesMonthBx.ResetText()
        FeesYearBx.ResetText()
        FeesAmountBx.Clear()

    End Sub

    Private Sub FeesSubmitBtn_Click(sender As Object, e As EventArgs) Handles FeesSubmitBtn.Click

        'Establishing And opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to Submit/Enter new fees to the MS Database via new data entered in the textboxes.
        insertDT = "INSERT INTO Fees_Table 
([Fees ID], 
[Student ID], 
[Payment Method],
[Fees Date],
[Fees Month],
[Fees Year],
[Fees Amount])   
        VALUES 
('" & FeesIDBx.Text & "', 
'" & FeesStudentIDBx.Text & "', 
'" & PaymentTypeBx.Text & "', 
'" & FeesDayBx.Text & "',
'" & FeesMonthBx.Text & "',
'" & FeesYearBx.Text & "',
'" & FeesAmountBx.Text & "')"

        'Executing the Query above to the function below.
        Dim cmdInsert As New OleDbCommand(insertDT, con)


        'Function to execute the Query, closing the connection, or handle the error if the fees couldn't be added.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdInsert.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("New fees record added successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not save new fees record. Try again later.", "Error")
        End Try

    End Sub

    Private Sub FeesDeleteBtn_Click(sender As Object, e As EventArgs) Handles FeesDeleteBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to delete fee details to the MS Database via overwriting data entered in the textboxes.
        deleteDT = "DELETE * FROM Fees_Table WHERE [Fees ID] = '" & FeesIDBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdDelete As New OleDbCommand(deleteDT, con)

        'Function to execute the Query & check authorization to delete, closing the connection, or handle the error if the fees couldn't be removed.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                Dim result As DialogResult = MsgBox("Are you sure you want to delete this record?", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    cmdDelete.ExecuteNonQuery()
                    con.Dispose()
                    con.Close()
                    MessageBox.Show("Record deleted successfully.", "Success")
                End If
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not find or delete record. Try again later.", "Error")
        End Try

    End Sub

    Private Sub FeesFindBtn_Click(sender As Object, e As EventArgs) Handles FeesFindBtn.Click

        'Query to find the entered Fees for the Student ID from the textbox in the MS DB
        findDT = ("SELECT * from Fees_Table WHERE [Student ID] = '" & (FeesStudentIDBx.Text) & "'")

        'Establishing and opening a connection to the DB where the table is
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        cmd = New OleDbCommand(findDT, con)
        con.Open()

        'Executing the Query above to the function below.
        Dim myRead As OleDb.OleDbDataReader = cmd.ExecuteReader

        'Function for finding fees and all their details via entering a Student ID, and returning the values in the textboxes.
        With myRead
            If myRead.Read() Then
                FeesIDBx.Text = myRead("Fees ID")
                FeesStudentIDBx.Text = myRead("Student ID")
                PaymentTypeBx.Text = myRead("Payment Method")
                FeesDayBx.Text = myRead("Fees Date")
                FeesMonthBx.Text = myRead("Fees Month")
                FeesYearBx.Text = myRead("Fees Year")
                FeesAmountBx.Text = myRead("Fees Amount")
            Else
                'Error handling for the scenario of a student not found/non existent.
                MessageBox.Show("Invalid ID or record not found.", "Invalid Entry")
            End If
        End With

        'Closing the connection to the DB where the table is.
        con.Close()

    End Sub

    Private Sub FeesModBtn_Click(sender As Object, e As EventArgs) Handles FeesModBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
        con.Open()

        'Query to update student details to the MS Database via overwriting data entered in the textboxes.
        modifyDT = "UPDATE Fees_Table set [Fees ID] = '" & FeesIDBx.Text & "', 
        [Student ID] = '" & FeesStudentIDBx.Text & "', 
        [Payment Method]= '" & PaymentTypeBx.Text & "', 
        [Fees Date]= '" & FeesDayBx.Text & "', 
        [Fees Month]= '" & FeesMonthBx.Text & "', 
        [Fees Year] = '" & FeesYearBx.Text & "', 
        [Fees Amount] = '" & FeesAmountBx.Text & "'"

        'Executing the Query above to the function below.
        Dim cmdModify As New OleDbCommand(modifyDT, con)

        'Function to execute the Query and check verification, closing the connection, or handle the error if the fees couldn't be updated.
        Try
            Dim passMsg As String
            passMsg = InputBox("Enter verifivation passcode to proceed.", "Verify Authorization", "Enter passcode here", 450, 450)
            If passMsg = passcode Then
                cmdModify.ExecuteNonQuery()
                con.Dispose()
                con.Close()
                MessageBox.Show("Record updated successfully.", "Success")
            Else
                MessageBox.Show("Passcode Incorrect. Try again later.", "Unauthorized Access")
                Loading.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MessageBox.Show("Could not update record. Try again later.", "Error")
        End Try


    End Sub

    'Fee Section End -------------------------------------------------------------------------------------------------------------------------------

    'Profile Section Start -------------------------------------------------------------------------------------------------------------------------------
    Private Sub PswChangeBtn_Click(sender As Object, e As EventArgs) Handles PswChangeBtn.Click

        'Function to ensure that the new password isn't too short or too long.
        If NewPswordBox.Text.Length < 4 And NewPswordBox.Text.Length > 12 Then
            MessageBox.Show("New Password should not be longer that 12 characters, nor shorter than 4 characters.", "Invalid Information")
            UserNameBx.Text = ""
            OldPswordBox.Text = ""
            NewPswordBox.Text = ""
            ConPswordBox.Text = ""

            'Function to ensure that the new password is not the same as old password.
        ElseIf OldPswordBox.Text = NewPswordBox.Text Then
            MessageBox.Show("New Password cannot be the same as Old Password. Try again.", "Invalid Information")
            UserNameBx.Text = ""
            OldPswordBox.Text = ""
            NewPswordBox.Text = ""
            ConPswordBox.Text = ""

            'Function to that new password is the same as the confirmed password, and execute the query.
        ElseIf (NewPswordBox.Text = ConPswordBox.Text) Then
            Try
                con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb")
                con.Open()
                cmd = New OleDbCommand("UPDATE User_Table set [Password] = '" + NewPswordBox.Text + "' where [Username] = '" + UserNameBx.Text + "'", con)
                cmd.ExecuteNonQuery()
                MessageBox.Show("User Password updated successfully.", "Successful Update")
                con.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            'Error handling for when password confirmation does not match the new password.
        Else
            MsgBox("Confirmed password does not match. Try again.", "Invalid Information")
            UserNameBx.Text = ""
            OldPswordBox.Text = ""
            NewPswordBox.Text = ""
            ConPswordBox.Text = ""
        End If
        con.Close()

    End Sub

    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles ExtBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If

    End Sub

    Private Sub clearPswBtn_Click(sender As Object, e As EventArgs) Handles clearPswBtn.Click

        'Function to clear all textboxes once clicked.
        UserNameBx.Clear()
        OldPswordBox.Clear()
        NewPswordBox.Clear()
        ConPswordBox.Clear()

    End Sub

    'Profile Section End -------------------------------------------------------------------------------------------------------------------------------

End Class