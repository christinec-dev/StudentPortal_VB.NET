Imports System.Data.OleDb
Public Class LoginForm

    'Global DB variables to be used between Private Subs.
    Dim DbCon As New OleDb.OleDbConnection
    Dim dbUp As New OleDb.OleDbCommand
    Dim Read As OleDb.OleDbDataReader

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Clears the Username and Password Textboxes when the form loads up.
        UsernameBox.Clear()
        PswordBox.Clear()

    End Sub

    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click

        'Function to exit the application, but ensuring the user is sure before proceeding.
        If MessageBox.Show("Are you sure that you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else
        End If

    End Sub

    Private Sub LoginBtn_Click(sender As Object, e As EventArgs) Handles LoginBtn.Click

        'Establishing and opening a connection to the DB where the table is.
        DbCon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\HP Notebook 15\Desktop\Login_DB_RGIT.accdb"
        DbCon.Open()
        dbUp.Connection = DbCon
        dbUp.CommandType = CommandType.Text

        'Query to find the password and username from the DB via new data entered in the textboxes.
        dbUp.CommandText = "SELECT Username, Password from User_Table WHERE [Username] = '" & Trim(UsernameBox.Text) & "'  AND [Password] = '" & Trim(PswordBox.Text) & "' "
        dbUp.Parameters.Add("Username", Data.OleDb.OleDbType.Variant)
        dbUp.Parameters.Add("Password", Data.OleDb.OleDbType.Variant)
        dbUp.Parameters("Username").Value = UsernameBox.Text
        dbUp.Parameters("Password").Value = PswordBox.Text
        Read = dbUp.ExecuteReader

        'Function to execute the Query, closing the connection, or handle the error if the password/username is incorrect.
        With Read
            If .Read Then
                Me.Hide()
                Loading.Show()
            Else
                UsernameBox.Clear()
                PswordBox.Clear()
                MessageBox.Show("Invalid Username/Password Entered")
                UsernameBox.Focus()
            End If
        End With

        'Closing the DB connection
        DbCon.Close()

    End Sub

    Private Sub login_title_Click(sender As Object, e As EventArgs) Handles login_title.Click

    End Sub
End Class
