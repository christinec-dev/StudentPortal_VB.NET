Public Class AboutForm

    Private Sub ExitBtn_Click(sender As Object, e As EventArgs) Handles ExitBtn.Click

        'Closes the about form and executes the Loading form to go back to the Control Form.
        Me.Close()
        Loading.Show()

    End Sub

    Private Sub AboutForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class