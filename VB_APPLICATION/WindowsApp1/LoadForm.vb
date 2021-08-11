Public Class Loading
    Private Sub Loading_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Creating a timer and allocating a time for it to load up the progressbar.
        Timer1.Interval = 50
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 100
        Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'Function to execute the timer to move the progressbar so it can load the Control Form.
        If ProgressBar1.Value > 0 Then
            ProgressBar1.Value = ProgressBar1.Value - 1
        Else
            ProgressBar1.Value = 100
            Me.Hide()
            ControlForm.Show()
        End If
    End Sub

End Class