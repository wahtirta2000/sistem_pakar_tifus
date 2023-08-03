Public NotInheritable Class SplashScreen1

    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            With Timer1
                .Enabled = True

            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 1
        If ProgressBar1.Value >= ProgressBar1.Maximum Then
            Timer1.Enabled = False
            Me.Close()
            Form2.Show()
            Me.Hide()
        End If
    End Sub
End Class
