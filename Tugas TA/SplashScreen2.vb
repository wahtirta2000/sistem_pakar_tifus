Public NotInheritable Class SplashScreen2
    Dim OpacityRate As Double = 0.0
    Dim MaximizeRate As Boolean = True
    Dim MinimizeRate As Boolean = False

    Private Sub SplashScreen2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Opacity = 0.0
        Timer1.Interval = 60
        Timer1.Enabled = True
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If OpacityRate >= 1.0 Then
            OpacityRate = OpacityRate + 1.0
            If OpacityRate >= 20.0 Then
                OpacityRate = 0.99
                Me.Opacity = OpacityRate
            End If
        ElseIf MaximizeRate Then
            OpacityRate = OpacityRate + 0.025
            Me.Opacity = OpacityRate
            If OpacityRate >= 1.0 Then
                MaximizeRate = False
                MinimizeRate = True
            End If
        ElseIf MinimizeRate Then
            OpacityRate = OpacityRate - 0.025
            If OpacityRate < 0 Then
                OpacityRate = 0
            End If
            Me.Opacity = OpacityRate
            If Opacity <= 0.0 Then
                MinimizeRate = False
                MaximizeRate = False
            End If
        Else
            Timer1.Stop()
            Timer1.Enabled = False
            Timer1.Dispose()

            Me.Visible = False
            Dim Login As New Form1
            Login.Show()
        End If
    End Sub

End Class
