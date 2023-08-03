Imports System.Data
Imports System.Data.OleDb
Public Class Form2

    Sub Pengguna()
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Pengguna_Aplikasi where Level='" & ToolStripStatusLabel4.Text & "' and Pengguna='" & ToolStripStatusLabel2.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
    End Sub
    Private Sub ProfilToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProfilToolStripMenuItem.Click
        Call Pengguna()
        If D_reader.HasRows Then
            Form3.Show()
            Form3.TextBox1.Text = D_reader.Item("Pengguna")
            Form3.TextBox2.Text = D_reader.Item("Password")
            Form3.TextBox3.Text = D_reader.Item("Level")

            If Form3.TextBox3.Text = "User" Then
                Form3.TextBox1.Enabled = False
                Form3.TextBox2.Enabled = False
                Form3.TextBox3.Enabled = False
            End If
            If Form3.TextBox3.Text = "Pakar" Then
                Form3.TextBox1.Enabled = False
                Form3.TextBox2.Enabled = False
                Form3.TextBox3.Enabled = False
            End If
        End If
    End Sub
    Private Sub ProgramAplikasiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgramAplikasiToolStripMenuItem.Click
        Form8.Show()
    End Sub
    Private Sub IdentitasPerancangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IdentitasPerancangToolStripMenuItem.Click
        Form9.Show()
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel8.Text = Format(Now, "hh:mm:ss")
    End Sub
    Private Sub TambahPenggunaAplikasiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Pengguna()
        If D_reader.HasRows Then
            Form4.Show()
        End If
    End Sub
    Private Sub UbahPasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UbahPasswordToolStripMenuItem.Click
        Call Pengguna()
        If D_reader.HasRows Then
            Form7.Show()
            Form7.TextBox1.Focus()
            Form7.TextBox4.Text = D_reader.Item("Pengguna")
            If Form7.TextBox4.Text = "User" Then
                Form7.TextBox4.Enabled = False
            End If
        End If
    End Sub

    Private Sub LogoutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem1.Click
        Me.Dispose()
        My.Forms.Form1.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Dim Tanya = MessageBox.Show("Apakah Anda Yakin, Ingin Keluar Dari Aplikasi Ini ?", "Notification", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Tanya = Windows.Forms.DialogResult.Yes Then
            End
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click, Button14.Click
        Form9.Show()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click, Button13.Click
        Form8.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call Pengguna()
        If D_reader.HasRows Then
            Form7.Show()
            Form7.TextBox1.Focus()
            Form7.TextBox4.Text = D_reader.Item("Pengguna")
            If Form7.TextBox4.Text = "User" Then
                Form7.TextBox4.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
            Form4.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Pengguna()
        If D_reader.HasRows Then
            Form3.Show()
            Form3.TextBox1.Text = D_reader.Item("Pengguna")
            Form3.TextBox2.Text = D_reader.Item("Password")
            Form3.TextBox3.Text = D_reader.Item("Level")

            If Form3.TextBox3.Text = "User" Then
                Form3.TextBox1.Enabled = False
                Form3.TextBox2.Enabled = False
                Form3.TextBox3.Enabled = False
            End If
            If Form3.TextBox3.Text = "Pakar" Then
                Form3.TextBox1.Enabled = False
                Form3.TextBox2.Enabled = False
                Form3.TextBox3.Enabled = False
            End If
        End If
    End Sub

    Private Sub DataPengetahuanPakarToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPengetahuanPakarToolStripMenuItem1.Click
        Form10.Show()
    End Sub

    Private Sub InputJenisKerusakanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputJenisKerusakanToolStripMenuItem.Click
        Form11.Show()
    End Sub


    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click, Button4.Click
        Form11.Show()
    End Sub


    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click, Button3.Click
        Form10.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click, Button12.Click
        Form14.Show()
    End Sub

    Private Sub BantuanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BantuanToolStripMenuItem.Click
        Form13.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click, Button15.Click
        Form13.Show()
    End Sub

    Private Sub PendiagnosaanKerusakanSepedaMotorKarbulatorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PendiagnosaanKerusakanSepedaMotorKarbulatorToolStripMenuItem.Click
        Form14.Show()
    End Sub

    Private Sub ManagePenggunaAplikasiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManagePenggunaAplikasiToolStripMenuItem.Click
        Form4.Show()
    End Sub


End Class