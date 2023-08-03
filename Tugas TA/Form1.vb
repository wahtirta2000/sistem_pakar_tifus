Imports System.Data
Imports System.Data.OleDb
Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()

        ComboBox1.Text = "- - - "
        ComboBox1.Items.Add("User")
        ComboBox1.Items.Add("Pakar")

        TextBox1.MaxLength = 10
        TextBox2.MaxLength = 10
        TextBox2.PasswordChar = ""

        TextBox2.UseSystemPasswordChar = True
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim a As Char
        a = e.KeyChar
        If a = Chr(13) Then
            'validasi kosong

            If TextBox1.Text = "" Then
                MsgBox("Isi Sesuai Account Atau Username Yang Anda Punya", MsgBoxStyle.Exclamation, "Warning")
                Exit Sub
            End If
            TextBox2.Focus()
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        Dim b As Char
        b = e.KeyChar
        If b = Chr(13) Then

            If TextBox2.Text = "" Then
                MsgBox("Isi Dengan Password Yang Benar", MsgBoxStyle.Exclamation, "Warning")
                Exit Sub
            End If
            Button1.Focus()
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim Jam, Tanggal As String
        Jam = Format(Now, "hh:mm:ss")
        Tanggal = Format(Now, "dddd dd-MM-yyyy")
        Label5.Text = " " & Jam & ", " & Tanggal & ""
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            TextBox2.UseSystemPasswordChar = True
        Else
            TextBox2.UseSystemPasswordChar = False
        End If
    End Sub
    Sub Login()
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Account dan Password Yang Anda Masuki Belum Lengkap", MsgBoxStyle.Exclamation, "Warning")
            TextBox1.Focus()
            Exit Sub
        Else
            Call Koneksi()
            Cmd = New OleDbCommand("Select * From Pengguna_Aplikasi Where Pengguna='" & TextBox1.Text & "' and Password='" & TextBox2.Text & "' and Level='" & ComboBox1.Text & "'", Connection_db)
            D_reader = Cmd.ExecuteReader
            D_reader.Read()
            If D_reader.HasRows Then
                Me.Visible = False

                SplashScreen1.Show()
                Me.Hide()

                Form2.ToolStripStatusLabel2.Text = D_reader.Item("Pengguna")
                Form2.ToolStripStatusLabel4.Text = D_reader.Item("Level")

                Form2.ToolStripStatusLabel6.Text = Format(Now.Date, "dd MMM yyyy")

                If Form2.ToolStripStatusLabel4.Text = "User" Then
                    Form2.DataPenggunaAplikasiToolStripMenuItem.Enabled = False
                    Form2.DataPengetahuanPakarToolStripMenuItem.Enabled = False
                    Form2.DataPengetahuanPakarToolStripMenuItem1.Enabled = False
                    Form2.InputJenisKerusakanToolStripMenuItem.Enabled = False
                    Form2.InputSolusiPermasalahanToolStripMenuItem.Enabled = False

                    Form2.ManagePenggunaAplikasiToolStripMenuItem.Enabled = False

                    Form2.Button1.Enabled = True
                    Form2.Button2.Enabled = False
                    Form2.Button3.Enabled = False
                    Form2.Button4.Enabled = False
                    Form2.Button5.Enabled = True
                    Form2.Button6.Enabled = True
                    Form2.Button7.Enabled = True
                    Form2.Button8.Enabled = True
                    Form2.Button9.Enabled = True
                    Form2.Button10.Enabled = False
                    Form2.Button11.Enabled = False
                Else
                    If Form2.ToolStripStatusLabel4.Text = "Pakar" Then
                        Form2.DataPengetahuanPakarToolStripMenuItem.Enabled = True
                        Form2.DataPengetahuanPakarToolStripMenuItem1.Enabled = True
                        Form2.InputJenisKerusakanToolStripMenuItem.Enabled = True
                        Form2.InputSolusiPermasalahanToolStripMenuItem.Enabled = True

                        Form2.ManagePenggunaAplikasiToolStripMenuItem.Enabled = True

                        Form2.Button1.Enabled = True
                        Form2.Button2.Enabled = True
                        Form2.Button5.Enabled = True
                        Form2.Button6.Enabled = True
                        Form2.Button7.Enabled = True
                        Form2.Button8.Enabled = True
                        Form2.Button9.Enabled = True
                        Form2.Button10.Enabled = True
                        Form2.Button11.Enabled = True
                    End If
                End If
            Else
                MsgBox("Login Anda Salah, Harap Ulangi lagi", MsgBoxStyle.Exclamation, "Warning")
                TextBox1.Clear()
                TextBox2.Clear()

                TextBox1.Focus()
            End If
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Login()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

End Class
