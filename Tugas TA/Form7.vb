Imports System.Data
Imports System.Data.OleDb
Public Class Form7
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Pengguna_Aplikasi where Pengguna='" & TextBox4.Text & "' and Password='" & TextBox1.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If Not D_reader.HasRows Then
            MsgBox("Password Lama Anda Tidak Sesuai Dengan Yang Ada Didalam Database Kami", MsgBoxStyle.Exclamation, "Warning")
            TextBox1.Clear()
            TextBox1.Focus()
        ElseIf TextBox2.Text = "" Then
            MsgBox("Isi Password Baru, Tidak Diperbolehkan Kosong", MsgBoxStyle.Exclamation, "Warning")
            TextBox2.Focus()
            TextBox2.Clear()
        ElseIf TextBox3.Text <> TextBox2.Text Then
            MsgBox("Konfirmasi Password Anda Berbeda Dengan Yang Sebelumnya", MsgBoxStyle.Exclamation, "Warning")
            TextBox3.Focus()
            TextBox3.Clear()
        Else
            Dim Ubah As String = "UPDATE Pengguna_Aplikasi SET [Password] = '" & TextBox3.Text & "' WHERE [Pengguna] = '" & TextBox4.Text & "'"
            Cmd = New OleDbCommand(Ubah, Connection_db)
            Cmd.ExecuteNonQuery()
            MsgBox("Password Anda Berhasil Diubah Dalam Database", MsgBoxStyle.Information, "Notification")
            Call kosong()
            TextBox1.Focus()
            Button1.Refresh()
        End If
    End Sub
    Sub kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Form7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        TextBox4.Enabled = False

        TextBox1.MaxLength = 10
        TextBox2.MaxLength = 10
        TextBox3.MaxLength = 10

        TextBox1.PasswordChar = ""
        TextBox2.PasswordChar = ""
        TextBox3.PasswordChar = ""

        TextBox1.UseSystemPasswordChar = True
        TextBox2.UseSystemPasswordChar = True
        TextBox3.UseSystemPasswordChar = True

        MaximizeBox = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class