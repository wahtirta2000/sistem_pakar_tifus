Imports System.Data
Imports System.Data.OleDb

Public Class Form3

    Sub Pengguna()
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Pengguna_Aplikasi where Level='" & Form2.ToolStripStatusLabel4.Text & "' and Pengguna='" & Form2.ToolStripStatusLabel2.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
    End Sub
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Pengguna()
        If D_reader.HasRows Then
            TextBox1.Text = D_reader.Item("Pengguna")
            TextBox2.Text = D_reader.Item("Password")
            TextBox3.Text = D_reader.Item("Level")

            If TextBox3.Text = "Admin" Then
                TextBox1.Enabled = False
                TextBox2.Enabled = False
                TextBox3.Enabled = False
            End If
            If TextBox3.Text = "Pakar" Then
                TextBox1.Enabled = False
                TextBox2.Enabled = False
                TextBox3.Enabled = False
            End If
        End If
        MaximizeBox = False
        TextBox2.UseSystemPasswordChar = True
    End Sub


End Class