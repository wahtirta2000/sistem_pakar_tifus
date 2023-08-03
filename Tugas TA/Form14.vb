Imports System.Data
Imports System.Data.OleDb

Public Class Form14

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Cmd = New OleDbCommand("Select * From Jenis_Penyakit where Penyakit='" & ComboBox1.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If D_reader.HasRows = True Then
            TextBox1.Text = D_reader.Item(1)
        Else
            MsgBox("Jenis Penyakit tidak tersedia")
        End If
    End Sub

    Sub Tampil_Penyakit()
        Cmd = New OleDbCommand("Select Penyakit From Jenis_Penyakit", Connection_db)
        D_reader = Cmd.ExecuteReader
        Do While D_reader.Read
            ComboBox1.Items.Add(D_reader.Item(0))
        Loop
    End Sub

    Sub Ya()
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Solusi_Permasalahan where ID_Solusi='" & TextBox2.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
    End Sub

    Sub Tidak()
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Solusi_Permasalahan where ID_Solusi='" & TextBox3.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
    End Sub

    Sub Non_Aktif()
        TextBox1.Hide()
        TextBox2.Hide()
        TextBox3.Hide()
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        RichTextBox1.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub Form14_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call Tampil_Penyakit()
        Call Non_Aktif()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Pengetahuan_Pakar where ID_Pengetahuan='" & TextBox2.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If Not D_reader.HasRows Then
            Call Ya()
            Me.Close()
            Form15.Show()
            TextBox2.Text = D_reader.Item("ID_Solusi")
            If TextBox2.Text = D_reader("ID_Solusi") Then
                Form15.RichTextBox1.Text = D_reader("Solusi_Permasalahan")
                Form15.RichTextBox2.Text = D_reader("Mengatasi_Solusi_Permasalahan")
            End If
        Else
            TextBox2.Text = D_reader.Item("ID_Pengetahuan")
            If TextBox2.Text = D_reader("ID_Pengetahuan") Then
                RichTextBox1.Text = D_reader("Pertanyaan")
                TextBox2.Text = D_reader("Ya")
                TextBox3.Text = D_reader("Tidak")
                ComboBox1.Enabled = False
                Button1.Enabled = True
                Button2.Enabled = True
                Form15.RichTextBox1.Enabled = True
                Form15.RichTextBox2.Enabled = True
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        Cmd = New OleDbCommand("Select * From Pengetahuan_Pakar where ID_Pengetahuan='" & TextBox3.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If Not D_reader.HasRows Then
            Call Tidak()
            Me.Close()
            Form15.Show()
            TextBox3.Text = D_reader.Item("ID_Solusi")
            If TextBox3.Text = D_reader("ID_Solusi") Then
                Form15.RichTextBox1.Text = D_reader("Solusi_Permasalahan")
                Form15.RichTextBox2.Text = D_reader("Mengatasi_Solusi_Permasalahan")
            End If
        Else
            TextBox3.Text = D_reader.Item("ID_Pengetahuan")
            If TextBox3.Text = D_reader("ID_Pengetahuan") Then
                RichTextBox1.Text = D_reader("Pertanyaan")
                TextBox2.Text = D_reader("Ya")
                TextBox3.Text = D_reader("Tidak")
                ComboBox1.Enabled = False
                Button1.Enabled = True
                Button2.Enabled = True
                Form15.RichTextBox1.Enabled = False
                Form15.RichTextBox2.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Cmd = New OleDbCommand("Select * From Pengetahuan_Pakar where Jenis_Penyakit='" & TextBox1.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If D_reader.HasRows Then
            TextBox1.Text = D_reader.Item("Jenis_Penyakit")
            If TextBox1.Text = D_reader("Jenis_Penyakit") Then
                RichTextBox1.Text = D_reader("Pertanyaan")
                TextBox2.Text = D_reader("Ya")
                TextBox3.Text = D_reader("Tidak")
                ComboBox1.Enabled = False
                Button1.Enabled = True
                Button2.Enabled = True
            Else
                MsgBox("Silahkan Pilih Jenis Penyakit Yang Tersedia..!")
            End If
        End If
    End Sub

End Class