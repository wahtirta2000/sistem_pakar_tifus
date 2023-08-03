Imports System.Data
Imports System.Data.OleDb
Public Class Form10

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Or ComboBox3.Text = "" Or ComboBox4.Text = "" Then
            MsgBox("Data Masih Belum Lengkap / Isi Terlebih Dahulu Kolom Yang Sudah Disediakan", MsgBoxStyle.Exclamation, "Warning")
            TextBox1.Focus()
            Exit Sub
        Else
        End If

        Select Case State
            Case "Update"

                Call Koneksi()

                Cmd.Connection = Connection_db
                Cmd.CommandType = CommandType.Text
                Cmd.CommandText = "Update Pengetahuan_Pakar set ID_Penyakit='" & ComboBox1.Text & _
                    "',Jenis_Penyakit='" & TextBox6.Text & "',Pertanyaan='" & TextBox2.Text & _
                    "',PertanyaanFakta_Ya='" & TextBox3.Text & "',PertanyaanFakta_Tidak='" & TextBox4.Text & "',Ya='" & ComboBox3.Text & "',Tidak='" & ComboBox4.Text & "' where ID_Pengetahuan='" & TextBox1.Text & "'"
                Cmd.ExecuteNonQuery()
                MsgBox("Salah Satu Data Pengetahuan Pakar Berhasil Diubah, Selanjutnya Tekan Tombol Refresh Untuk Bisa Menampilkan Data Terbaru", MsgBoxStyle.Information, "Notification")

            Case "Tambah"

                Call Koneksi()

                Cmd.Connection = Connection_db
                Cmd.CommandType = CommandType.Text
                Cmd.CommandText = "Insert into Pengetahuan_Pakar(ID_Pengetahuan,ID_Penyakit,Jenis_Penyakit,Pertanyaan,PertanyaanFakta_Ya,PertanyaanFakta_Tidak,Ya,Tidak)values " & "('" & TextBox1.Text & "','" & ComboBox1.Text & "','" & TextBox6.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox3.Text & "','" & ComboBox4.Text & "')"
                Cmd.ExecuteNonQuery()
                MsgBox("Data Pengetahuan Pakar Berhasil Disimpan Dalam Database, Selanjutnya Tekan Tombol Refresh Untuk Bisa Menampilkan Data Terbaru", MsgBoxStyle.Information, "Notification")

        End Select
        TextBox1.Focus()
        Call SetState("Tampil")
        Call Tampil_Data_JenisPenyakit()
        Call TampilSolusi_Penyakit()
        Call Tampil_Data_PengetahuanPakar()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        State = "Update"
        Call SetState("Update")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        State = "Tambah"
        MyBookMark = Manager.Position
        Call SetState("Add")
        Manager.AddNew()
        TextBox1.Focus()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Manager.CancelCurrentEdit()
        If MyState = "Add" Then
            Manager.Position = MyBookMark
        End If
        Call SetState("Tampil")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            Call tampil()

            Cmd.Connection = Connection_db
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = "Delete From Pengetahuan_Pakar where ID_Pengetahuan='" & TextBox1.Text & "'"
            Cmd.ExecuteNonQuery()

            If MessageBox.Show("Apakah Anda Yakin Untuk Menghapus Data Ini ? Jika 'YA' Tekan Tombol OK dan Selanjutnya Tekan Tombol Refresh Untuk Bisa Menampilkan Data Terbaru", "Notification", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                Manager.RemoveAt(Manager.Position)
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

        Call SetState("Tampil")
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Call tampil()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Close()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Manager.Position = 0
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Manager.Position -= 1
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Manager.Position += 1
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Manager.Position = Manager.Count - 1
    End Sub

    Sub tampil()

        Call Koneksi()

        Cmd = New OleDbCommand("SELECT * FROM Pengetahuan_Pakar", Connection_db)
        D_set = New DataSet
        D_adapter.Fill(D_set)
        Record.DataSource = D_set
        Record.DataMember = D_set.Tables(0).ToString()
        DataGridView3.DataSource = Record

    End Sub

    Private Sub SetState(ByVal AppState As String)
        MyState = AppState
        Select Case AppState
            Case "Tampil"

                Button1.Enabled = False
                Button2.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = False
                Button11.Enabled = True
                Button9.Enabled = True
                Button8.Enabled = True
                Button4.Enabled = True
                Button3.Enabled = True

                TextBox1.ReadOnly = True
                ComboBox1.Enabled = False
                TextBox6.ReadOnly = True
                TextBox2.ReadOnly = True
                TextBox3.ReadOnly = True
                TextBox4.ReadOnly = True
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False

            Case "Update", "Add"

                Button1.Enabled = True
                Button2.Enabled = False
                Button5.Enabled = False
                Button6.Enabled = True
                Button11.Enabled = False
                Button9.Enabled = False
                Button8.Enabled = False
                Button4.Enabled = False
                Button3.Enabled = False

                TextBox1.ReadOnly = False
                ComboBox1.Enabled = True
                TextBox6.ReadOnly = False
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = False
                TextBox4.ReadOnly = False
                ComboBox3.Enabled = True
                ComboBox4.Enabled = True

        End Select
    End Sub

    Sub Data_Grid_View3()
        DataGridView3.Columns(0).Width = 100
        DataGridView3.Columns(1).Width = 100
        DataGridView3.Columns(2).Width = 100
        DataGridView3.Columns(3).Width = 400
        DataGridView3.Columns(4).Width = 100
        DataGridView3.Columns(5).Width = 100
        DataGridView3.Columns(6).Width = 50
        DataGridView3.Columns(7).Width = 50

        DataGridView3.Columns(0).HeaderText = "ID Pengetahuan"
        DataGridView3.Columns(1).HeaderText = "ID Penyakit"
        DataGridView3.Columns(2).HeaderText = "Jenis_Penyakit"
        DataGridView3.Columns(3).HeaderText = "Pertanyaan"
        DataGridView3.Columns(4).HeaderText = "Pertanyaan Fakta Ya"
        DataGridView3.Columns(5).HeaderText = "Pertanyaan Fakta Tidak"
        DataGridView3.Columns(6).HeaderText = "Ya"
        DataGridView3.Columns(7).HeaderText = "Tidak"

        DataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.Aqua
        DataGridView3.GridColor = Color.Silver

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim Cari_Data As String = "SELECT * FROM Pengetahuan_Pakar WHERE ID_Pengetahuan='" & TextBox5.Text & "'"
        Try
            Call Koneksi()
            Dim Cmd As New OleDbCommand(Cari_Data, Connection_db)
            Dim D_reader As OleDbDataReader = Cmd.ExecuteReader
            D_reader.Read()
            If D_reader.HasRows Then
                TextBox1.Text = D_reader("ID_Pengetahuan")
                ComboBox1.Text = D_reader("ID_Penyakit")
                TextBox6.Text = D_reader("Jenis_Penyakit")
                TextBox2.Text = D_reader("Pertanyaan")
                TextBox3.Text = D_reader("PertanyaanFakta_Ya")
                TextBox4.Text = D_reader("PertanyaanFakta_Tidak")
                ComboBox3.Text = D_reader("Ya")
                ComboBox4.Text = D_reader("Tidak")
            Else
                MsgBox("Data Dari Pengetahuan Pakar Tidak Dapat Ditemukan Dalam Database", MsgBoxStyle.Exclamation, "Warning")
            End If
            D_reader.Close()
            Cmd.Dispose()
            Connection_db.Close()
            Connection_db.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error - Tidak Dapat Menemukan Data")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Cmd = New OleDbCommand("Select * From Jenis_Penyakit where ID_Penyakit='" & ComboBox1.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If D_reader.HasRows = True Then
            TextBox6.Text = D_reader.Item(1).ToString()
        Else
            MsgBox("Data Jenis Kerusakan, Tidak Terdapat Dalam Database / Tidak Tersedia", MsgBoxStyle.Exclamation, "Warning")
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Cmd = New OleDbCommand("Select * From Pengetahuan_Pakar where ID_Pengetahuan='" & ComboBox3.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If D_reader.HasRows = True Then
            TextBox7.Text = D_reader.Item(1).ToString()
        Else
            Call Tampil_TIDAK()
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Cmd = New OleDbCommand("Select * From Solusi_Permasalahan where ID_Solusi='" & ComboBox4.Text & "'", Connection_db)
        D_reader = Cmd.ExecuteReader
        D_reader.Read()
        If D_reader.HasRows = True Then
            TextBox8.Text = D_reader.Item(1).ToString()
        Else
            Call Tampil_YA()
        End If
    End Sub

    Sub Tampil_TIDAK()
        Cmd = New OleDbCommand("Select ID_Solusi From Solusi_Permasalahan", Connection_db)
        D_reader = Cmd.ExecuteReader
        Do While D_reader.Read
            ComboBox4.Items.Add(D_reader.Item(0))
        Loop
    End Sub

    Sub Tampil_YA()
        Cmd = New OleDbCommand("Select ID_Pengetahuan From Pengetahuan_Pakar", Connection_db)
        D_reader = Cmd.ExecuteReader
        Do While D_reader.Read
            ComboBox3.Items.Add(D_reader.Item(0))
        Loop
    End Sub

    Sub Tampil_Data_JenisPenyakit()
        D_adapter = New OleDbDataAdapter("Select * From Jenis_Penyakit", Connection_db)
        D_set = New DataSet
        D_set.Clear()
        D_adapter.Fill(D_set, "Data_Penyakit")
        DataGridView1.DataSource = D_set.Tables("Data_Penyakit")
        DataGridView1.Refresh()
    End Sub

    Sub TampilSolusi_Penyakit()
        D_adapter = New OleDbDataAdapter("Select * From Solusi_Permasalahan", Connection_db)
        D_set = New DataSet
        D_set.Clear()
        D_adapter.Fill(D_set, "Solusi_Permasalahan")
        DataGridView2.DataSource = D_set.Tables("Solusi_Permasalahan")
        DataGridView2.Refresh()
    End Sub

    Sub Tampil_Data_PengetahuanPakar()
        D_adapter = New OleDbDataAdapter("Select * From Pengetahuan_Pakar", Connection_db)
        D_set = New DataSet
        D_set.Clear()
        D_adapter.Fill(D_set, "Pengetahuan_Pakar")
        DataGridView3.DataSource = D_set.Tables("Pengetahuan_Pakar")
        DataGridView3.Refresh()
    End Sub

    Sub Tampildi_IDPenyakit()
        Cmd = New OleDbCommand("Select ID_Penyakit From Jenis_Penyakit", Connection_db)
        D_reader = Cmd.ExecuteReader
        Do While D_reader.Read
            ComboBox1.Items.Add(D_reader.Item(0))
        Loop
    End Sub

    Sub Data_Grid_View2()
        DataGridView2.Columns(0).Width = 100
        DataGridView2.Columns(1).Width = 1000
        DataGridView2.Columns(2).Width = 1000

        DataGridView2.Columns(0).HeaderText = "ID Solusi"
        DataGridView2.Columns(1).HeaderText = "Solusi Permasalahan"
        DataGridView2.Columns(2).HeaderText = "Mengatasi Solusi Permasalahan"

        DataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Aqua
        DataGridView2.GridColor = Color.Silver
    End Sub

    Sub Data_Grid_View1()
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 1100

        DataGridView1.Columns(0).HeaderText = "ID Penyakit"
        DataGridView1.Columns(1).HeaderText = "Penyakit"

        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Aqua
        DataGridView1.GridColor = Color.Silver
    End Sub

    Private Sub Form10_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()

        Cmd = New OleDbCommand("SELECT * FROM Pengetahuan_Pakar ORDER BY ID_Pengetahuan", Connection_db)
        D_adapter = New OleDbDataAdapter()
        D_adapter.SelectCommand = Cmd
        Table = New DataTable()
        D_adapter.Fill(Table)
        TextBox1.DataBindings.Add("Text", Table, "ID_Pengetahuan")
        ComboBox1.DataBindings.Add("Text", Table, "ID_Penyakit")
        TextBox6.DataBindings.Add("Text", Table, "Jenis_Penyakit")
        TextBox2.DataBindings.Add("Text", Table, "Pertanyaan")
        TextBox3.DataBindings.Add("Text", Table, "PertanyaanFakta_Ya")
        TextBox4.DataBindings.Add("Text", Table, "PertanyaanFakta_Tidak")
        ComboBox3.DataBindings.Add("Text", Table, "Ya")
        ComboBox4.DataBindings.Add("Text", Table, "Tidak")
        DataGridView3.DataSource = Record
        Manager = DirectCast(Me.BindingContext(Table), CurrencyManager)

        Call SetState("Tampil")
        Call tampil()
        Call Data_Grid_View3()

        Call Tampil_Data_PengetahuanPakar()
        TextBox7.Hide()
        TextBox8.Hide()
        MaximizeBox = False
    End Sub

    Private Sub Form_Data_Pengetahuan_Pakar_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Koneksi()
        Call Tampil_Data_JenisPenyakit()
        Call Tampildi_IDPenyakit()
        Call Tampil_Data_PengetahuanPakar()
        Call TampilSolusi_Penyakit()
        Call Tampil_YA()
        Call Tampil_TIDAK()
        Call Data_Grid_View1()
        Call Data_Grid_View2()
        Call Data_Grid_View3()
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class