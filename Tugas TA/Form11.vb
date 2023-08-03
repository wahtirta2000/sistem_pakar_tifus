Imports System.Data
Imports System.Data.OleDb
Public Class Form11
    
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data ID dan Penyakit Tidak Boleh Dikosongkan", MsgBoxStyle.Exclamation, "Warning")
            TextBox1.Focus()
            Exit Sub
        Else
        End If

        Select Case State
            Case "Update"

                Call Koneksi()

                Cmd.Connection = Connection_db
                Cmd.CommandType = CommandType.Text
                Cmd.CommandText = "Update Jenis_Penyakit Set Penyakit='" & TextBox2.Text & "' where ID_Penyakit='" & TextBox1.Text & "'"
                Cmd.ExecuteNonQuery()
                MsgBox("Salah Satu Data Jenis Penyakit Berhasil Diubah, Selanjutnya Tekan Tombol Refresh Untuk Bisa Menampilkan Data Terbaru", MsgBoxStyle.Information, "Notification")

            Case "Tambah"

                Call Koneksi()

                Cmd.Connection = Connection_db
                Cmd.CommandType = CommandType.Text
                Cmd.CommandText = "Insert into Jenis_Penyakit(ID_Penyakit, Penyakit)" & _
                    "  values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                Cmd.ExecuteNonQuery()
                MsgBox("Data Jenis Penyakit Berhasil Disimpan Dalam Database, Selanjutnya Tekan Tombol Refresh Untuk Bisa Menampilkan Data Terbaru", MsgBoxStyle.Information, "Notification")

        End Select
        TextBox1.Focus()
        Call SetState("Tampil")

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        State = "Update"
        Call SetState("Update")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try

            Call tampil()

            Cmd.Connection = Connection_db
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = "Delete From Jenis_Penyakit where ID_Penyakit='" & TextBox1.Text & "'"
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Manager.Position = Manager.Count - 1
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Manager.Position += 1
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Manager.Position -= 1
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Manager.Position = 0
    End Sub

    Private Sub Form11_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()

        Cmd = New OleDbCommand("SELECT * FROM Jenis_Penyakit ORDER BY ID_Penyakit", Connection_db)
        D_adapter = New OleDbDataAdapter()
        D_adapter.SelectCommand = Cmd
        Table = New DataTable()
        D_adapter.Fill(Table)
        TextBox1.DataBindings.Add("Text", Table, "ID_Penyakit")
        TextBox2.DataBindings.Add("Text", Table, "Penyakit")
        DataGridView1.DataSource = Record
        Manager = DirectCast(Me.BindingContext(Table), CurrencyManager)

        Call SetState("Tampil")
        Call tampil()
        Call Data_Grid_View1()
        MaximizeBox = False
    End Sub

    Sub tampil()

        Call Koneksi()

        Cmd = New OleDbCommand("SELECT * FROM Jenis_Penyakit", Connection_db)
        D_set = New DataSet
        D_adapter.Fill(D_set)
        Record.DataSource = D_set
        Record.DataMember = D_set.Tables(0).ToString()
        DataGridView1.DataSource = Record

    End Sub

    Private Sub SetState(ByVal AppState As String)
        MyState = AppState
        Select Case AppState
            Case "Tampil"

                Button8.Enabled = False
                Button12.Enabled = True
                Button7.Enabled = True
                Button6.Enabled = False
                Button9.Enabled = True
                Button3.Enabled = True
                Button2.Enabled = True
                Button1.Enabled = True
                Button5.Enabled = True

                TextBox1.ReadOnly = True
                TextBox2.ReadOnly = True

            Case "Update"

                Button8.Enabled = True
                Button12.Enabled = False
                Button7.Enabled = False
                Button6.Enabled = True
                Button9.Enabled = False
                Button3.Enabled = False
                Button2.Enabled = False
                Button1.Enabled = False
                Button5.Enabled = False

                TextBox1.ReadOnly = True
                TextBox2.ReadOnly = False

            Case "Add"

                Button8.Enabled = True
                Button12.Enabled = False
                Button7.Enabled = False
                Button6.Enabled = True
                Button9.Enabled = False
                Button3.Enabled = False
                Button2.Enabled = False
                Button1.Enabled = False
                Button5.Enabled = False

                TextBox1.ReadOnly = False
                TextBox2.ReadOnly = False

        End Select
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim Cari_Data As String = "SELECT * FROM Jenis_Penyakit WHERE Penyakit='" & TextBox3.Text & "'"
        Try
            Call Koneksi()
            Dim Cmd As New OleDbCommand(Cari_Data, Connection_db)
            Dim D_reader As OleDbDataReader = Cmd.ExecuteReader
            D_reader.Read()
            If D_reader.HasRows Then
                TextBox1.Text = D_reader("ID_Penyakit")
                TextBox2.Text = D_reader("Penyakit")
            Else
                MsgBox("Data Dari Jenis Penyakit Tidak Dapat Ditemukan Dalam Database", MsgBoxStyle.Exclamation, "Warning")
            End If
            D_reader.Close()
            Cmd.Dispose()
            Connection_db.Close()
            Connection_db.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error - Tidak Dapat Menemukan Data")
        End Try
    End Sub

    Sub Data_Grid_View1()
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 1100

        DataGridView1.Columns(0).HeaderText = "ID Penyakit"
        DataGridView1.Columns(1).HeaderText = "Penyakit"

        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Aqua
        DataGridView1.GridColor = Color.Silver
    End Sub

End Class