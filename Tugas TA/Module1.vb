Imports System.Data
Imports System.Data.OleDb
Module Module1

    Public Connection_db As OleDbConnection
    Public D_adapter As OleDbDataAdapter
    Public D_set As DataSet
    Public Table As DataTable
    Public Cmd As OleDbCommand
    Public D_reader As OleDbDataReader
    Public Record As New BindingSource
    Public Manager As CurrencyManager
    Public Codd As String
    Public MyState As String
    Public State As String
    Public MyBookMark As Integer

    Public Sub Koneksi()
        Codd = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\Database_Penyakit.accdb"
        Connection_db = New OleDbConnection(Codd)
        If Connection_db.State = ConnectionState.Closed Then
            Connection_db.Open()
        End If
    End Sub
End Module

