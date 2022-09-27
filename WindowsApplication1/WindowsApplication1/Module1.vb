Imports System.Data.OleDb
Module Module1
    Public Conn As OleDbConnection
    Public da As OleDbDataAdapter
    Public ds As DataSet
    Public cmd As OleDbCommand
    Public rd As OleDbDataReader
    Public str As String

    Public Sub koneksi()
        str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\databasenyajurnal.mdb"
        Conn = New OleDbConnection(str)
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
    End Sub
End Module
