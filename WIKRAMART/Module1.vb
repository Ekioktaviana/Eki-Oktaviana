Imports System.Data.OleDb
Module Module1
    Dim Lokasi As String
    Public konek As OleDbConnection
    Public DA As OleDbDataAdapter
    Public DR As OleDbDataReader
    Public DS As DataSet
    Public CMD As OleDbCommand
    Public DT As DataTable


    Sub koneksi()
        Try
            Lokasi = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=db_Wikramart.mdb"
            konek = New OleDbConnection(Lokasi)
            If konek.State = ConnectionState.Closed Then
                konek.Open()

            End If

        Catch ex As Exception
            MsgBox("Koneksi Gagal")
        End Try
    End Sub
End Module
