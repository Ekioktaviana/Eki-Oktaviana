Imports System.Data.OleDb
Public Class FormLogin

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        koneksi()
        CMD = New OleDbCommand("Select * FROM TB_Karyawan WHERE Username='" & TextUser.Text & "' and Sandi='" & TextPass.Text & "' ", konek)
        DR = CMD.ExecuteReader

        If TextUser.Text = "" Or TextPass.Text = "" Then
            MsgBox("text tidak boleh kosong")
            TextUser.Focus()
        End If
        DR.Read()
        If DR.HasRows() Then
            If DR("Jabatan") = "Admin" Then
                Formadmin2.Show()
                Me.Hide()
            ElseIf DR("Jabatan") = "Kasir" Then
                Formkasir.Show()
                Me.Hide()
            ElseIf DR("Jabatan") = "Owner" Then
                Formowner.Show()
                Me.Hide()
            End If
            Else
                MsgBox("Gagal Login")
            End If
        Call kosong()

    End Sub

    Private Sub FormLogin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        koneksi()
    End Sub
    Sub kosong()
        TextUser.Text = ""
        TextPass.Text = ""

    End Sub

    Private Sub TextPass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextPass.TextChanged

    End Sub
End Class
