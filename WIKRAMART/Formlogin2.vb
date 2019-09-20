Imports System.Data.OleDb
Public Class FormLogin2

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
            If DR("Jabatan") = "Owner" Then
                Formowner.Show()
                Me.Hide()
            End If
        Else
            MsgBox("data yang anda masukan salah")
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MessageBox.Show("Apakah anda ingin Kembali?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Formtest.Show()
            Me.Hide()
        End If
    End Sub
End Class
