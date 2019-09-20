Imports System.Data.OleDb
Public Class Formiuser

    Public Sub Formiuser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        ComboBox1.Items.Add("Admin")
        ComboBox1.Items.Add("Kasir")
        ComboBox1.Items.Add("Owner")
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Call mati()

    End Sub
    Public Sub tampil()
        koneksi()
        DA = New OleDbDataAdapter("Select * from TB_Karyawan", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_Karyawan")
        DataGridView1.DataSource = (DS.Tables("TB_Karyawan"))

    End Sub

    Sub panggil()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
        ComboBox1.Text = DataGridView1.Item(4, i).Value
    End Sub


    Sub kosong()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Tidak boleh kosong")
        Else
            koneksi()
            Dim cekdata As String
            cekdata = "SELECT * From TB_Karyawan WHERE ID_Karyawan = '" & TextBox1.Text & "'"
            CMD = New OleDbCommand(cekdata, konek)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Dim simpan As String
                simpan = "Insert into TB_Karyawan values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
                CMD = New OleDbCommand(simpan, konek)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil disimpan")
                Call tampil()
                Call kosong()
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button5.Enabled = True
                Call mati()
            Else
                MsgBox("Data sudah ada!!")
            End If
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim hapus As String
        If TextBox1.Text = "" Then
            MsgBox("Silahkan klik data yang hendak dihapus")
        Else
            If MessageBox.Show("Apakah anda yakin akan menghapus data ini?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                koneksi()
                hapus = ("DELETE FROM TB_Karyawan WHERE ID_Karyawan = '" & TextBox1.Text & "'")
                CMD = New OleDbCommand(hapus, konek)
                CMD.ExecuteNonQuery()
                CMD.Dispose()
                MsgBox("Data berhasil dihapus")
                Call kosong()
                Call tampil()
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button5.Enabled = True
                Call mati()


            End If
        End If


    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call panggil()
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button5.Enabled = False
        Call respawn()

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        koneksi()
        Dim edit As String
        edit = "UPDATE TB_Karyawan set Nama = '" & TextBox2.Text & "' ,Username = '" & TextBox3.Text & "' ,Jabatan ='" & ComboBox1.Text & "' ,Sandi ='" & TextBox4.Text & "' WHERE ID_Karyawan = '" & TextBox1.Text & "'"
        CMD = New OleDbCommand(edit, konek)
        CMD.ExecuteNonQuery()
        MsgBox("Data berhasil terubah")

        Call tampil()
        Call kosong()
        Call mati()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button5.Enabled = True

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FormAdmin.Show()
        Me.Hide()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Button1.Enabled = True
        Button5.Enabled = False
        Call respawn()

    End Sub
    Sub mati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        ComboBox1.Enabled = False
    End Sub
    Sub respawn()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        ComboBox1.Enabled = True

    End Sub
End Class