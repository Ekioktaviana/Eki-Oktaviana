Imports System.Data.OleDb
Public Class Formijenisbarang

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Call panggil()
    End Sub

    Public Sub Formijenisbarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        Call mati()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button6.Enabled = True
    End Sub
    Sub panggil()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
    End Sub
    Public Sub tampil()
        koneksi()
        DA = New OleDbDataAdapter("Select * from TB_JenisBarang", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_JenisBarang")
        DataGridView1.DataSource = (DS.Tables("TB_JenisBarang"))
    End Sub
    Sub kosong()
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Tidak boleh kosong")
        Else
            koneksi()
            Dim cekdata As String
            cekdata = "SELECT * From TB_JenisBarang WHERE JenisBarang = '" & TextBox1.Text & "'"
            CMD = New OleDbCommand(cekdata, konek)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Dim simpan As String
                simpan = "Insert into TB_JenisBarang values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                CMD = New OleDbCommand(simpan, konek)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil disimpan")
                Call tampil()
                Call kosong()
                Call mati()
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button6.Enabled = True
            Else
                MsgBox("Data sudah ada!!")
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call panggil()
        Call respawn()
        Button2.Enabled = True
        Button3.Enabled = True
        Button6.Enabled = False

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim hapus As String
        If TextBox1.Text = "" Then
            MsgBox("Silahkan klik data yang hendak dihapus")
        Else
            If MessageBox.Show("Apakah anda yakin akan menghapus data ini?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                koneksi()
                hapus = ("DELETE FROM TB_JenisBarang WHERE ID_JenisBarang = '" & TextBox1.Text & "'")
                CMD = New OleDbCommand(hapus, konek)
                CMD.ExecuteNonQuery()
                CMD.Dispose()
                MsgBox("Data berhasil dihapus")
                Call kosong()
                Call tampil()
                Call mati()
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button6.Enabled = True


            End If
        End If

    End Sub

    Sub edit()
        Try
            koneksi()
            Dim edit As String
            edit = "UPDATE TB_JenisBarang set Jenis_Barang = '" & TextBox2.Text & "' WHERE ID_JenisBarang = '" & TextBox1.Text & "'"
            CMD = New OleDbCommand(edit, konek)
            CMD.ExecuteNonQuery()
            MsgBox("Data berhasil terubah")
        Catch ex As Exception
            MsgBox("Gagal")
        End Try
        Call tampil()
        Call kosong()

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call edit()
        Call mati()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button6.Enabled = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        koneksi()
        Dim cari As String
        cari = "SELECT * FROM TB_JenisBarang WHERE Jenis_Barang like '%" & TextBox3.Text & "%'"
        DA = New OleDbDataAdapter(cari, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_JenisBarang")
        DataGridView1.DataSource = (DS.Tables("TB_JenisBarang"))
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FormLogin.Show()
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call respawn()
        Button1.Enabled = True
        Button6.Enabled = False
    End Sub
    Sub mati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
    End Sub
    Sub respawn()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Formowner.Show()
        Me.Hide()

    End Sub
End Class