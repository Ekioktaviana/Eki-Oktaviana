Imports System.Data.OleDb
Public Class Formbarang
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FormAdmin.Show()
        Me.Hide()
    End Sub

    Private Sub FormBARANG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        tampilkombo()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Call mati()

    End Sub
    Sub tampilkombo()
        Dim kombo As String
        kombo = "select * from TB_JenisBarang"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        Do While DR.Read
            ComboBox1.Items.Add(DR.Item(1))
            ComboBox1.SelectedItem = TextBox2.Text
        Loop
    End Sub

    Public Sub tampil()
        koneksi()
        DA = New OleDbDataAdapter("select * from TB_Barang", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_Barang")
        DataGridView1.DataSource = (DS.Tables("TB_Barang"))
    End Sub
    Sub kosong()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        ComboBox1.Text = ""

    End Sub
    Sub panggil()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
        TextBox5.Text = DataGridView1.Item(4, i).Value
        TextBox7.Text = DataGridView1.Item(5, i).Value
        TextBox8.Text = DataGridView1.Item(6, i).Value
        TextBox9.Text = DataGridView1.Item(7, i).Value

        Dim kombo As String
        kombo = "select * from TB_JenisBarang WHERE ID_JenisBarang ='" & TextBox2.Text & "'"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        Do While DR.Read()

            ComboBox1.Items.Add(DR.Item(1))
            ComboBox1.SelectedItem = DR.Item(1)
        Loop
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
            MsgBox("data tidak boleh kosong")
        Else

            koneksi()
            Dim cekdata As String
            cekdata = "select * from TB_Barang WHERE ID_Barang = '" & TextBox1.Text & "'"
            CMD = New OleDbCommand(cekdata, konek)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Dim simpan As String
                simpan = "Insert into TB_Barang values('" & TextBox1.Text & "', '" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "')"
                CMD = New OleDbCommand(simpan, konek)
                CMD.ExecuteNonQuery()
                CMD.Dispose()
                MsgBox("Data berhasil disimpan")
                Call tampil()
                Call kosong()
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button7.Enabled = True
                Call mati()

            Else
                MsgBox("Data sudah ada coyy !")
                Call tampil()
            End If
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim hapus As String
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Silahkan klik data yang hendak dihapus")
        Else
            If MessageBox.Show("Apakah anda yakin akan menghapus data ini?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                koneksi()
                hapus = ("DELETE FROM TB_Barang WHERE ID_Barang = '" & TextBox1.Text & "'")
                CMD = New OleDbCommand(hapus, konek)
                CMD.ExecuteNonQuery()
                CMD.Dispose()
                MsgBox("Data berhasil dihapus")
                Call kosong()
                Call tampil()
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button7.Enabled = True
                Call mati()


            End If
        End If

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        koneksi()
        Dim edit As String
        edit = "UPDATE TB_Barang set ID_Barang = '" & TextBox1.Text & "', ID_JenisBarang = '" & TextBox2.Text & "' , Nama_Barang = '" & TextBox3.Text & "', Harga_Barang = '" & TextBox4.Text & "', Stock = '" & TextBox5.Text & "', Nomor Mesin = '" & TextBox7.Text & "', Warna = '" & TextBox8.Text & "', Kapasitas Mesin = '" & TextBox9.Text & "' WHERE ID_Barang ='" & TextBox1.Text & "'"
        CMD = New OleDbCommand(edit, konek)
        CMD.ExecuteNonQuery()
        MsgBox("data berhasil diperbarui !!")
        MsgBox("(Tidak berlaku jika edit ID Barang)")
        Call tampil()
        Call kosong()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button7.Enabled = True
        Call mati()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call panggil()
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button7.Enabled = False
        Call respawn()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        koneksi()
        Dim cari As String
        cari = "select  * from TB_Barang WHERE Nama_Barang like '%" & TextBox6.Text & "%'"
        DA = New OleDbDataAdapter(cari, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_Barang")
        DataGridView1.DataSource = (DS.Tables("TB_Barang"))
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox6.Text = ""
        Call tampil()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        CMD = New OleDbCommand("select * from TB_JenisBarang where JenisBarang= '" & ComboBox1.Text & "'", konek)
        DR = CMD.ExecuteReader
        DR.Read()
        If DR.HasRows Then
            TextBox2.Text = DR.Item(0)
        End If
    End Sub


    Private Sub TextBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.Click
        TextBox6.Text = ""
    End Sub
    Sub mati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False


    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Button1.Enabled = True
        Button7.Enabled = False
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        Call respawn()


    End Sub
    Sub respawn()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        ComboBox1.Enabled = True

    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Formowner.Show()
        Me.Hide()

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
End Class
