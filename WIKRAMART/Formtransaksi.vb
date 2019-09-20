Imports System.Data.OleDb
Public Class Formtransaksi
    Private Sub Formtransaksi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        tampil2()
        tampilkombo2()
        tampilkombo3()
        TextBox8.Visible = False

        Call mati()
        Call kunci()
        Call mati2()
        Call kunci2()

    End Sub


    Public Sub tampil()
        koneksi()
        DA = New OleDbDataAdapter("Select * from TB_DetailTransaksi", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_DetailTransaksi")
        DataGridView1.DataSource = (DS.Tables("TB_DetailTransaksi"))
    End Sub
    Sub Kosong()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""


    End Sub
    Sub panggil()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
        TextBox5.Text = DataGridView1.Item(4, i).Value
        TextBox6.Text = DataGridView1.Item(5, i).Value

        Dim kombo As String
        kombo = "select * from TB_JenisBarang WHERE ID_JenisBarang ='" & TextBox2.Text & "'"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        Do While DR.Read()

            ComboBox3.Items.Add(DR.Item(3))
            ComboBox3.SelectedItem = DR.Item(3)
        Loop
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Data tidak boleh kosong")
        Else
            koneksi()
            Dim cekdata As String
            cekdata = " SELECT * FROM TB_DetailTransaksi WHERE ID_DetailTransaksi = '" & TextBox1.Text & "'"
            CMD = New OleDbCommand(cekdata, konek)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Dim simpan As String
                simpan = "insert into TB_DetailTransaksi values('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & TextBox6.Text & "')"
                CMD = New OleDbCommand(simpan, konek)
                CMD.ExecuteReader()
                MsgBox("Data Berhasil Disimpan")
                Call editstock()
                Call tampil()
                Call Kosong()

            Else
                MsgBox("Data Sudah Ada!")


            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call panggil()
        Button8.Enabled = False
        Call hidup()
        Call kuncibuka()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim hapus As String
        If TextBox1.Text = "" Then
            MsgBox("Silahkan klik data yang hendak dihapus")
        Else
            If MessageBox.Show("apakah anda yakin menghapus data ini?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                koneksi()
                hapus = ("DELETE FROM TB_DetailTransaksi WHERE ID_DetailTransaksi = '" & TextBox1.Text & "'")
                CMD = New OleDbCommand(hapus, konek)
                CMD.ExecuteNonQuery()
                CMD.Dispose()
                MsgBox("Data berhasil dihapus")
                Call Kosong()
                Call tampil()
            End If

        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        koneksi()
        Dim edit As String
        edit = "UPDATE TB_DetailTransaksi SET ID_DetailTransaksi = '" & TextBox1.Text & "', ID_Transaksi = '" & TextBox2.Text & "', ID_Barang = '" & TextBox3.Text & "', Harga = '" & TextBox4.Text & "', Jumlah = '" & TextBox5.Text & "', Total_Harga = '" & TextBox6.Text & "'  WHERE ID_DetailTransaksi = '" & TextBox1.Text & "' "
        CMD = New OleDbCommand(edit, konek)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil Diperbaharui")
        Call tampil()
        Call panggil()
    End Sub
    'FORM Transaksi'
    Public Sub tampilkombo2()
        Dim kombo As String
        kombo = "SELECT * FROM TB_Karyawan"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        Do While DR.Read
            ComboBox2.Items.Add(DR.Item(1))
        Loop
    End Sub
    Public Sub tampil2()
        koneksi()
        DA = New OleDbDataAdapter("Select * from TB_Transaksi", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "TB_Transaksi")
        DataGridView2.DataSource = (DS.Tables("TB_Transaksi"))
        Call tampilkombo2()

    End Sub
    Sub Kosong2()
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox9.Text = ""



    End Sub
    Sub panggil2()
        Dim i As Integer
        i = DataGridView2.CurrentRow.Index
        TextBox7.Text = DataGridView2.Item(0, i).Value
        TextBox8.Text = DataGridView2.Item(3, i).Value
        TextBox9.Text = DataGridView2.Item(2, i).Value
        TextBox10.Text = DataGridView2.Item(1, i).Value
        Dim kombo As String
        kombo = "SELECT * FROM TB_Karyawan WHERE ID_Karyawan  = '" & TextBox8.Text & "'"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        DR.Read()
        If DR.HasRows Then
            ComboBox2.SelectedItem = DR.Item(1)
        End If


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Then
            MsgBox("Data tidak boleh kosong")
        Else
            koneksi()
            Dim cekdata2 As String
            cekdata2 = " SELECT * FROM TB_Transaksi WHERE ID_Transaksi = '" & TextBox7.Text & "'"
            CMD = New OleDbCommand(cekdata2, konek)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Dim simpan2 As String
                simpan2 = "insert into TB_Transaksi values('" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & TextBox9.Text & "', '" & TextBox10.Text & "')"
                CMD = New OleDbCommand(simpan2, konek)
                CMD.ExecuteReader()
                MsgBox("Data Berhasil Disimpan")
                Call tampil2()
                Call Kosong2()
            Else
                MsgBox("Data Sudah Ada!")


            End If
        End If
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Call panggil2()
        TextBox2.Text = TextBox7.Text
        TextBox11.Text = TextBox10.Text

        Button9.Enabled = False
        Call hidup2()
        Call kuncibuka221()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim hapus2 As String
        If TextBox7.Text = "" Then
            MsgBox("Silahkan klik data yang hendak dihapus")
        Else
            If MessageBox.Show("apakah anda yakin menghapus data ini?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                koneksi()
                hapus2 = ("DELETE FROM TB_Transaksi WHERE ID_Transaksi = '" & TextBox7.Text & "'")
                CMD = New OleDbCommand(hapus2, konek)
                CMD.ExecuteNonQuery()
                CMD.Dispose()
                MsgBox("Data berhasil dihapus")
                Call Kosong2()
                Call tampil2()
            End If

        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        koneksi()
        Dim edit2 As String
        edit2 = "UPDATE TB_Transaksi SET ID_Transaksi = '" & TextBox7.Text & "', ID_Karyawan = '" & TextBox8.Text & "', Tanggal = '" & TextBox9.Text & "'"
        CMD = New OleDbCommand(edit2, konek)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil Diperbaharui")
        Call tampil2()
        Call panggil2()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim kombo As String
        kombo = "SELECT * FROM TB_Karyawan WHERE Nama = '" & ComboBox2.Text & "'"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        DR.Read()
        If DR.HasRows Then
            TextBox8.Text = DR.Item(0)
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Hide()
        Formkasir.Show()
    End Sub

    Sub editstock()
        Dim setok As String
        Dim sisa As Integer
        sisa = Val(TextBox13.Text) - Val(TextBox5.Text)
        setok = "UPDATE TB_Barang SET Stock = '" & sisa & "' WHERE Nama_Barang = '" & ComboBox3.Text & "'  "
        CMD = New OleDbCommand(setok, konek)
        CMD.ExecuteNonQuery()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim kombo As String
        kombo = "SELECT * FROM TB_Barang WHERE Nama_Barang = '" & ComboBox3.Text & "'"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        DR.Read()
        If DR.HasRows Then
            TextBox13.Text = DR.Item(4)
            TextBox4.Text = DR.Item(3)
            TextBox3.Text = DR.Item(0)
        End If
    End Sub

    Public Sub tampilkombo3()
        Dim kombo As String
        kombo = "SELECT * FROM TB_Barang"
        CMD = New OleDbCommand(kombo, konek)
        DR = CMD.ExecuteReader()
        Do While DR.Read
            ComboBox3.Items.Add(DR.Item(2))
        Loop
    End Sub
    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        TextBox6.Text = Val(TextBox4.Text) * Val(TextBox5.Text)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Button1.Enabled = True
        Button8.Enabled = False
        Call kuncibuka2()

    End Sub

    Sub hidup()
        Button2.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = False

    End Sub

    Sub mati()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button3.Enabled = False
        Button8.Enabled = True
    End Sub

    Sub kunci()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        ComboBox3.Enabled = False
    End Sub

    Sub kuncibuka()
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        ComboBox3.Enabled = True
    End Sub

    Sub kuncibuka2()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        ComboBox3.Enabled = True
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Button9.Enabled = False
        Button4.Enabled = True
        Call kuncibuka221()
    End Sub
    Sub hidup2()
        Button4.Enabled = False
        Button5.Enabled = True
        Button6.Enabled = True

    End Sub

    Sub mati2()
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button9.Enabled = True
    End Sub

    Sub kunci2()
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        ComboBox2.Enabled = False
    End Sub

    Sub kuncibuka22()
        TextBox7.Enabled = False
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox10.Enabled = True
        ComboBox2.Enabled = True
    End Sub

    Sub kuncibuka221()
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox10.Enabled = True
        ComboBox2.Enabled = True
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Formstruk.Show()
        Me.Hide()

    End Sub
End Class