Imports System.Data.OleDb
Public Class FormTambahPelanggan1

    Sub KondisiAwal()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_PELANGGAN", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "TBL_PELANGGAN")
        DataGridView1.DataSource = (Ds.Tables("TBL_PELANGGAN"))

        
        Call NomorOtomatis()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        Button1.Text = "INPUT"

        Call DataPelanggan()
    End Sub

    Sub Terbuka()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
    End Sub

    Private Sub FormTambahPelanggan1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Sub DataPelanggan()
        Dim JmlData
        JmlData = DataGridView1.RowCount
        Label7.Text = JmlData - 1
    End Sub

    Sub NomorOtomatis()
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_PELANGGAN Where KodePelanggan in (Select Max(KodePelanggan) From TBL_PELANGGAN)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            UrutanKode = "PLGN_" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            UrutanKode = "PLGN_" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox1.Text = UrutanKode
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            Call Terbuka()
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Button1.Text = "SIMPAN"
            TextBox2.Focus()
        Else
            If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Semua Data Pelanggan wajib diisii!!", MsgBoxStyle.Exclamation, "PERINGATANN BROO !!")
            Else
                Call Koneksi()
                Dim SimpanData As String = "insert into TBL_PELANGGAN values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                cmd = New OleDbCommand(SimpanData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Anda Berhasil di input", MsgBoxStyle.Information, "SUKSES TAMABH PELANGGAN")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Button2.Text = "RUBAH"
            Call Terbuka()
        Else
            If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Data harus terisi semua agar bisa di EDIT", MsgBoxStyle.Exclamation, "PERINGATAN EDITT")
            Else
                Call Koneksi()
                Dim EditData As String = "UPDATE TBL_PELANGGAN Set NamaPelanggan='" & TextBox2.Text & "', AlamatPelanggan='" & TextBox3.Text & "', TelpPelanggan='" & TextBox4.Text & "' Where KodePelanggan='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(EditData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil DI-Edit", MsgBoxStyle.Information, "BERHASIL TER-EDIT!!")
                Button2.Text = "EDIT"
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Or TextBox1.Text = "" Then
            MsgBox("Data harap diisi semua!!!", MsgBoxStyle.Information, "INFOAJE")
        Else
            Call Koneksi()
            If MessageBox.Show("Anda Yakin ingin menghapus data pelanggan yang anda pilih?", "HAPUSDATA", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim HapusData As String = "DELETE From TBL_PELANGGAN Where KodePelanggan='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(HapusData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Pelanggan sudah Terhapus dari DATABASE!")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_PELANGGAN Where KodePelanggan LIKE '%" & TextBox5.Text & "%'", Conn)
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Call Koneksi()
            Da = New OleDbDataAdapter("Select * From TBL_PELANGGAN Where KodePelanggan LIKE '%" & TextBox5.Text & "%'", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "KetemuData")
            DataGridView1.DataSource = (Ds.Tables("KetemuData"))
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index

        On Error Resume Next
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call KondisiAwal()
    End Sub
End Class