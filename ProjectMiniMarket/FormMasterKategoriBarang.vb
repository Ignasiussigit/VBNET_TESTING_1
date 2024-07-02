Imports System.Data.OleDb
Public Class FormMasterKategoriBarang

    Sub KondisiAwal()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_KATEGORI", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "TBL_KATEGORI")
        DataGridView1.DataSource = (Ds.Tables("TBL_KATEGORI"))
    End Sub

    Private Sub FormMasterKategoriBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Harap DATA diisi semua ...", "INFO")
        Else
            Call Koneksi()
            Dim SimpanData As String = "Insert into TBL_KATEGORI values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')"
            cmd = New OleDbCommand(SimpanData, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Ter-Input ...")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Harap DATA diisi semua ...", "INFO")
        Else
            Call Koneksi()
            Dim EditData As String = "UPDATE TBL_KATEGORI Set NamaKategori='" & TextBox2.Text & "', KeteranganKategori='" & TextBox3.Text & "' where KodeKategori='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(EditData, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Ter-Edit ...")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Harap DATA diisi semua ...", "INFO")
        Else
            Call Koneksi()
            Dim HapusData As String = "DELETE From TBL_KATEGORI where KodeKategori='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(HapusData, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Ter-Hapus ...")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_KATEGORI where KodeKategori='" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaKategori")
                TextBox3.Text = Rd.Item("KeteranganKategori")
            Else
                MsgBox("Data Tidak Ada ...")
            End If
        End If



    End Sub
End Class