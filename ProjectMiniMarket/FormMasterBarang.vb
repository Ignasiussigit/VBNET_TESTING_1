Imports System.Data.OleDb
Public Class FormMasterBarang

    Sub KondisiAwal()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_BARANG", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "TBL_BARANG")
        DataGridView1.DataSource = (Ds.Tables("TBL_BARANG"))
        Label6.Text = ""

        ComboBox1.Items.Clear()
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList


        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("PCS")
        ComboBox2.Items.Add("LUSIN")
        ComboBox2.DropDownStyle = ComboBoxStyle.DropDownList

        Call MunculKategori()
    End Sub

    Sub TextMute()
        TextBox1.Enabled = True
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        ComboBox2.Enabled = False
    End Sub

    Sub TextOpen()
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        ComboBox2.Enabled = True
    End Sub

    Sub DataBersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub FormMasterBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        Call TextMute()
        Call DataBarang()
    End Sub

    Sub MunculKategori()
        cmd = New OleDbCommand("Select * From TBL_KATEGORI", Conn)
        Rd = cmd.ExecuteReader
        Do While Rd.Read
            ComboBox1.Items.Add(Rd.Item(0))
        Loop
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            Call TextOpen()
            Button1.Text = "SIMPAN"
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox2.Text = "" Then
                MsgBox("Data harus diisikan Semua ....")
            Else
                Call Koneksi()
                Dim SimpanData As String = "Insert into TBL_BARANG Values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox2.Text & "')"
                cmd = New OleDbCommand(SimpanData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Ter-Input...")
                Call KondisiAwal()
                Call DataBersih()
                Call TextMute()
                Button1.Text = "INPUT"
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_KATEGORI Where KodeKategori = '" & ComboBox1.Text & "'", Conn)
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Label6.Text = Rd.Item("NamaKategori")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Call TextOpen()
            Button2.Text = "SIMPAN"
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox2.Text = "" Then
                MsgBox("Data harus diisikan Semua ....")
            Else
                Call Koneksi()
                If MessageBox.Show("Anda Yakin Ingin mengedit kolom ini ?", "Kolom Edit", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Dim EditData As String = "UPDATE TBL_BARANG set NamaBarang='" & TextBox2.Text & "', KodeKategori='" & ComboBox1.Text & "', HargaBarang='" & TextBox3.Text & "', JumlahBarang='" & TextBox4.Text & "', SatuanBarang='" & ComboBox2.Text & "' Where KodeBarang='" & TextBox1.Text & "'"
                    cmd = New OleDbCommand(EditData, Conn)
                    cmd.ExecuteNonQuery()
                    MsgBox("Data berhasil Di-Edit... ")
                    Call KondisiAwal()
                    Call DataBersih()
                    Call TextMute()
                    Button2.Text = "EDIT"
                Else
                    MsgBox("Data GAGAL Ter-edit :) ")
                    Call DataBersih()
                    Call TextMute()
                    Button2.Text = "EDIT"
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("KodeBarang WAJIB DI isi")
        Else
            Call Koneksi()
            If MessageBox.Show("Anda yakin ingin HAPUS data ini  ?", "Kolom HAPUS", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                Dim HapusData As String = "DELETE From TBL_BARANG where KodeBarang='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(HapusData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data anda berhasil di HAPUS ...")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_BARANG where KodeBarang='" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaBarang")
                TextBox3.Text = Rd.Item("HargaBarang")
                TextBox4.Text = Rd.Item("JumlahBarang")
                ComboBox2.Text = Rd.Item("SatuanBarang")
                ComboBox1.Text = Rd.Item("KodeKategori")

                'TextBox3.Text = Rd.Item("HargaBarang")
                'TextBox4.Text = Rd.Item(" JumlahBarang")
                'ComboBox2.Text = Rd.Item("SatuaBarang")
            Else
                MsgBox("Data yang anda cari tidak ada...")
                TextBox1.Text = ""
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub DataBarang()
        Dim JmlData
        JmlData = DataGridView1.RowCount
        Label9.Text = JmlData - 1
    End Sub
End Class
