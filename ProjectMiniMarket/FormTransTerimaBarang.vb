Imports System.Data.OleDb
Public Class FormTransTerimaBarang

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label16.Text = TimeOfDay
    End Sub
    Sub KondisiAwal()
        Label7.Text = ""
        Label8.Text = ""
        Label9.Text = ""
        Label15.Text = Today
        Label17.Text = FormMenuUtama.USER2.Text
        TextBox2.Enabled = False
        Call NomorOtomatis()
        Call BuatKolom()
    End Sub


    Private Sub FormTransTerimaBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call MunculKodeSupplier()
        Call KondisiAwal()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    'kode dibawah untuk mengambil table database supplier di dalam combobox1
    Sub MunculKodeSupplier()
        Call Koneksi()
        ComboBox1.Items.Clear()
        cmd = New OleDbCommand("Select * From TBL_SUPPLIER", Conn)
        Rd = cmd.ExecuteReader
        Do While Rd.Read
            ComboBox1.Items.Add(Rd.Item(0))
        Loop
    End Sub

    'kode dibawah ini untuk mengambil data sesuai dipilih yang berada di combobox1
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_SUPPLIER where KodeSupplier = '" & ComboBox1.Text & "'", Conn)
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Label7.Text = Rd!NamaSupplier
            Label8.Text = Rd!AlamatSupplier
            Label9.Text = Rd!TelpSupplier
        End If
    End Sub

    'CODE-NYA DI SESUAIKAM
    'Menjeneret kode secara otomatis Format nya seperti ini "T tahun,Bulan,Tanggal" dan data nya bertambah 1 setian inputan baru
    Sub NomorOtomatis()
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_TERIMA Where NoTerima in (Select max(NoTerima) From TBL_TERIMA)", Conn)
        Dim UrutanKode As String
        Dim Hiung As Long
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            UrutanKode = "T" + Format(Now, "yyMMdd") + "001"
        Else
            Hiung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            UrutanKode = "T" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hiung, 3)
        End If
        LBLNoTerima.Text = UrutanKode
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Kode", "Kode")
        DataGridView1.Columns.Add("Nama", "NamaBarang")
        DataGridView1.Columns.Add("Harga", "Harga")
        DataGridView1.Columns.Add("Jumlah", "Jumlah")
        DataGridView1.Columns.Add("Subtotal", "Subtotal")
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_BARANG Where KodeBarang='" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                MsgBox("Kode Barang Tidak ditemukan", MsgBoxStyle.Exclamation, "Peringatan")
            Else
                Label20.Text = Rd.Item("NamaBarang")
                Label22.Text = Rd.Item("HargaBarang")
                TextBox2.Enabled = True
            End If
        End If
    End Sub

    'Berfungsi ketika user klik Button "Insert" makan akan mucul di datagridview yang sudah kita bikin dengan class "BuatKolom"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label20.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Silahkan Masukan KodeBarang dan Tekan ENTER!", MsgBoxStyle.Exclamation, "Peringatan")
        Else
            DataGridView1.Rows.Add(New String() {TextBox1.Text, Label20.Text, Label22.Text, TextBox2.Text, Val(Label22.Text) * Val(TextBox2.Text)})
            Call RumusSubtotal()
            TextBox1.Text = ""
            Label20.Text = ""
            Label22.Text = ""
            TextBox2.Text = ""
            TextBox2.Enabled = False
            Call RumusCariItem()
        End If
    End Sub
    'menghitung total jumlah harga dari jumlah item yang user beli
    Sub RumusSubtotal()
        Dim hitung As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitung = hitung + DataGridView1.Rows(i).Cells(4).Value
            Label11.Text = hitung
        Next
    End Sub
    'menghitung tottal item user beli
    Sub RumusCariItem()
        Dim HitungItem As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            HitungItem = HitungItem + DataGridView1.Rows(i).Cells(3).Value
            Label25.Text = HitungItem
        Next
    End Sub
    'Fungsi TOMBOL Simpan
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Label7.Text = "" Or Label11.Text = "" Then
            MsgBox("Di Mohon untuk MemilihKodeSupplier dan KodeBarang", MsgBoxStyle.Exclamation, "Peringatan!")
        Else
            Call Koneksi()
            Dim SimpanTerima As String = "Insert into TBL_TERIMA Values ('" & LBLNoTerima.Text & "', '" & Label15.Text & "', '" & Label16.Text & "', '" & Label25.Text & "','" & Label11.Text & "','" & ComboBox1.Text & "','" & FormMenuUtama.USER2.Text & "')"
            cmd = New OleDbCommand(SimpanTerima, Conn)
            cmd.ExecuteNonQuery()

            'Code Dibawah ini untuk insert ke TBL_DETAILTERIMA
            For Baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Dim SimpanDetail As String = "Insert into TBL_DETAILTERIMA Values('" & LBLNoTerima.Text & "','" & DataGridView1.Rows(Baris).Cells(0).Value & "', '" & DataGridView1.Rows(Baris).Cells(1).Value & "', '" & DataGridView1.Rows(Baris).Cells(2).Value & "', '" & DataGridView1.Rows(Baris).Cells(3).Value & "', '" & DataGridView1.Rows(Baris).Cells(4).Value & "')"
                cmd = New OleDbCommand(SimpanDetail, Conn)
                cmd.ExecuteNonQuery()

                'cmd = New OleDbCommand("Select * From TBL_BARANG Where KodeBarang='" & DataGridView1.Rows(Baris).Cells(0).Value & "'", Conn)
                'Rd = cmd.ExecuteReader
                'Rd.Read()
                'Dim KurangiStock As String = "Update TBL_BARANG Set JumlahBarang = '" & Rd.Item("JumlahBarang") + DataGridView1.Rows(Baris).Cells(3).Value & "' Where KodeBarang = '" & DataGridView1.Rows(Baris).Cells(0).Value & "'"
                'cmd = New OleDbCommand(KurangiStock, Conn)
                'cmd.ExecuteNonQuery()

                ' Ambil jumlah stok lama dari TBL_BARANG
                Dim StokLama As Integer = 0
                cmd = New OleDbCommand("Select JumlahBarang From TBL_BARANG Where KodeBarang='" & DataGridView1.Rows(Baris).Cells(0).Value & "'", Conn)
                Rd = cmd.ExecuteReader()
                If Rd.HasRows Then
                    Rd.Read()
                    StokLama = CInt(Rd("JumlahBarang")) ' Pastikan konversi ke Integer
                End If
                Rd.Close()

                ' Hitung stok baru
                Dim StokBaru As Integer = StokLama + CInt(DataGridView1.Rows(Baris).Cells(3).Value)

                ' Update stok barang di TBL_BARANG
                Dim UpdateStok As String = "Update TBL_BARANG Set JumlahBarang = " & StokBaru & " Where KodeBarang = '" & DataGridView1.Rows(Baris).Cells(0).Value & "'"
                cmd = New OleDbCommand(UpdateStok, Conn)
                cmd.ExecuteNonQuery()
            Next
            Call KondisiAwal()
            MsgBox("Transaksi Telah Berhasil Disimpan !", MsgBoxStyle.Information, "PERHATIAN!")
        End If
    End Sub

End Class