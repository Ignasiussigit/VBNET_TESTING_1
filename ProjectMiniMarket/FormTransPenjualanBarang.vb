Imports System.Data.OleDb
Public Class FormTransPenjualanBarang

    Sub KondisiAwal()
        Label17.Text = FormMenuUtama.USER2.Text
        Call NomorOtomatis()
        Call MunculPelanggan()
        TextBox2.Enabled = False
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList

        Call BuatKolom()
    End Sub

    Private Sub FormTransPenjualanBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        Label15.Text = Today
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label16.Text = TimeOfDay
    End Sub

    Sub NomorOtomatis()
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_JUAL Where NoJual in (Select Max(NoJual) From TBL_JUAL)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            UrutanKode = "JU-" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            UrutanKode = "JU-" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        LBLNoJual.Text = UrutanKode
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FormTambahPelanggan1.ShowDialog()
    End Sub
    '=================================================================
    'Code dibawah ini untuk memunculkan data pelanggan pada combobox1
    '=================================================================
    Sub MunculPelanggan()
        Call Koneksi()
        ComboBox1.Items.Clear()
        cmd = New OleDbCommand("Select * From TBL_PELANGGAN", Conn)
        Rd = cmd.ExecuteReader
        Do While Rd.Read
            ComboBox1.Items.Add(Rd.Item(0))
        Loop
    End Sub
    '==============================================================================
    'kode dibawah ini untuk mengambil data sesuai dipilih yang berada di combobox1
    '==============================================================================
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_PELANGGAN Where KodePelanggan = '" & ComboBox1.Text & "'", Conn)
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Label7.Text = Rd!NamaPelanggan
            Label8.Text = Rd!AlamatPelanggan
            Label9.Text = Rd!TelpPelanggan
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call KondisiAwal()
    End Sub
    '========================================================================================
    'Kondisi dimana ketika input kode barang lalu di ENTER akan muncul nama,harga barang-nya
    '========================================================================================
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_BARANG Where KodeBarang = '" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                MsgBox("Data tidak ada dalam Database !", MsgBoxStyle.Exclamation, "PERHATIAN")
            Else
                Label20.Text = Rd.Item("NamaBarang")
                Label22.Text = Rd.Item("HargaBarang")
                TextBox2.Enabled = True
            End If
        End If
    End Sub
    '===================================
    'Untuk Buat Kolom pada DataGridView 
    '===================================
    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Kode", "Kode")
        DataGridView1.Columns.Add("Nama", "NamaBarang")
        DataGridView1.Columns.Add("Harga", "Harga")
        DataGridView1.Columns.Add("Jumlah", "Jumlah")
        DataGridView1.Columns.Add("Subtotal", "Subtotal")
    End Sub
    '==============================================================
    'menghitung total jumlah harga dari jumlah item yang user beli
    '==============================================================
    Sub RumusSubtotal()
        Dim Hitung As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            Hitung = Hitung + DataGridView1.Rows(i).Cells(4).Value
            Label11.Text = Hitung
        Next

    End Sub
    '=================================
    'menghitung tottal item user beli
    '=================================
    Sub RumusCariItem()
        Dim HitungItem As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            HitungItem = HitungItem + DataGridView1.Rows(i).Cells(3).Value
            Label25.Text = HitungItem
        Next
    End Sub
    '=======================================================================
    'Berfungsi, ketika user klik "INSERT" akan muncul ke dalam dataGridView
    '=======================================================================
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label20.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Silahkan masukan KodeSupplier dan Tekan ENTER", MsgBoxStyle.Exclamation, "PERINGATAN")
        Else
            Call Koneksi()
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
    '=========================================================================================
    'Kondisi dimana meng-inputkan jumlah uang pelanggan ketika di enter akan muncul kembalian
    '=========================================================================================
    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox3.Text) < Val(Label11.Text) Then
                MsgBox("Pembayaran Kurang!", MsgBoxStyle.Information, "PERINGATAN KEMBALIAN")
            ElseIf Val(TextBox3.Text) = Val(Label11.Text) Then
                Label29.Text = 0
            ElseIf Val(TextBox3.Text) > Val(Label11.Text) Then
                Label29.Text = Val(TextBox3.Text) - Val(Label11.Text)
                Button2.Focus()
            End If

        End If
    End Sub
    '====================================================================================================================
    'Kondisi dimana TOMBOL SIMPAN akan menyimpan kedalam TBL_JUA,TBL_DETAILJUAL dan Upadate JumlahBarang pada TBL_BARANG
    '====================================================================================================================
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Label7.Text = "" Or Label11.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Di Mohon untuk Memilih KodePelangan, KodeBarang dan mengisikan Transaksi !", MsgBoxStyle.Exclamation, "Peringatan!")
        Else
            Call Koneksi()
            Dim SimpanJual As String = "Insert into TBL_JUAL Values ('" & LBLNoJual.Text & "', '" & Label15.Text & "', '" & Label16.Text & "', '" & Label25.Text & "','" & Label11.Text & "','" & TextBox3.Text & "','" & Label29.Text & "','" & ComboBox1.Text & "','" & FormMenuUtama.USER2.Text & "')"
            cmd = New OleDbCommand(SimpanJual, Conn)
            cmd.ExecuteNonQuery()

            'Code Dibawah ini untuk insert ke TBL_DETAILJUAL
            For Baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Dim SimpanDetail As String = "Insert into TBL_DETAILJUAL Values('" & LBLNoJual.Text & "','" & DataGridView1.Rows(Baris).Cells(0).Value & "', '" & DataGridView1.Rows(Baris).Cells(1).Value & "', '" & DataGridView1.Rows(Baris).Cells(2).Value & "', '" & DataGridView1.Rows(Baris).Cells(3).Value & "', '" & DataGridView1.Rows(Baris).Cells(4).Value & "')"
                cmd = New OleDbCommand(SimpanDetail, Conn)
                cmd.ExecuteNonQuery()


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
                Dim StokBaru As Integer = StokLama - CInt(DataGridView1.Rows(Baris).Cells(3).Value)

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