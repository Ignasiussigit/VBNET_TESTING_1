'Imports System.Data.OleDb
Public Class FormMenuUtama

    Sub Terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogOutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Visible = False
        TransaksiToolStripMenuItem.Visible = False
        LaporanToolStripMenuItem.Visible = False
    End Sub

    'Ketika belum login maka untuk strip Master , Transaksi dan laporan akan hilang atau tidak muncul
    'Sub ToolHilang()
    '    MasterToolStripMenuItem.Visible = False
    'End Sub
    'UNTUK SUB TERBUKA-NYA ADA DI FORM LOGIN


    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        Dim TabLogin As New FormLogin()

        TabLogin.Show()
    End Sub

    Private Sub FormMenuUtama_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call ToolHilang()
        Call Terkunci()
        TANGGAL2.Text = Today
    End Sub

    Private Sub LogOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogOutToolStripMenuItem.Click
        If MessageBox.Show("Anda Yakin Ingin LogOut? ", "LogOut", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            MsgBox("Berhasil LogOut")
            Call Terkunci()
            FormLogin.Show()
            FormLogin.TextBox1.Focus()
        Else
            MsgBox("Data Masih Aman")
        End If
      
    End Sub

 
    Private Sub AdminToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AdminToolStripMenuItem.Click
        Dim Master As New FormMasterAdmin()

        Master.ShowDialog()
    End Sub

    Private Sub SupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierToolStripMenuItem.Click
        FormmMasterASupplier.Show()
    End Sub

    Private Sub KategoriBarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KategoriBarangToolStripMenuItem.Click
        FormMasterKategoriBarang.Show()
    End Sub

    Private Sub BarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarangToolStripMenuItem.Click
        FormMasterBarang.Show()
    End Sub

   
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        JAM2.Text = TimeOfDay
    End Sub

    Private Sub TransaksiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransaksiToolStripMenuItem.Click

    End Sub

    Private Sub FormPenerimaanBarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormPenerimaanBarangToolStripMenuItem.Click
        FormTransTerimaBarang.ShowDialog()
    End Sub

    Private Sub PenjualanBarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenjualanBarangToolStripMenuItem.Click
        FormTransPenjualanBarang.ShowDialog()
    End Sub
End Class
