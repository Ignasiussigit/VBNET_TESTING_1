'Imports System.Data.OleDb
Public Class FormMenuUtama

    Sub Terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogOutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Enabled = False
        TransaksiToolStripMenuItem.Enabled = False
        LaporanToolStripMenuItem.Enabled = False
    End Sub
    'UNTUK SUB TERBUKA-NYA ADA DI FORM LOGIN


    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        Dim TabLogin As New FormLogin()

        TabLogin.Show()
    End Sub

    Private Sub FormMenuUtama_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Terkunci()
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

        Master.Show()
    End Sub

    Private Sub SupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierToolStripMenuItem.Click
        FormmMasterASupplier.Show()
    End Sub

    Private Sub KategoriBarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KategoriBarangToolStripMenuItem.Click
        FormMasterKategoriBarang.Show()
    End Sub
End Class
