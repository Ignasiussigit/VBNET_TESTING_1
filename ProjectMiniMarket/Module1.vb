Imports System.Data.OleDb
Module Module1
    Public Conn As OleDbConnection
    Public Da As OleDbDataAdapter
    Public Ds As DataSet
    Public LokasiDb As String
    Public cmd As OleDbCommand
    Public Rd As OleDbDataReader

    Sub Terkunci()
        FormMenuUtama.LoginToolStripMenuItem.Enabled = True
        FormMenuUtama.LogOutToolStripMenuItem.Enabled = False
        FormMenuUtama.MasterToolStripMenuItem.Enabled = False
        FormMenuUtama.TransaksiToolStripMenuItem.Enabled = False
        FormMenuUtama.LaporanToolStripMenuItem.Enabled = False
    End Sub

    Sub Koneksi()
        LokasiDb = "Provider=Microsoft.ACE.OleDb.12.0;Data Source=DBMiniMarket.accdb"
        Conn = New OleDbConnection(LokasiDb)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
End Module
