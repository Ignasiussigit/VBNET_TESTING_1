Imports System.Data.OleDb
Public Class FormLogin

    Sub TextBersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Sub Terbuka()
        FormMenuUtama.LoginToolStripMenuItem.Enabled = False
        FormMenuUtama.LogOutToolStripMenuItem.Enabled = True
        FormMenuUtama.MasterToolStripMenuItem.Visible = True
        FormMenuUtama.TransaksiToolStripMenuItem.Visible = True
        FormMenuUtama.LaporanToolStripMenuItem.Visible = True
        'Dibawah ini untuk membuka toolstrip  
        FormMenuUtama.MasterToolStripMenuItem.Enabled = True
        FormMenuUtama.TransaksiToolStripMenuItem.Enabled = True
        FormMenuUtama.LaporanToolStripMenuItem.Enabled = True
    End Sub
    Private Sub FormLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Terbuka()
    End Sub

    '===============================================
    ' CODE DIBAWAH UNTUK TOMBOL LOGIN
    '===============================================

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Terkunci()
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("KodeAdmin dan Password Harap diisi !", "INFO")
        Else
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "' and PasswordAdmin='" & TextBox2.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                MsgBox("Berhasil Masuk ...")
                Me.Close()
                'Call Terkunci()
                Call Terbuka()
                FormMenuUtama.USER2.Text = Rd!NamaAdmin
                FormMenuUtama.LEVEL2.Text = Rd!LevelAdmin
                'FormTransTerimaBarang.Label17.Text = Rd!NamaAdmin
            Else
                MsgBox("Data Tidak Ada ...")
                Call Terkunci()
                Call TextBersih()
                TextBox1.Focus()
            End If
        End If
    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            TextBox2.PasswordChar = ""
        Else
            TextBox2.PasswordChar = "•"
        End If
    End Sub
End Class