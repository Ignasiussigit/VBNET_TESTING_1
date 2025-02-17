Imports System.Data.OleDb

Public Class FormMasterAdmin

    Sub KondisiAwal()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_ADMIN", Conn)
        Ds = New DataSet()
        Ds.Clear()
        Da.Fill(Ds, "TBL_ADMIN")
        DataGridView1.DataSource = (Ds.Tables("TBL_ADMIN"))

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("ADMIN")
        ComboBox1.Items.Add("USER")
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList

        TextBox2.Enabled = False
        TextBox3.Enabled = False
    End Sub

    Sub TextBersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
    End Sub


    Sub TextMuteHapus()
        TextBox2.Enabled = False
        TextBox3.Enabled = False
    End Sub

    Sub inpt()
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        TextBox1.Enabled = False
    End Sub




    Private Sub FormMasterAdmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            Button1.Text = "SIMPAN"
            Call inpt()
            TextBox1.Text = ""
            Button2.Enabled = False
            Button3.Enabled = False
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MessageBox.Show("Data Harus TerIsi Semua-nya Cuyy ...", "INFO AJA")
                Button1.Text = "INPUT"
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                Button2.Enabled = True
                Button3.Enabled = True
            Else
                Call Koneksi()
                Dim SimpanData As String = "insert into TBL_ADMIN values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
                cmd = New OleDbCommand(SimpanData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Ter-Input ...")
                Button2.Enabled = True
                Button3.Enabled = True
                Call KondisiAwal()
                Call TextBersih()
            End If
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            Button2.Text = "SIMPAN"
            Button1.Enabled = False
            Button3.Enabled = False
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MessageBox.Show("Data Harus TerIsi Semua-nya Cuyy ...", "INFO AJA")
                Button2.Text = "EDIT"
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                Button1.Enabled = True
                Button3.Enabled = True
            Else
                Call Koneksi()
                Dim EditData As String = "UPDATE TBL_ADMIN set NamaAdmin='" & TextBox2.Text & "', PasswordAdmin='" & TextBox3.Text & "', LevelAdmin='" & ComboBox1.Text & "' where KodeAdmin='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(EditData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Ter-Edit ...")
                Button1.Enabled = True
                Button3.Enabled = True
                Call KondisiAwal()
                Call TextBersih()
            End If
        End If
       
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Data Harus TerIsi Semua-nya Cuyy ...", "INFO AJA")
        Else
            Call Koneksi()
            If MessageBox.Show("Anda Yakin Mengkhapus Kolom Ini ? ", "PERHATIAN", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim HapusData As String = "DELETE from TBL_ADMIN  where KodeAdmin='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(HapusData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Ter-Hapus ...")
                Call KondisiAwal()
                Call TextBersih()
            End If    
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaAdmin")
                TextBox3.Text = Rd.Item("PasswordAdmin")
                ComboBox1.Text = Rd.Item("LevelAdmin")
            Else
                MsgBox("Data Tidak Di Temukan ...")
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            TextBox3.PasswordChar = ""
        Else
            TextBox3.PasswordChar = "•"
        End If
    End Sub
End Class