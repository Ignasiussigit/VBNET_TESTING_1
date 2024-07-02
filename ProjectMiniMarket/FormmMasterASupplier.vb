Imports System.Data.OleDb
Public Class FormmMasterASupplier
    'Dim Conn As OleDbConnection
    'Dim Da As OleDbDataAdapter
    'Dim Ds As DataSet
    'Dim LokasiDb As String
    'Dim cmd As OleDbCommand
    'Dim Rd As OleDbDataReader

    Sub TablesMe()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_SUPPLIER", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "TBL_SUPPLIER")
        DataGridView1.DataSource = (Ds.Tables("TBL_SUPPLIER"))
    End Sub

    Sub TextKosong()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub


    Private Sub FormmMasterASupplier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TablesMe()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MessageBox.Show("Data harus tersisi semau-nya ...")
        Else
            Call Koneksi()
            Dim SimpanData As String = "Insert into TBL_SUPPLIER values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
            cmd = New OleDbCommand(SimpanData, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data berhasil Ter-Input ...")
            Call TextKosong()
            Call TablesMe()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MessageBox.Show("Data harus tersisi semau-nya ...")
        Else
            Call Koneksi()
            Dim EditData As String = "UPDATE TBL_SUPPLIER set NamaSupplier='" & TextBox2.Text & "', AlamatSupplier='" & TextBox3.Text & "', TelpSupplier='" & TextBox4.Text & "' where KodeSupplier='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(EditData, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data berhasil Ter-Edit ...")
            Call TextKosong()
            Call TablesMe()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Data KodeSupplier harus terisi ...")
        Else
            If MessageBox.Show("Anda Yakin Ingin menghapus ini ? ", "INFO CUYY", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                cmd = New OleDbCommand("Select * From TBL_SUPPLIER where KodeSupplier='" & TextBox1.Text & "'", Conn)
                Rd = cmd.ExecuteReader
                Rd.Read()
                If Rd.HasRows Then
                    Dim HapusData As String = "DELETE From TBL_SUPPLIER where KodeSupplier='" & TextBox1.Text & "'"
                    cmd = New OleDbCommand(HapusData, Conn)
                    cmd.ExecuteNonQuery()
                    MsgBox("Data berhasil Ter-Hapus ...")
                    Call TextKosong()
                    Call TablesMe()
                Else
                    MsgBox(" Data tidak ada ... ")
                End If
            End If 
        End If

            
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_SUPPLIER where KodeSupplier='" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaSupplier")
                TextBox3.Text = Rd.Item("AlamatSupplier")
                TextBox4.Text = Rd.Item("TelpSupplier")
            Else
                MsgBox(" Data tidak ada ... ")
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class