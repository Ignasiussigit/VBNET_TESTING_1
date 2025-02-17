<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenuUtama
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoginToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogOutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KeluarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdminToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KategoriBarangToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BarangToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransaksiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FormPenerimaanBarangToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.PenjualanBarangToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.USER1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.USER2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LEVEL1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LEVEL2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.JAM1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.JAM2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TANGGAL1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TANGGAL2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.CadetBlue
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.MasterToolStripMenuItem, Me.TransaksiToolStripMenuItem, Me.LaporanToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(607, 28)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LoginToolStripMenuItem, Me.LogOutToolStripMenuItem, Me.KeluarToolStripMenuItem})
        Me.FileToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(45, 24)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'LoginToolStripMenuItem
        '
        Me.LoginToolStripMenuItem.Name = "LoginToolStripMenuItem"
        Me.LoginToolStripMenuItem.Size = New System.Drawing.Size(130, 24)
        Me.LoginToolStripMenuItem.Text = "LogIn"
        '
        'LogOutToolStripMenuItem
        '
        Me.LogOutToolStripMenuItem.Name = "LogOutToolStripMenuItem"
        Me.LogOutToolStripMenuItem.Size = New System.Drawing.Size(130, 24)
        Me.LogOutToolStripMenuItem.Text = "LogOut"
        '
        'KeluarToolStripMenuItem
        '
        Me.KeluarToolStripMenuItem.Name = "KeluarToolStripMenuItem"
        Me.KeluarToolStripMenuItem.Size = New System.Drawing.Size(130, 24)
        Me.KeluarToolStripMenuItem.Text = "Keluar"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AdminToolStripMenuItem, Me.SupplierToolStripMenuItem, Me.KategoriBarangToolStripMenuItem, Me.BarangToolStripMenuItem})
        Me.MasterToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(62, 24)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'AdminToolStripMenuItem
        '
        Me.AdminToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem"
        Me.AdminToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.AdminToolStripMenuItem.Text = "Admin"
        '
        'SupplierToolStripMenuItem
        '
        Me.SupplierToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SupplierToolStripMenuItem.Name = "SupplierToolStripMenuItem"
        Me.SupplierToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.SupplierToolStripMenuItem.Text = "Supplier"
        '
        'KategoriBarangToolStripMenuItem
        '
        Me.KategoriBarangToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KategoriBarangToolStripMenuItem.Name = "KategoriBarangToolStripMenuItem"
        Me.KategoriBarangToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.KategoriBarangToolStripMenuItem.Text = "KategoriBarang"
        '
        'BarangToolStripMenuItem
        '
        Me.BarangToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BarangToolStripMenuItem.Name = "BarangToolStripMenuItem"
        Me.BarangToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.BarangToolStripMenuItem.Text = "Barang"
        '
        'TransaksiToolStripMenuItem
        '
        Me.TransaksiToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FormPenerimaanBarangToolStripMenuItem, Me.ToolStripMenuItem1, Me.PenjualanBarangToolStripMenuItem})
        Me.TransaksiToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TransaksiToolStripMenuItem.Name = "TransaksiToolStripMenuItem"
        Me.TransaksiToolStripMenuItem.Size = New System.Drawing.Size(86, 24)
        Me.TransaksiToolStripMenuItem.Text = "Transaksi"
        '
        'FormPenerimaanBarangToolStripMenuItem
        '
        Me.FormPenerimaanBarangToolStripMenuItem.Name = "FormPenerimaanBarangToolStripMenuItem"
        Me.FormPenerimaanBarangToolStripMenuItem.Size = New System.Drawing.Size(215, 24)
        Me.FormPenerimaanBarangToolStripMenuItem.Text = "Penerimaan Barang"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(212, 6)
        '
        'PenjualanBarangToolStripMenuItem
        '
        Me.PenjualanBarangToolStripMenuItem.Name = "PenjualanBarangToolStripMenuItem"
        Me.PenjualanBarangToolStripMenuItem.Size = New System.Drawing.Size(215, 24)
        Me.PenjualanBarangToolStripMenuItem.Text = "Penjualan Barang"
        '
        'LaporanToolStripMenuItem
        '
        Me.LaporanToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LaporanToolStripMenuItem.Name = "LaporanToolStripMenuItem"
        Me.LaporanToolStripMenuItem.Size = New System.Drawing.Size(78, 24)
        Me.LaporanToolStripMenuItem.Text = "Laporan"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.USER1, Me.USER2, Me.LEVEL1, Me.LEVEL2, Me.JAM1, Me.JAM2, Me.TANGGAL1, Me.TANGGAL2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 350)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(607, 22)
        Me.StatusStrip1.TabIndex = 1
        '
        'USER1
        '
        Me.USER1.Name = "USER1"
        Me.USER1.Size = New System.Drawing.Size(40, 17)
        Me.USER1.Text = "USER :"
        '
        'USER2
        '
        Me.USER2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.USER2.Name = "USER2"
        Me.USER2.Size = New System.Drawing.Size(14, 17)
        Me.USER2.Text = "?"
        '
        'LEVEL1
        '
        Me.LEVEL1.Name = "LEVEL1"
        Me.LEVEL1.Size = New System.Drawing.Size(44, 17)
        Me.LEVEL1.Text = "LEVEL :"
        '
        'LEVEL2
        '
        Me.LEVEL2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LEVEL2.Name = "LEVEL2"
        Me.LEVEL2.Size = New System.Drawing.Size(14, 17)
        Me.LEVEL2.Text = "?"
        '
        'JAM1
        '
        Me.JAM1.Name = "JAM1"
        Me.JAM1.Size = New System.Drawing.Size(36, 17)
        Me.JAM1.Text = "JAM :"
        '
        'JAM2
        '
        Me.JAM2.Name = "JAM2"
        Me.JAM2.Size = New System.Drawing.Size(0, 17)
        '
        'TANGGAL1
        '
        Me.TANGGAL1.Name = "TANGGAL1"
        Me.TANGGAL1.Size = New System.Drawing.Size(65, 17)
        Me.TANGGAL1.Text = "TANGGAL :"
        '
        'TANGGAL2
        '
        Me.TANGGAL2.Name = "TANGGAL2"
        Me.TANGGAL2.Size = New System.Drawing.Size(0, 17)
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'FormMenuUtama
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(607, 372)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenuUtama"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoginToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LogOutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KeluarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AdminToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KategoriBarangToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BarangToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransaksiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents USER1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents USER2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents LEVEL1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents LEVEL2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents JAM1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents JAM2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TANGGAL1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TANGGAL2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents FormPenerimaanBarangToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents PenjualanBarangToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
