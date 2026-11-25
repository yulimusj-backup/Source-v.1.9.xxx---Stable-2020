<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMain
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
        Me.components = New System.ComponentModel.Container
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FrmMasterProduct1 = New System.Windows.Forms.ToolStripMenuItem
        Me.FrmMasterDiskon1 = New System.Windows.Forms.ToolStripMenuItem
        Me.SessionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MemberToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PayoutDropCashToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PaymentOptionToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.PengecekanFormatDecimalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ProductToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.InventoryTransactionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.POSSettingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SessionToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.POSPaymentDisbursementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.POSPaymentDisbursemenEditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.POSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.WindowsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CascadeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TileVericallyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TileHorizontallyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CloseAllWindowsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.FrmTrans_Data_Transfer1 = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MasterToolStripMenuItem, Me.InventoryTransactionToolStripMenuItem, Me.WindowsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.MdiWindowListItem = Me.WindowsToolStripMenuItem
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(4, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(628, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FrmMasterProduct1, Me.FrmMasterDiskon1, Me.SessionToolStripMenuItem, Me.MemberToolStripMenuItem, Me.PayoutDropCashToolStripMenuItem, Me.PaymentOptionToolStripMenuItem1, Me.PengecekanFormatDecimalToolStripMenuItem, Me.ProductToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasterToolStripMenuItem.Text = "&Master"
        '
        'FrmMasterProduct1
        '
        Me.FrmMasterProduct1.Name = "FrmMasterProduct1"
        Me.FrmMasterProduct1.Size = New System.Drawing.Size(225, 22)
        Me.FrmMasterProduct1.Text = "Branch"
        '
        'FrmMasterDiskon1
        '
        Me.FrmMasterDiskon1.Name = "FrmMasterDiskon1"
        Me.FrmMasterDiskon1.Size = New System.Drawing.Size(225, 22)
        Me.FrmMasterDiskon1.Text = "Discount"
        '
        'SessionToolStripMenuItem
        '
        Me.SessionToolStripMenuItem.Name = "SessionToolStripMenuItem"
        Me.SessionToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.SessionToolStripMenuItem.Text = "Session"
        '
        'MemberToolStripMenuItem
        '
        Me.MemberToolStripMenuItem.Name = "MemberToolStripMenuItem"
        Me.MemberToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.MemberToolStripMenuItem.Text = "Member"
        '
        'PayoutDropCashToolStripMenuItem
        '
        Me.PayoutDropCashToolStripMenuItem.Name = "PayoutDropCashToolStripMenuItem"
        Me.PayoutDropCashToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.PayoutDropCashToolStripMenuItem.Text = "Payout drop cash"
        '
        'PaymentOptionToolStripMenuItem1
        '
        Me.PaymentOptionToolStripMenuItem1.Name = "PaymentOptionToolStripMenuItem1"
        Me.PaymentOptionToolStripMenuItem1.Size = New System.Drawing.Size(225, 22)
        Me.PaymentOptionToolStripMenuItem1.Text = "Payment option"
        '
        'PengecekanFormatDecimalToolStripMenuItem
        '
        Me.PengecekanFormatDecimalToolStripMenuItem.Name = "PengecekanFormatDecimalToolStripMenuItem"
        Me.PengecekanFormatDecimalToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.PengecekanFormatDecimalToolStripMenuItem.Text = "Pengecekan Format Decimal"
        '
        'ProductToolStripMenuItem
        '
        Me.ProductToolStripMenuItem.Name = "ProductToolStripMenuItem"
        Me.ProductToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.ProductToolStripMenuItem.Text = "Product"
        '
        'InventoryTransactionToolStripMenuItem
        '
        Me.InventoryTransactionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.POSSettingToolStripMenuItem, Me.SessionToolStripMenuItem1, Me.POSPaymentDisbursementToolStripMenuItem, Me.POSPaymentDisbursemenEditToolStripMenuItem, Me.POSToolStripMenuItem, Me.FrmTrans_Data_Transfer1})
        Me.InventoryTransactionToolStripMenuItem.Name = "InventoryTransactionToolStripMenuItem"
        Me.InventoryTransactionToolStripMenuItem.Size = New System.Drawing.Size(68, 20)
        Me.InventoryTransactionToolStripMenuItem.Text = "&Transaksi"
        '
        'POSSettingToolStripMenuItem
        '
        Me.POSSettingToolStripMenuItem.Name = "POSSettingToolStripMenuItem"
        Me.POSSettingToolStripMenuItem.Size = New System.Drawing.Size(245, 22)
        Me.POSSettingToolStripMenuItem.Text = "POS Setting"
        '
        'SessionToolStripMenuItem1
        '
        Me.SessionToolStripMenuItem1.Name = "SessionToolStripMenuItem1"
        Me.SessionToolStripMenuItem1.Size = New System.Drawing.Size(245, 22)
        Me.SessionToolStripMenuItem1.Text = "Session"
        '
        'POSPaymentDisbursementToolStripMenuItem
        '
        Me.POSPaymentDisbursementToolStripMenuItem.Name = "POSPaymentDisbursementToolStripMenuItem"
        Me.POSPaymentDisbursementToolStripMenuItem.Size = New System.Drawing.Size(245, 22)
        Me.POSPaymentDisbursementToolStripMenuItem.Text = "POS Payment Disbursement"
        '
        'POSPaymentDisbursemenEditToolStripMenuItem
        '
        Me.POSPaymentDisbursemenEditToolStripMenuItem.Name = "POSPaymentDisbursemenEditToolStripMenuItem"
        Me.POSPaymentDisbursemenEditToolStripMenuItem.Size = New System.Drawing.Size(245, 22)
        Me.POSPaymentDisbursemenEditToolStripMenuItem.Text = "POS Payment Disbursement Edit"
        '
        'POSToolStripMenuItem
        '
        Me.POSToolStripMenuItem.Name = "POSToolStripMenuItem"
        Me.POSToolStripMenuItem.Size = New System.Drawing.Size(245, 22)
        Me.POSToolStripMenuItem.Text = "POS "
        '
        'WindowsToolStripMenuItem
        '
        Me.WindowsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CascadeToolStripMenuItem, Me.TileVericallyToolStripMenuItem, Me.TileHorizontallyToolStripMenuItem, Me.CloseAllWindowsToolStripMenuItem})
        Me.WindowsToolStripMenuItem.Name = "WindowsToolStripMenuItem"
        Me.WindowsToolStripMenuItem.Size = New System.Drawing.Size(68, 20)
        Me.WindowsToolStripMenuItem.Text = "&Windows"
        '
        'CascadeToolStripMenuItem
        '
        Me.CascadeToolStripMenuItem.Name = "CascadeToolStripMenuItem"
        Me.CascadeToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.CascadeToolStripMenuItem.Text = "Cascade"
        '
        'TileVericallyToolStripMenuItem
        '
        Me.TileVericallyToolStripMenuItem.Name = "TileVericallyToolStripMenuItem"
        Me.TileVericallyToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.TileVericallyToolStripMenuItem.Text = "Tile Vertically"
        '
        'TileHorizontallyToolStripMenuItem
        '
        Me.TileHorizontallyToolStripMenuItem.Name = "TileHorizontallyToolStripMenuItem"
        Me.TileHorizontallyToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.TileHorizontallyToolStripMenuItem.Text = "Tile Horizontally"
        '
        'CloseAllWindowsToolStripMenuItem
        '
        Me.CloseAllWindowsToolStripMenuItem.Name = "CloseAllWindowsToolStripMenuItem"
        Me.CloseAllWindowsToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.CloseAllWindowsToolStripMenuItem.Text = "Close All Windows"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 353)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 12, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(628, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Text = "Internal Mail Service - click here to display email page"
        Me.NotifyIcon1.Visible = True
        '
        'FrmTrans_Data_Transfer1
        '
        Me.FrmTrans_Data_Transfer1.Name = "FrmTrans_Data_Transfer1"
        Me.FrmTrans_Data_Transfer1.Size = New System.Drawing.Size(245, 22)
        Me.FrmTrans_Data_Transfer1.Text = "Transfer Data"
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(628, 375)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FrmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "POS Module"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents InventoryTransactionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents WindowsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CascadeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TileVericallyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TileHorizontallyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CloseAllWindowsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FrmMasterProduct1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FrmMasterDiskon1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SessionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POSSettingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SessionToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POSPaymentDisbursementToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POSPaymentDisbursemenEditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MemberToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PayoutDropCashToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PaymentOptionToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PengecekanFormatDecimalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProductToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FrmTrans_Data_Transfer1 As System.Windows.Forms.ToolStripMenuItem
End Class
