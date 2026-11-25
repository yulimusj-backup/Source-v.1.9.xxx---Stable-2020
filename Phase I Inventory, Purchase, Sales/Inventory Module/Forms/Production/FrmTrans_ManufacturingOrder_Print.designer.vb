<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_Print
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
        Me.CrvMO = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SetDefaultPrinterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BtnClose = New System.Windows.Forms.Button
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CrvMO
        '
        Me.CrvMO.ActiveViewIndex = -1
        Me.CrvMO.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvMO.ContextMenuStrip = Me.ContextMenuStrip1
        Me.CrvMO.DisplayGroupTree = False
        Me.CrvMO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvMO.EnableDrillDown = False
        Me.CrvMO.Location = New System.Drawing.Point(0, 0)
        Me.CrvMO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CrvMO.Name = "CrvMO"
        Me.CrvMO.SelectionFormula = ""
        Me.CrvMO.ShowCloseButton = False
        Me.CrvMO.ShowGroupTreeButton = False
        Me.CrvMO.ShowRefreshButton = False
        Me.CrvMO.Size = New System.Drawing.Size(1045, 690)
        Me.CrvMO.TabIndex = 0
        Me.CrvMO.ViewTimeSelectionFormula = ""
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SetDefaultPrinterToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(199, 28)
        '
        'SetDefaultPrinterToolStripMenuItem
        '
        Me.SetDefaultPrinterToolStripMenuItem.Name = "SetDefaultPrinterToolStripMenuItem"
        Me.SetDefaultPrinterToolStripMenuItem.Size = New System.Drawing.Size(198, 24)
        Me.SetDefaultPrinterToolStripMenuItem.Text = "Set default printer"
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(929, 15)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 28)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'FrmTrans_ManufacturingOrder_Print
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(1045, 690)
        Me.Controls.Add(Me.CrvMO)
        Me.Controls.Add(Me.BtnClose)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "FrmTrans_ManufacturingOrder_Print"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Manufacturing Order"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrvMO As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents SetDefaultPrinterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
