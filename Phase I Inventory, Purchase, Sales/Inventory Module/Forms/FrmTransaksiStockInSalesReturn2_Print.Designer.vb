<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInSalesReturn2_Print
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
        Me.BtnClose = New System.Windows.Forms.Button
        Me.CrvStockInSalesReturn = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(341, 13)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 0
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'CrvStockInSalesReturn
        '
        Me.CrvStockInSalesReturn.ActiveViewIndex = -1
        Me.CrvStockInSalesReturn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvStockInSalesReturn.DisplayGroupTree = False
        Me.CrvStockInSalesReturn.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvStockInSalesReturn.EnableDrillDown = False
        Me.CrvStockInSalesReturn.Location = New System.Drawing.Point(0, 0)
        Me.CrvStockInSalesReturn.Name = "CrvStockInSalesReturn"
        Me.CrvStockInSalesReturn.SelectionFormula = ""
        Me.CrvStockInSalesReturn.ShowCloseButton = False
        Me.CrvStockInSalesReturn.ShowGroupTreeButton = False
        Me.CrvStockInSalesReturn.ShowRefreshButton = False
        Me.CrvStockInSalesReturn.Size = New System.Drawing.Size(866, 514)
        Me.CrvStockInSalesReturn.TabIndex = 1
        Me.CrvStockInSalesReturn.ViewTimeSelectionFormula = ""
        '
        'FrmTransaksiStockInSalesReturn_Cetak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(866, 514)
        Me.Controls.Add(Me.CrvStockInSalesReturn)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(874, 541)
        Me.Name = "FrmTransaksiStockInSalesReturn_Cetak"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Stock In Sales Return"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Public WithEvents CrvStockInSalesReturn As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
