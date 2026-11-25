<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakTransaksiStockInAdj
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
        Me.crvStockCard = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'crvStockCard
        '
        Me.crvStockCard.ActiveViewIndex = -1
        Me.crvStockCard.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvStockCard.DisplayGroupTree = False
        Me.crvStockCard.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvStockCard.Location = New System.Drawing.Point(0, 0)
        Me.crvStockCard.Name = "crvStockCard"
        Me.crvStockCard.SelectionFormula = ""
        Me.crvStockCard.ShowCloseButton = False
        Me.crvStockCard.ShowGroupTreeButton = False
        Me.crvStockCard.ShowRefreshButton = False
        Me.crvStockCard.Size = New System.Drawing.Size(858, 505)
        Me.crvStockCard.TabIndex = 1
        Me.crvStockCard.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(469, 131)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(173, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmCetakTransaksiStockInAdj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(858, 505)
        Me.Controls.Add(Me.crvStockCard)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(874, 541)
        Me.Name = "FrmCetakTransaksiStockInAdj"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock In - Adjust"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crvStockCard As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
