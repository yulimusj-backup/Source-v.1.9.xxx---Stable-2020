<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStockTaking_FinalCheck
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
        Me.vcrv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'vcrv
        '
        Me.vcrv.ActiveViewIndex = -1
        Me.vcrv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.vcrv.DisplayGroupTree = False
        Me.vcrv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.vcrv.Location = New System.Drawing.Point(0, 0)
        Me.vcrv.Name = "vcrv"
        Me.vcrv.SelectionFormula = ""
        Me.vcrv.ShowCloseButton = False
        Me.vcrv.ShowGroupTreeButton = False
        Me.vcrv.ShowRefreshButton = False
        Me.vcrv.Size = New System.Drawing.Size(891, 493)
        Me.vcrv.TabIndex = 2
        Me.vcrv.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(365, 87)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 33)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmLaporanStockTaking_FinalCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(891, 493)
        Me.Controls.Add(Me.vcrv)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(907, 529)
        Me.Name = "FrmLaporanStockTaking_FinalCheck"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Final Data Check"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents vcrv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
