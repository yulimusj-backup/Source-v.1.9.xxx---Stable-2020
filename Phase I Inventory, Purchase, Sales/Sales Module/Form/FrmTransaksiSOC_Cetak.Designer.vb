<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSOC_Cetak
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
        Me.CrvSOC = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'CrvSOC
        '
        Me.CrvSOC.ActiveViewIndex = -1
        Me.CrvSOC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvSOC.DisplayGroupTree = False
        Me.CrvSOC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvSOC.EnableDrillDown = False
        Me.CrvSOC.Location = New System.Drawing.Point(0, 0)
        Me.CrvSOC.Name = "CrvSOC"
        Me.CrvSOC.SelectionFormula = ""
        Me.CrvSOC.ShowCloseButton = False
        Me.CrvSOC.ShowGroupTreeButton = False
        Me.CrvSOC.ShowRefreshButton = False
        Me.CrvSOC.Size = New System.Drawing.Size(866, 514)
        Me.CrvSOC.TabIndex = 2
        Me.CrvSOC.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(184, 39)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmCetakSOC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(866, 514)
        Me.Controls.Add(Me.CrvSOC)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(874, 541)
        Me.Name = "FrmCetakSOC"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Sales Order Cancellation"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrvSOC As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
