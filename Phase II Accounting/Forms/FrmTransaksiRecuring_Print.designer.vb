<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiRecuring_Print
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
        Me.CrvRecurring = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(12, 12)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 0
        Me.BtnClose.Text = "Button1"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'CrvRecurring
        '
        Me.CrvRecurring.ActiveViewIndex = -1
        Me.CrvRecurring.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvRecurring.DisplayGroupTree = False
        Me.CrvRecurring.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvRecurring.EnableDrillDown = False
        Me.CrvRecurring.Location = New System.Drawing.Point(0, 0)
        Me.CrvRecurring.Name = "CrvRecurring"
        Me.CrvRecurring.SelectionFormula = ""
        Me.CrvRecurring.ShowCloseButton = False
        Me.CrvRecurring.ShowGroupTreeButton = False
        Me.CrvRecurring.ShowRefreshButton = False
        Me.CrvRecurring.Size = New System.Drawing.Size(784, 561)
        Me.CrvRecurring.TabIndex = 1
        Me.CrvRecurring.ViewTimeSelectionFormula = ""
        '
        'FrmTransaksiRecuring_Print
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.CrvRecurring)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimumSize = New System.Drawing.Size(800, 600)
        Me.Name = "FrmTransaksiRecuring_Print"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Recurring"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents CrvRecurring As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
