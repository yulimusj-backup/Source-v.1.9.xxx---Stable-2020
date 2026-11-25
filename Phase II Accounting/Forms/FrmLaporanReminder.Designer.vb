<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanReminder
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
        Me.CrvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'CrvReport
        '
        Me.CrvReport.ActiveViewIndex = -1
        Me.CrvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvReport.DisplayGroupTree = False
        Me.CrvReport.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvReport.Location = New System.Drawing.Point(0, 0)
        Me.CrvReport.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CrvReport.Name = "CrvReport"
        Me.CrvReport.SelectionFormula = ""
        Me.CrvReport.ShowCloseButton = False
        Me.CrvReport.ShowGroupTreeButton = False
        Me.CrvReport.ShowRefreshButton = False
        Me.CrvReport.Size = New System.Drawing.Size(992, 633)
        Me.CrvReport.TabIndex = 282
        Me.CrvReport.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(572, 171)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(134, 28)
        Me.Button1.TabIndex = 283
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmLaporanReminder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.CrvReport)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanReminder"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Account Receivable Due"
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents CrvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
