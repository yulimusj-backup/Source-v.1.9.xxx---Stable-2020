<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Print
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
        Me.CrvLaporan = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'CrvLaporan
        '
        Me.CrvLaporan.ActiveViewIndex = -1
        Me.CrvLaporan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvLaporan.DisplayGroupTree = False
        Me.CrvLaporan.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvLaporan.EnableDrillDown = False
        Me.CrvLaporan.Location = New System.Drawing.Point(0, 0)
        Me.CrvLaporan.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CrvLaporan.Name = "CrvLaporan"
        Me.CrvLaporan.SelectionFormula = ""
        Me.CrvLaporan.ShowCloseButton = False
        Me.CrvLaporan.ShowGroupTreeButton = False
        Me.CrvLaporan.ShowRefreshButton = False
        Me.CrvLaporan.Size = New System.Drawing.Size(995, 634)
        Me.CrvLaporan.TabIndex = 3
        Me.CrvLaporan.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(557, 100)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_Print
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(995, 634)
        Me.Controls.Add(Me.CrvLaporan)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1003, 671)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Print"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Planning"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrvLaporan As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
