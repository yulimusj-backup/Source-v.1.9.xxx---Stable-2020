<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaporanInventoryCheckTag
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
        Me.CrvLaporan.Location = New System.Drawing.Point(0, 0)
        Me.CrvLaporan.Name = "CrvLaporan"
        Me.CrvLaporan.SelectionFormula = ""
        Me.CrvLaporan.ShowCloseButton = False
        Me.CrvLaporan.ShowGroupTreeButton = False
        Me.CrvLaporan.ShowRefreshButton = False
        Me.CrvLaporan.Size = New System.Drawing.Size(890, 471)
        Me.CrvLaporan.TabIndex = 2
        Me.CrvLaporan.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(488, 138)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(219, 28)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmLaporanInventoryCheckTag
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(890, 471)
        Me.Controls.Add(Me.CrvLaporan)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(896, 497)
        Me.Name = "frmLaporanInventoryCheckTag"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Check Tag"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrvLaporan As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
