<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakFakturDP
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
        Me.CRVfaktur = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'CRVfaktur
        '
        Me.CRVfaktur.ActiveViewIndex = -1
        Me.CRVfaktur.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRVfaktur.DisplayGroupTree = False
        Me.CRVfaktur.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CRVfaktur.Location = New System.Drawing.Point(0, 0)
        Me.CRVfaktur.Name = "CRVfaktur"
        Me.CRVfaktur.SelectionFormula = ""
        Me.CRVfaktur.ShowCloseButton = False
        Me.CRVfaktur.ShowGroupTreeButton = False
        Me.CRVfaktur.ShowRefreshButton = False
        Me.CRVfaktur.Size = New System.Drawing.Size(858, 505)
        Me.CRVfaktur.TabIndex = 2
        Me.CRVfaktur.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button1.Location = New System.Drawing.Point(611, 136)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmCetakFakturDP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(858, 505)
        Me.Controls.Add(Me.CRVfaktur)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(874, 541)
        Me.Name = "FrmCetakFakturDP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tax Invoice"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CRVfaktur As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
