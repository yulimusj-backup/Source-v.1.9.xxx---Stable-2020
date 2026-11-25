<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_Cetak
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
        Me.CRV1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BtnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'CRV1
        '
        Me.CRV1.ActiveViewIndex = -1
        Me.CRV1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRV1.DisplayGroupTree = False
        Me.CRV1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CRV1.Location = New System.Drawing.Point(0, 0)
        Me.CRV1.Name = "CRV1"
        Me.CRV1.SelectionFormula = ""
        Me.CRV1.ShowCloseButton = False
        Me.CRV1.ShowGroupTreeButton = False
        Me.CRV1.ShowRefreshButton = False
        Me.CRV1.Size = New System.Drawing.Size(840, 403)
        Me.CRV1.TabIndex = 0
        Me.CRV1.ViewTimeSelectionFormula = ""
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(619, 77)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(166, 28)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_POS_Cetak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(840, 403)
        Me.Controls.Add(Me.CRV1)
        Me.Controls.Add(Me.BtnClose)
        Me.Name = "FrmTrans_POS_Cetak"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cetak"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CRV1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BtnClose As System.Windows.Forms.Button
End Class
