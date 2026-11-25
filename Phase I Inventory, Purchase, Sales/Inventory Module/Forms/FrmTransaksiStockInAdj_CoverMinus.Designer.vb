<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInAdj_CoverMinus
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
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.prb1 = New System.Windows.Forms.ProgressBar
        Me.btnCancel = New System.Windows.Forms.Button
        Me.LblProgres = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'BackgroundWorker1
        '
        '
        'prb1
        '
        Me.prb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prb1.Location = New System.Drawing.Point(12, 12)
        Me.prb1.Name = "prb1"
        Me.prb1.Size = New System.Drawing.Size(475, 23)
        Me.prb1.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(493, 12)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'LblProgres
        '
        Me.LblProgres.AutoSize = True
        Me.LblProgres.Location = New System.Drawing.Point(12, 38)
        Me.LblProgres.Name = "LblProgres"
        Me.LblProgres.Size = New System.Drawing.Size(51, 17)
        Me.LblProgres.TabIndex = 2
        Me.LblProgres.Text = "Label1"
        '
        'FrmTransaksiStockInAdj_CoverMinus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(580, 77)
        Me.ControlBox = False
        Me.Controls.Add(Me.LblProgres)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.prb1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(586, 83)
        Me.Name = "FrmTransaksiStockInAdj_CoverMinus"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Creating Stock In Adjust Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents prb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents LblProgres As System.Windows.Forms.Label

End Class
