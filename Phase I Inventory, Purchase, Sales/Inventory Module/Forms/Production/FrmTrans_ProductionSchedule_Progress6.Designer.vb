<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Progress6
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
        Me.LblProgress = New System.Windows.Forms.Label
        Me.prg1 = New System.Windows.Forms.ProgressBar
        Me.LblDescription = New System.Windows.Forms.Label
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.SuspendLayout()
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(12, 63)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 5
        Me.LblProgress.Text = "Label1"
        '
        'prg1
        '
        Me.prg1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prg1.Location = New System.Drawing.Point(12, 33)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(625, 27)
        Me.prg1.TabIndex = 4
        '
        'LblDescription
        '
        Me.LblDescription.AutoSize = True
        Me.LblDescription.Location = New System.Drawing.Point(12, 13)
        Me.LblDescription.Name = "LblDescription"
        Me.LblDescription.Size = New System.Drawing.Size(51, 17)
        Me.LblDescription.TabIndex = 3
        Me.LblDescription.Text = "Label1"
        '
        'FrmTrans_ProductionSchedule_Progress6
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(649, 93)
        Me.ControlBox = False
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.prg1)
        Me.Controls.Add(Me.LblDescription)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(657, 130)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(657, 130)
        Me.Name = "FrmTrans_ProductionSchedule_Progress6"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Rescheduling, please wait"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblDescription As System.Windows.Forms.Label
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
End Class
