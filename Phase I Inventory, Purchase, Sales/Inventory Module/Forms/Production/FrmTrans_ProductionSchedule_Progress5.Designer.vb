<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Progress5
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
        Me.prg1 = New System.Windows.Forms.ProgressBar
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.SuspendLayout()
        '
        'prg1
        '
        Me.prg1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prg1.Location = New System.Drawing.Point(12, 12)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(625, 27)
        Me.prg1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.prg1.TabIndex = 10
        '
        'FrmTrans_ProductionSchedule_Progress5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(649, 63)
        Me.ControlBox = False
        Me.Controls.Add(Me.prg1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(657, 130)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(657, 100)
        Me.Name = "FrmTrans_ProductionSchedule_Progress5"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Showing Data"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
End Class
