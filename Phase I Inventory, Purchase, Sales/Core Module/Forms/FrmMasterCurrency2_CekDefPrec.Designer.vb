<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCurrency2_CekDefPrec
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
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.btnCancel = New System.Windows.Forms.Button
        Me.prg1 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'bgw1
        '
        Me.bgw1.WorkerReportsProgress = True
        Me.bgw1.WorkerSupportsCancellation = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(634, 12)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'prg1
        '
        Me.prg1.Location = New System.Drawing.Point(12, 12)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(616, 23)
        Me.prg1.TabIndex = 1
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(12, 38)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 2
        Me.LblProgress.Text = "Label1"
        '
        'FrmMasterCurrency2_CekDefPrec
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(721, 85)
        Me.ControlBox = False
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.prg1)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.Name = "FrmMasterCurrency2_CekDefPrec"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Checking for inconsistency"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
End Class
