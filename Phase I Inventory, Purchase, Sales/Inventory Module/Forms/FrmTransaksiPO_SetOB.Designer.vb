<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPO_SetOB
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
        Me.btnStart = New System.Windows.Forms.Button
        Me.Prb01 = New System.Windows.Forms.ProgressBar
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(324, 9)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(54, 21)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Prb01
        '
        Me.Prb01.Location = New System.Drawing.Point(12, 9)
        Me.Prb01.Name = "Prb01"
        Me.Prb01.Size = New System.Drawing.Size(306, 21)
        Me.Prb01.TabIndex = 1
        '
        'BackgroundWorker1
        '
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(324, 36)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(54, 21)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Start"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiPO_SetOB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(390, 54)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Prb01)
        Me.Controls.Add(Me.btnStart)
        Me.Name = "FrmTransaksiPO_SetOB"
        Me.Text = "FrmTransaksiPO_SetOB"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents Prb01 As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
