<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2_UpdateCategory
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
        Me.lbl1 = New System.Windows.Forms.Label
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.SuspendLayout()
        '
        'prg1
        '
        Me.prg1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prg1.Location = New System.Drawing.Point(12, 12)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(573, 23)
        Me.prg1.TabIndex = 0
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(12, 38)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(51, 17)
        Me.lbl1.TabIndex = 1
        Me.lbl1.Text = "Label1"
        '
        'bgw1
        '
        '
        'Form2_UpdateCategory
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(597, 75)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.prg1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form2_UpdateCategory"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update Category"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
End Class
