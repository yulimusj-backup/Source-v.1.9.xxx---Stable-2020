<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransferBeginingBalance_Error
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
        Me.txtError = New System.Windows.Forms.TextBox
        Me.btnContinue = New System.Windows.Forms.Button
        Me.btnCancel2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtError
        '
        Me.txtError.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtError.Location = New System.Drawing.Point(16, 14)
        Me.txtError.Margin = New System.Windows.Forms.Padding(4)
        Me.txtError.Multiline = True
        Me.txtError.Name = "txtError"
        Me.txtError.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtError.Size = New System.Drawing.Size(657, 377)
        Me.txtError.TabIndex = 0
        '
        'btnContinue
        '
        Me.btnContinue.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnContinue.Location = New System.Drawing.Point(467, 399)
        Me.btnContinue.Margin = New System.Windows.Forms.Padding(4)
        Me.btnContinue.Name = "btnContinue"
        Me.btnContinue.Size = New System.Drawing.Size(100, 28)
        Me.btnContinue.TabIndex = 1
        Me.btnContinue.Text = "Continue"
        Me.btnContinue.UseVisualStyleBackColor = True
        '
        'btnCancel2
        '
        Me.btnCancel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel2.Location = New System.Drawing.Point(575, 399)
        Me.btnCancel2.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel2.Name = "btnCancel2"
        Me.btnCancel2.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel2.TabIndex = 2
        Me.btnCancel2.Text = "Cancel"
        Me.btnCancel2.UseVisualStyleBackColor = True
        '
        'FrmTransferBeginingBalance_Error
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel2
        Me.ClientSize = New System.Drawing.Size(691, 433)
        Me.Controls.Add(Me.btnCancel2)
        Me.Controls.Add(Me.btnContinue)
        Me.Controls.Add(Me.txtError)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(699, 470)
        Me.Name = "FrmTransferBeginingBalance_Error"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Confirmation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtError As System.Windows.Forms.TextBox
    Friend WithEvents btnContinue As System.Windows.Forms.Button
    Friend WithEvents btnCancel2 As System.Windows.Forms.Button
End Class
