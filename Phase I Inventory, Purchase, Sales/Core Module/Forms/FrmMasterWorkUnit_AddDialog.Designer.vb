<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkUnit_AddDialog
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
        Me.BtnOk = New System.Windows.Forms.Button
        Me.CboPilih = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'BtnOk
        '
        Me.BtnOk.Location = New System.Drawing.Point(324, 11)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(51, 23)
        Me.BtnOk.TabIndex = 0
        Me.BtnOk.Text = "OK"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'CboPilih
        '
        Me.CboPilih.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPilih.FormattingEnabled = True
        Me.CboPilih.Items.AddRange(New Object() {"Create New Work Unit", "Using Existing Work Unit"})
        Me.CboPilih.Location = New System.Drawing.Point(121, 12)
        Me.CboPilih.Name = "CboPilih"
        Me.CboPilih.Size = New System.Drawing.Size(197, 21)
        Me.CboPilih.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Add To Work Unit"
        '
        'FrmMasterWorkUnit_AddDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(389, 47)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboPilih)
        Me.Controls.Add(Me.BtnOk)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(405, 85)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(405, 85)
        Me.Name = "FrmMasterWorkUnit_AddDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Unit Option"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnOk As System.Windows.Forms.Button
    Friend WithEvents CboPilih As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
