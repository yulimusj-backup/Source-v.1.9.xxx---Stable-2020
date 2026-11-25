<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch5
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboProcess = New System.Windows.Forms.ComboBox
        Me.cboDependency = New System.Windows.Forms.ComboBox
        Me.TxtDelay = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtBackColor = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Process"
        '
        'cboProcess
        '
        Me.cboProcess.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProcess.FormattingEnabled = True
        Me.cboProcess.Location = New System.Drawing.Point(110, 10)
        Me.cboProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboProcess.Name = "cboProcess"
        Me.cboProcess.Size = New System.Drawing.Size(174, 21)
        Me.cboProcess.TabIndex = 0
        '
        'cboDependency
        '
        Me.cboDependency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDependency.FormattingEnabled = True
        Me.cboDependency.Location = New System.Drawing.Point(110, 34)
        Me.cboDependency.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboDependency.Name = "cboDependency"
        Me.cboDependency.Size = New System.Drawing.Size(174, 21)
        Me.cboDependency.TabIndex = 1
        '
        'TxtDelay
        '
        Me.TxtDelay.Location = New System.Drawing.Point(110, 58)
        Me.TxtDelay.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtDelay.Name = "TxtDelay"
        Me.TxtDelay.Size = New System.Drawing.Size(174, 20)
        Me.TxtDelay.TabIndex = 2
        Me.TxtDelay.Text = "00.00"
        Me.TxtDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 37)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Dependency Type"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 61)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Delay"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(226, 81)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(166, 81)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtBackColor
        '
        Me.TxtBackColor.Location = New System.Drawing.Point(110, 80)
        Me.TxtBackColor.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtBackColor.Name = "TxtBackColor"
        Me.TxtBackColor.Size = New System.Drawing.Size(53, 20)
        Me.TxtBackColor.TabIndex = 6
        Me.TxtBackColor.TabStop = False
        Me.TxtBackColor.Visible = False
        '
        'FrmTrans_ProductionSchedule_NewSch5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(286, 107)
        Me.Controls.Add(Me.TxtBackColor)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtDelay)
        Me.Controls.Add(Me.cboDependency)
        Me.Controls.Add(Me.cboProcess)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(302, 145)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(302, 145)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch5"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Dependency By Process"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboProcess As System.Windows.Forms.ComboBox
    Friend WithEvents cboDependency As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDelay As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtBackColor As System.Windows.Forms.TextBox
End Class
