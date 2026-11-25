<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkShift_AdvanceSettings
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
        Me.chkCalculateEarlyOTOnWeekend = New System.Windows.Forms.CheckBox
        Me.chkCalculateLateOnWeekend = New System.Windows.Forms.CheckBox
        Me.chkCalculateEarlyOutOnWeekend = New System.Windows.Forms.CheckBox
        Me.chkCalculateOTOnWeekend = New System.Windows.Forms.CheckBox
        Me.chkCalculateOTOnHoliday = New System.Windows.Forms.CheckBox
        Me.chkCalculateEarlyOutOnHoliday = New System.Windows.Forms.CheckBox
        Me.chkCalculateLateOnHoliday = New System.Windows.Forms.CheckBox
        Me.chkCalculateEarlyOTOnHoliday = New System.Windows.Forms.CheckBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'chkCalculateEarlyOTOnWeekend
        '
        Me.chkCalculateEarlyOTOnWeekend.AutoSize = True
        Me.chkCalculateEarlyOTOnWeekend.ForeColor = System.Drawing.Color.Red
        Me.chkCalculateEarlyOTOnWeekend.Location = New System.Drawing.Point(392, 119)
        Me.chkCalculateEarlyOTOnWeekend.Name = "chkCalculateEarlyOTOnWeekend"
        Me.chkCalculateEarlyOTOnWeekend.Size = New System.Drawing.Size(179, 17)
        Me.chkCalculateEarlyOTOnWeekend.TabIndex = 0
        Me.chkCalculateEarlyOTOnWeekend.Text = "Calculate Early OT on Weekend"
        Me.chkCalculateEarlyOTOnWeekend.UseVisualStyleBackColor = True
        Me.chkCalculateEarlyOTOnWeekend.Visible = False
        '
        'chkCalculateLateOnWeekend
        '
        Me.chkCalculateLateOnWeekend.AutoSize = True
        Me.chkCalculateLateOnWeekend.Location = New System.Drawing.Point(16, 13)
        Me.chkCalculateLateOnWeekend.Name = "chkCalculateLateOnWeekend"
        Me.chkCalculateLateOnWeekend.Size = New System.Drawing.Size(159, 17)
        Me.chkCalculateLateOnWeekend.TabIndex = 1
        Me.chkCalculateLateOnWeekend.Text = "Calculate Late on Weekend"
        Me.chkCalculateLateOnWeekend.UseVisualStyleBackColor = True
        '
        'chkCalculateEarlyOutOnWeekend
        '
        Me.chkCalculateEarlyOutOnWeekend.AutoSize = True
        Me.chkCalculateEarlyOutOnWeekend.Location = New System.Drawing.Point(16, 36)
        Me.chkCalculateEarlyOutOnWeekend.Name = "chkCalculateEarlyOutOnWeekend"
        Me.chkCalculateEarlyOutOnWeekend.Size = New System.Drawing.Size(181, 17)
        Me.chkCalculateEarlyOutOnWeekend.TabIndex = 2
        Me.chkCalculateEarlyOutOnWeekend.Text = "Calculate Early Out on Weekend"
        Me.chkCalculateEarlyOutOnWeekend.UseVisualStyleBackColor = True
        '
        'chkCalculateOTOnWeekend
        '
        Me.chkCalculateOTOnWeekend.AutoSize = True
        Me.chkCalculateOTOnWeekend.Checked = True
        Me.chkCalculateOTOnWeekend.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalculateOTOnWeekend.ForeColor = System.Drawing.Color.Red
        Me.chkCalculateOTOnWeekend.Location = New System.Drawing.Point(392, 188)
        Me.chkCalculateOTOnWeekend.Name = "chkCalculateOTOnWeekend"
        Me.chkCalculateOTOnWeekend.Size = New System.Drawing.Size(153, 17)
        Me.chkCalculateOTOnWeekend.TabIndex = 3
        Me.chkCalculateOTOnWeekend.Text = "Calculate OT on Weekend"
        Me.chkCalculateOTOnWeekend.UseVisualStyleBackColor = True
        Me.chkCalculateOTOnWeekend.Visible = False
        '
        'chkCalculateOTOnHoliday
        '
        Me.chkCalculateOTOnHoliday.AutoSize = True
        Me.chkCalculateOTOnHoliday.Checked = True
        Me.chkCalculateOTOnHoliday.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalculateOTOnHoliday.ForeColor = System.Drawing.Color.Red
        Me.chkCalculateOTOnHoliday.Location = New System.Drawing.Point(592, 188)
        Me.chkCalculateOTOnHoliday.Name = "chkCalculateOTOnHoliday"
        Me.chkCalculateOTOnHoliday.Size = New System.Drawing.Size(141, 17)
        Me.chkCalculateOTOnHoliday.TabIndex = 7
        Me.chkCalculateOTOnHoliday.Text = "Calculate OT on Holiday"
        Me.chkCalculateOTOnHoliday.UseVisualStyleBackColor = True
        Me.chkCalculateOTOnHoliday.Visible = False
        '
        'chkCalculateEarlyOutOnHoliday
        '
        Me.chkCalculateEarlyOutOnHoliday.AutoSize = True
        Me.chkCalculateEarlyOutOnHoliday.Location = New System.Drawing.Point(216, 36)
        Me.chkCalculateEarlyOutOnHoliday.Name = "chkCalculateEarlyOutOnHoliday"
        Me.chkCalculateEarlyOutOnHoliday.Size = New System.Drawing.Size(169, 17)
        Me.chkCalculateEarlyOutOnHoliday.TabIndex = 6
        Me.chkCalculateEarlyOutOnHoliday.Text = "Calculate Early Out on Holiday"
        Me.chkCalculateEarlyOutOnHoliday.UseVisualStyleBackColor = True
        '
        'chkCalculateLateOnHoliday
        '
        Me.chkCalculateLateOnHoliday.AutoSize = True
        Me.chkCalculateLateOnHoliday.Location = New System.Drawing.Point(216, 13)
        Me.chkCalculateLateOnHoliday.Name = "chkCalculateLateOnHoliday"
        Me.chkCalculateLateOnHoliday.Size = New System.Drawing.Size(147, 17)
        Me.chkCalculateLateOnHoliday.TabIndex = 5
        Me.chkCalculateLateOnHoliday.Text = "Calculate Late on Holiday"
        Me.chkCalculateLateOnHoliday.UseVisualStyleBackColor = True
        '
        'chkCalculateEarlyOTOnHoliday
        '
        Me.chkCalculateEarlyOTOnHoliday.AutoSize = True
        Me.chkCalculateEarlyOTOnHoliday.ForeColor = System.Drawing.Color.Red
        Me.chkCalculateEarlyOTOnHoliday.Location = New System.Drawing.Point(592, 119)
        Me.chkCalculateEarlyOTOnHoliday.Name = "chkCalculateEarlyOTOnHoliday"
        Me.chkCalculateEarlyOTOnHoliday.Size = New System.Drawing.Size(167, 17)
        Me.chkCalculateEarlyOTOnHoliday.TabIndex = 4
        Me.chkCalculateEarlyOTOnHoliday.Text = "Calculate Early OT on Holiday"
        Me.chkCalculateEarlyOTOnHoliday.UseVisualStyleBackColor = True
        Me.chkCalculateEarlyOTOnHoliday.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(328, 62)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 21)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(258, 62)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(64, 21)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmMasterWorkShift_AdvanceSettings
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(400, 90)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.chkCalculateOTOnHoliday)
        Me.Controls.Add(Me.chkCalculateEarlyOutOnHoliday)
        Me.Controls.Add(Me.chkCalculateLateOnHoliday)
        Me.Controls.Add(Me.chkCalculateEarlyOTOnHoliday)
        Me.Controls.Add(Me.chkCalculateOTOnWeekend)
        Me.Controls.Add(Me.chkCalculateEarlyOutOnWeekend)
        Me.Controls.Add(Me.chkCalculateLateOnWeekend)
        Me.Controls.Add(Me.chkCalculateEarlyOTOnWeekend)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(416, 129)
        Me.Name = "FrmMasterWorkShift_AdvanceSettings"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Shift Advance Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkCalculateEarlyOTOnWeekend As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateLateOnWeekend As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateEarlyOutOnWeekend As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateOTOnWeekend As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateOTOnHoliday As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateEarlyOutOnHoliday As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateLateOnHoliday As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateEarlyOTOnHoliday As System.Windows.Forms.CheckBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
