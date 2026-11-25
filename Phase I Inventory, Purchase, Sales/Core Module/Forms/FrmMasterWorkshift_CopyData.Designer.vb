<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkshift_CopyData
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
        Me.RboHolidayList = New System.Windows.Forms.RadioButton
        Me.RboWeeklyActivities = New System.Windows.Forms.RadioButton
        Me.RboBoth = New System.Windows.Forms.RadioButton
        Me.gbo_detail = New System.Windows.Forms.GroupBox
        Me.rboAttendanceSetting = New System.Windows.Forms.RadioButton
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.gbo_detail.SuspendLayout()
        Me.SuspendLayout()
        '
        'RboHolidayList
        '
        Me.RboHolidayList.AutoSize = True
        Me.RboHolidayList.Location = New System.Drawing.Point(6, 41)
        Me.RboHolidayList.Name = "RboHolidayList"
        Me.RboHolidayList.Size = New System.Drawing.Size(103, 17)
        Me.RboHolidayList.TabIndex = 1
        Me.RboHolidayList.TabStop = True
        Me.RboHolidayList.Text = "Holiday List Only"
        Me.RboHolidayList.UseVisualStyleBackColor = True
        '
        'RboWeeklyActivities
        '
        Me.RboWeeklyActivities.AutoSize = True
        Me.RboWeeklyActivities.Checked = True
        Me.RboWeeklyActivities.Location = New System.Drawing.Point(6, 21)
        Me.RboWeeklyActivities.Name = "RboWeeklyActivities"
        Me.RboWeeklyActivities.Size = New System.Drawing.Size(128, 17)
        Me.RboWeeklyActivities.TabIndex = 0
        Me.RboWeeklyActivities.TabStop = True
        Me.RboWeeklyActivities.Text = "Weekly Activites Only"
        Me.RboWeeklyActivities.UseVisualStyleBackColor = True
        '
        'RboBoth
        '
        Me.RboBoth.AutoSize = True
        Me.RboBoth.Location = New System.Drawing.Point(6, 87)
        Me.RboBoth.Name = "RboBoth"
        Me.RboBoth.Size = New System.Drawing.Size(89, 17)
        Me.RboBoth.TabIndex = 3
        Me.RboBoth.TabStop = True
        Me.RboBoth.Text = "Copy All Data"
        Me.RboBoth.UseVisualStyleBackColor = True
        '
        'gbo_detail
        '
        Me.gbo_detail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbo_detail.Controls.Add(Me.rboAttendanceSetting)
        Me.gbo_detail.Controls.Add(Me.RboHolidayList)
        Me.gbo_detail.Controls.Add(Me.RboWeeklyActivities)
        Me.gbo_detail.Controls.Add(Me.RboBoth)
        Me.gbo_detail.Location = New System.Drawing.Point(12, 12)
        Me.gbo_detail.Name = "gbo_detail"
        Me.gbo_detail.Size = New System.Drawing.Size(260, 129)
        Me.gbo_detail.TabIndex = 0
        Me.gbo_detail.TabStop = False
        '
        'rboAttendanceSetting
        '
        Me.rboAttendanceSetting.AutoSize = True
        Me.rboAttendanceSetting.Location = New System.Drawing.Point(6, 64)
        Me.rboAttendanceSetting.Name = "rboAttendanceSetting"
        Me.rboAttendanceSetting.Size = New System.Drawing.Size(140, 17)
        Me.rboAttendanceSetting.TabIndex = 2
        Me.rboAttendanceSetting.TabStop = True
        Me.rboAttendanceSetting.Text = "Attendance Setting Only"
        Me.rboAttendanceSetting.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(192, 147)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(98, 147)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmMasterWorkshift_CopyData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(292, 185)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.gbo_detail)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(300, 193)
        Me.Name = "FrmMasterWorkshift_CopyData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Workshift_Copy_Data"
        Me.gbo_detail.ResumeLayout(False)
        Me.gbo_detail.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RboHolidayList As System.Windows.Forms.RadioButton
    Friend WithEvents RboWeeklyActivities As System.Windows.Forms.RadioButton
    Friend WithEvents RboBoth As System.Windows.Forms.RadioButton
    Friend WithEvents gbo_detail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents rboAttendanceSetting As System.Windows.Forms.RadioButton
End Class
