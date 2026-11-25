<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Reorder_ColSetting
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
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ChkWorkHour = New System.Windows.Forms.CheckBox
        Me.ChkWONo = New System.Windows.Forms.CheckBox
        Me.ChkUnit = New System.Windows.Forms.CheckBox
        Me.ChkCodeNo = New System.Windows.Forms.CheckBox
        Me.ChkQty = New System.Windows.Forms.CheckBox
        Me.ChkDesc = New System.Windows.Forms.CheckBox
        Me.ChkRate = New System.Windows.Forms.CheckBox
        Me.ChkDayStart = New System.Windows.Forms.CheckBox
        Me.ChkProgress = New System.Windows.Forms.CheckBox
        Me.ChkDayFinish = New System.Windows.Forms.CheckBox
        Me.ChkProcess = New System.Windows.Forms.CheckBox
        Me.ChkMO = New System.Windows.Forms.CheckBox
        Me.ChkJO = New System.Windows.Forms.CheckBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BtnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ChkWorkHour
        '
        Me.ChkWorkHour.AutoSize = True
        Me.ChkWorkHour.Location = New System.Drawing.Point(111, 104)
        Me.ChkWorkHour.Name = "ChkWorkHour"
        Me.ChkWorkHour.Size = New System.Drawing.Size(78, 17)
        Me.ChkWorkHour.TabIndex = 9
        Me.ChkWorkHour.Text = "Work Hour"
        Me.ChkWorkHour.UseVisualStyleBackColor = True
        '
        'ChkWONo
        '
        Me.ChkWONo.AutoSize = True
        Me.ChkWONo.Location = New System.Drawing.Point(10, 58)
        Me.ChkWONo.Name = "ChkWONo"
        Me.ChkWONo.Size = New System.Drawing.Size(62, 17)
        Me.ChkWONo.TabIndex = 2
        Me.ChkWONo.Text = "WO No"
        Me.ChkWONo.UseVisualStyleBackColor = True
        '
        'ChkUnit
        '
        Me.ChkUnit.AutoSize = True
        Me.ChkUnit.Location = New System.Drawing.Point(111, 81)
        Me.ChkUnit.Name = "ChkUnit"
        Me.ChkUnit.Size = New System.Drawing.Size(45, 17)
        Me.ChkUnit.TabIndex = 8
        Me.ChkUnit.Text = "Unit"
        Me.ChkUnit.UseVisualStyleBackColor = True
        '
        'ChkCodeNo
        '
        Me.ChkCodeNo.AutoSize = True
        Me.ChkCodeNo.Location = New System.Drawing.Point(10, 104)
        Me.ChkCodeNo.Name = "ChkCodeNo"
        Me.ChkCodeNo.Size = New System.Drawing.Size(68, 17)
        Me.ChkCodeNo.TabIndex = 4
        Me.ChkCodeNo.Text = "Code No"
        Me.ChkCodeNo.UseVisualStyleBackColor = True
        '
        'ChkQty
        '
        Me.ChkQty.AutoSize = True
        Me.ChkQty.Location = New System.Drawing.Point(111, 58)
        Me.ChkQty.Name = "ChkQty"
        Me.ChkQty.Size = New System.Drawing.Size(42, 17)
        Me.ChkQty.TabIndex = 7
        Me.ChkQty.Text = "Qty"
        Me.ChkQty.UseVisualStyleBackColor = True
        '
        'ChkDesc
        '
        Me.ChkDesc.AutoSize = True
        Me.ChkDesc.Location = New System.Drawing.Point(111, 11)
        Me.ChkDesc.Name = "ChkDesc"
        Me.ChkDesc.Size = New System.Drawing.Size(79, 17)
        Me.ChkDesc.TabIndex = 5
        Me.ChkDesc.Text = "Description"
        Me.ChkDesc.UseVisualStyleBackColor = True
        '
        'ChkRate
        '
        Me.ChkRate.AutoSize = True
        Me.ChkRate.Location = New System.Drawing.Point(111, 34)
        Me.ChkRate.Name = "ChkRate"
        Me.ChkRate.Size = New System.Drawing.Size(49, 17)
        Me.ChkRate.TabIndex = 6
        Me.ChkRate.Text = "Rate"
        Me.ChkRate.UseVisualStyleBackColor = True
        '
        'ChkDayStart
        '
        Me.ChkDayStart.AutoSize = True
        Me.ChkDayStart.Location = New System.Drawing.Point(225, 11)
        Me.ChkDayStart.Name = "ChkDayStart"
        Me.ChkDayStart.Size = New System.Drawing.Size(114, 17)
        Me.ChkDayStart.TabIndex = 10
        Me.ChkDayStart.Text = "Day of Week Start"
        Me.ChkDayStart.UseVisualStyleBackColor = True
        '
        'ChkProgress
        '
        Me.ChkProgress.AutoSize = True
        Me.ChkProgress.Location = New System.Drawing.Point(225, 58)
        Me.ChkProgress.Name = "ChkProgress"
        Me.ChkProgress.Size = New System.Drawing.Size(67, 17)
        Me.ChkProgress.TabIndex = 12
        Me.ChkProgress.Text = "Progress"
        Me.ChkProgress.UseVisualStyleBackColor = True
        '
        'ChkDayFinish
        '
        Me.ChkDayFinish.AutoSize = True
        Me.ChkDayFinish.Location = New System.Drawing.Point(225, 34)
        Me.ChkDayFinish.Name = "ChkDayFinish"
        Me.ChkDayFinish.Size = New System.Drawing.Size(119, 17)
        Me.ChkDayFinish.TabIndex = 11
        Me.ChkDayFinish.Text = "Day of Week Finish"
        Me.ChkDayFinish.UseVisualStyleBackColor = True
        '
        'ChkProcess
        '
        Me.ChkProcess.AutoSize = True
        Me.ChkProcess.Location = New System.Drawing.Point(10, 81)
        Me.ChkProcess.Name = "ChkProcess"
        Me.ChkProcess.Size = New System.Drawing.Size(64, 17)
        Me.ChkProcess.TabIndex = 3
        Me.ChkProcess.Text = "Process"
        Me.ChkProcess.UseVisualStyleBackColor = True
        '
        'ChkMO
        '
        Me.ChkMO.AutoSize = True
        Me.ChkMO.Location = New System.Drawing.Point(10, 11)
        Me.ChkMO.Name = "ChkMO"
        Me.ChkMO.Size = New System.Drawing.Size(60, 17)
        Me.ChkMO.TabIndex = 0
        Me.ChkMO.Text = "MO No"
        Me.ChkMO.UseVisualStyleBackColor = True
        '
        'ChkJO
        '
        Me.ChkJO.AutoSize = True
        Me.ChkJO.Location = New System.Drawing.Point(10, 34)
        Me.ChkJO.Name = "ChkJO"
        Me.ChkJO.Size = New System.Drawing.Size(56, 17)
        Me.ChkJO.TabIndex = 1
        Me.ChkJO.Text = "JO No"
        Me.ChkJO.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(78, 129)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(66, 22)
        Me.BtnCancel.TabIndex = 14
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(6, 129)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(66, 22)
        Me.BtnOK.TabIndex = 13
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule_Reorder_ColSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(341, 162)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.ChkMO)
        Me.Controls.Add(Me.ChkJO)
        Me.Controls.Add(Me.ChkWorkHour)
        Me.Controls.Add(Me.ChkWONo)
        Me.Controls.Add(Me.ChkUnit)
        Me.Controls.Add(Me.ChkCodeNo)
        Me.Controls.Add(Me.ChkQty)
        Me.Controls.Add(Me.ChkDesc)
        Me.Controls.Add(Me.ChkRate)
        Me.Controls.Add(Me.ChkDayStart)
        Me.Controls.Add(Me.ChkProgress)
        Me.Controls.Add(Me.ChkDayFinish)
        Me.Controls.Add(Me.ChkProcess)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_ProductionSchedule_Reorder_ColSetting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Setting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkWorkHour As System.Windows.Forms.CheckBox
    Friend WithEvents ChkWONo As System.Windows.Forms.CheckBox
    Friend WithEvents ChkUnit As System.Windows.Forms.CheckBox
    Friend WithEvents ChkCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents ChkQty As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDesc As System.Windows.Forms.CheckBox
    Friend WithEvents ChkRate As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDayStart As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProgress As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDayFinish As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProcess As System.Windows.Forms.CheckBox
    Friend WithEvents ChkMO As System.Windows.Forms.CheckBox
    Friend WithEvents ChkJO As System.Windows.Forms.CheckBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BtnOK As System.Windows.Forms.Button
End Class
