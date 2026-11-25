<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_Actual_In_Report_ColSetting
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
        Me.chkShowAct = New System.Windows.Forms.CheckBox
        Me.ChkShowSch = New System.Windows.Forms.CheckBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkShowActBreak = New System.Windows.Forms.CheckBox
        Me.ChkShowSchBreak = New System.Windows.Forms.CheckBox
        Me.ChkShowShift = New System.Windows.Forms.CheckBox
        Me.ChkShowDept = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkShowAct
        '
        Me.chkShowAct.AutoSize = True
        Me.chkShowAct.Location = New System.Drawing.Point(8, 62)
        Me.chkShowAct.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowAct.Name = "chkShowAct"
        Me.chkShowAct.Size = New System.Drawing.Size(86, 17)
        Me.chkShowAct.TabIndex = 4
        Me.chkShowAct.Text = "Show Actual"
        Me.chkShowAct.UseVisualStyleBackColor = True
        '
        'ChkShowSch
        '
        Me.ChkShowSch.AutoSize = True
        Me.ChkShowSch.Location = New System.Drawing.Point(8, 41)
        Me.ChkShowSch.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowSch.Name = "ChkShowSch"
        Me.ChkShowSch.Size = New System.Drawing.Size(101, 17)
        Me.ChkShowSch.TabIndex = 2
        Me.ChkShowSch.Text = "Show Schedule"
        Me.ChkShowSch.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(212, 100)
        Me.btnOK.MaximumSize = New System.Drawing.Size(53, 21)
        Me.btnOK.MinimumSize = New System.Drawing.Size(53, 21)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(53, 21)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkShowDept)
        Me.GroupBox1.Controls.Add(Me.chkShowAct)
        Me.GroupBox1.Controls.Add(Me.ChkShowActBreak)
        Me.GroupBox1.Controls.Add(Me.ChkShowShift)
        Me.GroupBox1.Controls.Add(Me.ChkShowSch)
        Me.GroupBox1.Controls.Add(Me.ChkShowSchBreak)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(259, 90)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ChkShowActBreak
        '
        Me.ChkShowActBreak.AutoSize = True
        Me.ChkShowActBreak.Location = New System.Drawing.Point(125, 62)
        Me.ChkShowActBreak.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowActBreak.Name = "ChkShowActBreak"
        Me.ChkShowActBreak.Size = New System.Drawing.Size(117, 17)
        Me.ChkShowActBreak.TabIndex = 5
        Me.ChkShowActBreak.Text = "Show Actual Break"
        Me.ChkShowActBreak.UseVisualStyleBackColor = True
        '
        'ChkShowSchBreak
        '
        Me.ChkShowSchBreak.AutoSize = True
        Me.ChkShowSchBreak.Location = New System.Drawing.Point(125, 41)
        Me.ChkShowSchBreak.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowSchBreak.Name = "ChkShowSchBreak"
        Me.ChkShowSchBreak.Size = New System.Drawing.Size(132, 17)
        Me.ChkShowSchBreak.TabIndex = 3
        Me.ChkShowSchBreak.Text = "Show Schedule Break"
        Me.ChkShowSchBreak.UseVisualStyleBackColor = True
        '
        'ChkShowShift
        '
        Me.ChkShowShift.AutoSize = True
        Me.ChkShowShift.Location = New System.Drawing.Point(125, 18)
        Me.ChkShowShift.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowShift.Name = "ChkShowShift"
        Me.ChkShowShift.Size = New System.Drawing.Size(77, 17)
        Me.ChkShowShift.TabIndex = 1
        Me.ChkShowShift.Text = "Show Shift"
        Me.ChkShowShift.UseVisualStyleBackColor = True
        '
        'ChkShowDept
        '
        Me.ChkShowDept.AutoSize = True
        Me.ChkShowDept.Location = New System.Drawing.Point(8, 18)
        Me.ChkShowDept.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowDept.Name = "ChkShowDept"
        Me.ChkShowDept.Size = New System.Drawing.Size(111, 17)
        Me.ChkShowDept.TabIndex = 0
        Me.ChkShowDept.Text = "Show Department"
        Me.ChkShowDept.UseVisualStyleBackColor = True
        '
        'FrmTrans_Attendance_Actual_In_Report_ColSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(272, 133)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "FrmTrans_Attendance_Actual_In_Report_ColSetting"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ChkShowSch As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowAct As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkShowSchBreak As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowActBreak As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowDept As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowShift As System.Windows.Forms.CheckBox
End Class
