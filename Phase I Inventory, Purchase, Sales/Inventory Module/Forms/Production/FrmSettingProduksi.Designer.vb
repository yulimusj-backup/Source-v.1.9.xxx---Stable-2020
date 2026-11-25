<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSettingProduksi
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnSet_OT = New System.Windows.Forms.Button
        Me.pb_SchOT = New System.Windows.Forms.PictureBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnSet_SchDT = New System.Windows.Forms.Button
        Me.pb_SchDT = New System.Windows.Forms.PictureBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnSet_Offtime = New System.Windows.Forms.Button
        Me.pb_Offtime = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnSet_AllocatedWS = New System.Windows.Forms.Button
        Me.pb_AllocatedWS = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtEndDate = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtStartDate = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtRef = New System.Windows.Forms.TextBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.chkShowOffTime = New System.Windows.Forms.CheckBox
        Me.chkShowFreeTime = New System.Windows.Forms.CheckBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtFirmPlan = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.pb_SchOT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_SchDT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_Offtime, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_AllocatedWS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnSet_OT)
        Me.GroupBox1.Controls.Add(Me.pb_SchOT)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.btnSet_SchDT)
        Me.GroupBox1.Controls.Add(Me.pb_SchDT)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.btnSet_Offtime)
        Me.GroupBox1.Controls.Add(Me.pb_Offtime)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.btnSet_AllocatedWS)
        Me.GroupBox1.Controls.Add(Me.pb_AllocatedWS)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(572, 134)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Schedule Bar Color Settings"
        '
        'btnSet_OT
        '
        Me.btnSet_OT.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSet_OT.Location = New System.Drawing.Point(488, 104)
        Me.btnSet_OT.Name = "btnSet_OT"
        Me.btnSet_OT.Size = New System.Drawing.Size(75, 23)
        Me.btnSet_OT.TabIndex = 3
        Me.btnSet_OT.Text = "Set"
        Me.btnSet_OT.UseVisualStyleBackColor = True
        '
        'pb_SchOT
        '
        Me.pb_SchOT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb_SchOT.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_SchOT.Location = New System.Drawing.Point(160, 105)
        Me.pb_SchOT.Name = "pb_SchOT"
        Me.pb_SchOT.Size = New System.Drawing.Size(322, 22)
        Me.pb_SchOT.TabIndex = 11
        Me.pb_SchOT.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 17)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Scheduled Overtime"
        '
        'btnSet_SchDT
        '
        Me.btnSet_SchDT.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSet_SchDT.Location = New System.Drawing.Point(488, 76)
        Me.btnSet_SchDT.Name = "btnSet_SchDT"
        Me.btnSet_SchDT.Size = New System.Drawing.Size(75, 23)
        Me.btnSet_SchDT.TabIndex = 2
        Me.btnSet_SchDT.Text = "Set"
        Me.btnSet_SchDT.UseVisualStyleBackColor = True
        '
        'pb_SchDT
        '
        Me.pb_SchDT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb_SchDT.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_SchDT.Location = New System.Drawing.Point(160, 77)
        Me.pb_SchDT.Name = "pb_SchDT"
        Me.pb_SchDT.Size = New System.Drawing.Size(322, 22)
        Me.pb_SchDT.TabIndex = 8
        Me.pb_SchDT.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(140, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Scheduled Downtime"
        '
        'btnSet_Offtime
        '
        Me.btnSet_Offtime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSet_Offtime.Location = New System.Drawing.Point(488, 48)
        Me.btnSet_Offtime.Name = "btnSet_Offtime"
        Me.btnSet_Offtime.Size = New System.Drawing.Size(75, 23)
        Me.btnSet_Offtime.TabIndex = 1
        Me.btnSet_Offtime.Text = "Set"
        Me.btnSet_Offtime.UseVisualStyleBackColor = True
        '
        'pb_Offtime
        '
        Me.pb_Offtime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb_Offtime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_Offtime.Location = New System.Drawing.Point(160, 49)
        Me.pb_Offtime.Name = "pb_Offtime"
        Me.pb_Offtime.Size = New System.Drawing.Size(322, 22)
        Me.pb_Offtime.TabIndex = 5
        Me.pb_Offtime.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Off time"
        '
        'btnSet_AllocatedWS
        '
        Me.btnSet_AllocatedWS.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSet_AllocatedWS.Location = New System.Drawing.Point(488, 20)
        Me.btnSet_AllocatedWS.Name = "btnSet_AllocatedWS"
        Me.btnSet_AllocatedWS.Size = New System.Drawing.Size(75, 23)
        Me.btnSet_AllocatedWS.TabIndex = 0
        Me.btnSet_AllocatedWS.Text = "Set"
        Me.btnSet_AllocatedWS.UseVisualStyleBackColor = True
        '
        'pb_AllocatedWS
        '
        Me.pb_AllocatedWS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb_AllocatedWS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_AllocatedWS.Location = New System.Drawing.Point(160, 21)
        Me.pb_AllocatedWS.Name = "pb_AllocatedWS"
        Me.pb_AllocatedWS.Size = New System.Drawing.Size(322, 22)
        Me.pb_AllocatedWS.TabIndex = 1
        Me.pb_AllocatedWS.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(129, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Allocated Workshift"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(529, 342)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(448, 342)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.TxtFirmPlan)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.TxtEndDate)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.TxtStartDate)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 146)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(572, 143)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Schedule Horizon"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(408, 52)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(151, 17)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "days from current date"
        '
        'TxtEndDate
        '
        Me.TxtEndDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEndDate.Location = New System.Drawing.Point(160, 49)
        Me.TxtEndDate.Name = "TxtEndDate"
        Me.TxtEndDate.Size = New System.Drawing.Size(242, 22)
        Me.TxtEndDate.TabIndex = 1
        Me.TxtEndDate.Text = "0"
        Me.TxtEndDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(408, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(151, 17)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "days from current date"
        '
        'TxtStartDate
        '
        Me.TxtStartDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStartDate.Location = New System.Drawing.Point(160, 21)
        Me.TxtStartDate.Name = "TxtStartDate"
        Me.TxtStartDate.Size = New System.Drawing.Size(242, 22)
        Me.TxtStartDate.TabIndex = 0
        Me.TxtStartDate.Text = "0"
        Me.TxtStartDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 17)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "End Date"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 17)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Start Date"
        '
        'TxtRef
        '
        Me.TxtRef.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtRef.Location = New System.Drawing.Point(12, 343)
        Me.TxtRef.Name = "TxtRef"
        Me.TxtRef.Size = New System.Drawing.Size(100, 22)
        Me.TxtRef.TabIndex = 4
        Me.TxtRef.TabStop = False
        Me.TxtRef.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(592, 324)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(584, 295)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Global Settings"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.chkShowOffTime)
        Me.TabPage2.Controls.Add(Me.chkShowFreeTime)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(584, 234)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Personal Settings"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'chkShowOffTime
        '
        Me.chkShowOffTime.AutoSize = True
        Me.chkShowOffTime.Location = New System.Drawing.Point(6, 33)
        Me.chkShowOffTime.Name = "chkShowOffTime"
        Me.chkShowOffTime.Size = New System.Drawing.Size(122, 21)
        Me.chkShowOffTime.TabIndex = 1
        Me.chkShowOffTime.Text = "Show Off Time"
        Me.chkShowOffTime.UseVisualStyleBackColor = True
        '
        'chkShowFreeTime
        '
        Me.chkShowFreeTime.AutoSize = True
        Me.chkShowFreeTime.Location = New System.Drawing.Point(6, 6)
        Me.chkShowFreeTime.Name = "chkShowFreeTime"
        Me.chkShowFreeTime.Size = New System.Drawing.Size(132, 21)
        Me.chkShowFreeTime.TabIndex = 0
        Me.chkShowFreeTime.Text = "Show Free Time"
        Me.chkShowFreeTime.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(408, 80)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(151, 17)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "days from current date"
        '
        'TxtFirmPlan
        '
        Me.TxtFirmPlan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFirmPlan.Location = New System.Drawing.Point(160, 77)
        Me.TxtFirmPlan.Name = "TxtFirmPlan"
        Me.TxtFirmPlan.Size = New System.Drawing.Size(242, 22)
        Me.TxtFirmPlan.TabIndex = 12
        Me.TxtFirmPlan.Text = "0"
        Me.TxtFirmPlan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 80)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 17)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "Firm Plan"
        '
        'FrmSettingProduksi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(616, 377)
        Me.Controls.Add(Me.TxtRef)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimumSize = New System.Drawing.Size(437, 314)
        Me.Name = "FrmSettingProduksi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.pb_SchOT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_SchDT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_Offtime, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_AllocatedWS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents pb_AllocatedWS As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnSet_AllocatedWS As System.Windows.Forms.Button
    Friend WithEvents btnSet_OT As System.Windows.Forms.Button
    Friend WithEvents pb_SchOT As System.Windows.Forms.PictureBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnSet_SchDT As System.Windows.Forms.Button
    Friend WithEvents pb_SchDT As System.Windows.Forms.PictureBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSet_Offtime As System.Windows.Forms.Button
    Friend WithEvents pb_Offtime As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtEndDate As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtStartDate As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtRef As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents chkShowOffTime As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowFreeTime As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtFirmPlan As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
End Class
