<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_SendEmail
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
        Me.txtSignByPosition = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSignByName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtpDue = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.llbSettingEmail = New System.Windows.Forms.LinkLabel
        Me.cboEmailList = New System.Windows.Forms.ComboBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblProgress = New System.Windows.Forms.Label
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.btnSend = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtAcuan = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'txtSignByPosition
        '
        Me.txtSignByPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSignByPosition.Location = New System.Drawing.Point(151, 105)
        Me.txtSignByPosition.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSignByPosition.MaxLength = 50
        Me.txtSignByPosition.Name = "txtSignByPosition"
        Me.txtSignByPosition.Size = New System.Drawing.Size(370, 22)
        Me.txtSignByPosition.TabIndex = 45
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 109)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(126, 17)
        Me.Label3.TabIndex = 48
        Me.Label3.Text = "Signed By Position"
        '
        'txtSignByName
        '
        Me.txtSignByName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSignByName.Location = New System.Drawing.Point(151, 76)
        Me.txtSignByName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSignByName.MaxLength = 50
        Me.txtSignByName.Name = "txtSignByName"
        Me.txtSignByName.Size = New System.Drawing.Size(370, 22)
        Me.txtSignByName.TabIndex = 44
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 79)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 17)
        Me.Label2.TabIndex = 47
        Me.Label2.Text = "Signed By Name"
        '
        'txtSubject
        '
        Me.txtSubject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubject.Location = New System.Drawing.Point(151, 45)
        Me.txtSubject.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSubject.MaxLength = 50
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(210, 22)
        Me.txtSubject.TabIndex = 43
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 48)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(55, 17)
        Me.Label8.TabIndex = 46
        Me.Label8.Text = "Subject"
        '
        'dtpDue
        '
        Me.dtpDue.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDue.CustomFormat = "dd MMM yyyy"
        Me.dtpDue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDue.Location = New System.Drawing.Point(368, 45)
        Me.dtpDue.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpDue.Name = "dtpDue"
        Me.dtpDue.Size = New System.Drawing.Size(153, 22)
        Me.dtpDue.TabIndex = 42
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 17)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "Email Account"
        '
        'llbSettingEmail
        '
        Me.llbSettingEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSettingEmail.AutoSize = True
        Me.llbSettingEmail.Location = New System.Drawing.Point(441, 18)
        Me.llbSettingEmail.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbSettingEmail.Name = "llbSettingEmail"
        Me.llbSettingEmail.Size = New System.Drawing.Size(59, 17)
        Me.llbSettingEmail.TabIndex = 54
        Me.llbSettingEmail.TabStop = True
        Me.llbSettingEmail.Text = "Settings"
        Me.llbSettingEmail.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'cboEmailList
        '
        Me.cboEmailList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboEmailList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmailList.FormattingEnabled = True
        Me.cboEmailList.Location = New System.Drawing.Point(151, 13)
        Me.cboEmailList.Margin = New System.Windows.Forms.Padding(4)
        Me.cboEmailList.Name = "cboEmailList"
        Me.cboEmailList.Size = New System.Drawing.Size(280, 24)
        Me.cboEmailList.TabIndex = 53
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(442, 132)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 55
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(12, 158)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(49, 17)
        Me.lblProgress.TabIndex = 52
        Me.lblProgress.Text = "Ready"
        '
        'prgAll
        '
        Me.prgAll.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgAll.Location = New System.Drawing.Point(12, 133)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(343, 23)
        Me.prgAll.TabIndex = 51
        '
        'btnSend
        '
        Me.btnSend.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSend.Location = New System.Drawing.Point(361, 131)
        Me.btnSend.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(75, 25)
        Me.btnSend.TabIndex = 56
        Me.btnSend.Text = "Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(145, 162)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.MaxLength = 50
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(62, 22)
        Me.TxtReadOnly.TabIndex = 57
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtAcuan
        '
        Me.TxtAcuan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAcuan.Location = New System.Drawing.Point(215, 162)
        Me.TxtAcuan.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAcuan.MaxLength = 50
        Me.TxtAcuan.Name = "TxtAcuan"
        Me.TxtAcuan.Size = New System.Drawing.Size(62, 22)
        Me.TxtAcuan.TabIndex = 58
        Me.TxtAcuan.TabStop = False
        Me.TxtAcuan.Visible = False
        '
        'FrmTrans_Attendance_SendEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(531, 186)
        Me.Controls.Add(Me.TxtAcuan)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.llbSettingEmail)
        Me.Controls.Add(Me.cboEmailList)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.prgAll)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSignByPosition)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtSignByName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSubject)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtpDue)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(539, 223)
        Me.Name = "FrmTrans_Attendance_SendEmail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Send Email: Attendance Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSignByPosition As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSignByName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpDue As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents llbSettingEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents cboEmailList As System.Windows.Forms.ComboBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtAcuan As System.Windows.Forms.TextBox
End Class
