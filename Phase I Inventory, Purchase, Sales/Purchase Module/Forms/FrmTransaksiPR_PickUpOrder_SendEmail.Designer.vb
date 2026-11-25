<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPR_PickUpOrder_SendEmail
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
        Me.TxtAcuan = New System.Windows.Forms.TextBox
        Me.btnSend = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.llbSettingEmail = New System.Windows.Forms.LinkLabel
        Me.cboEmailList = New System.Windows.Forms.ComboBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblProgress = New System.Windows.Forms.Label
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSignByPosition = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSignByName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtpDue = New System.Windows.Forms.DateTimePicker
        Me.SuspendLayout()
        '
        'TxtAcuan
        '
        Me.TxtAcuan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAcuan.Location = New System.Drawing.Point(214, 157)
        Me.TxtAcuan.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAcuan.MaxLength = 50
        Me.TxtAcuan.Name = "TxtAcuan"
        Me.TxtAcuan.Size = New System.Drawing.Size(62, 22)
        Me.TxtAcuan.TabIndex = 74
        Me.TxtAcuan.TabStop = False
        Me.TxtAcuan.Visible = False
        '
        'btnSend
        '
        Me.btnSend.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSend.Location = New System.Drawing.Point(360, 126)
        Me.btnSend.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(75, 25)
        Me.btnSend.TabIndex = 72
        Me.btnSend.Text = "Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(144, 157)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.MaxLength = 50
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(62, 22)
        Me.TxtReadOnly.TabIndex = 73
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'llbSettingEmail
        '
        Me.llbSettingEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSettingEmail.AutoSize = True
        Me.llbSettingEmail.Location = New System.Drawing.Point(440, 13)
        Me.llbSettingEmail.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbSettingEmail.Name = "llbSettingEmail"
        Me.llbSettingEmail.Size = New System.Drawing.Size(59, 17)
        Me.llbSettingEmail.TabIndex = 70
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
        Me.cboEmailList.Location = New System.Drawing.Point(150, 8)
        Me.cboEmailList.Margin = New System.Windows.Forms.Padding(4)
        Me.cboEmailList.Name = "cboEmailList"
        Me.cboEmailList.Size = New System.Drawing.Size(280, 24)
        Me.cboEmailList.TabIndex = 69
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(441, 127)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 71
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(11, 153)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(49, 17)
        Me.lblProgress.TabIndex = 68
        Me.lblProgress.Text = "Ready"
        '
        'prgAll
        '
        Me.prgAll.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgAll.Location = New System.Drawing.Point(11, 128)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(343, 23)
        Me.prgAll.TabIndex = 67
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 17)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "Email Account"
        '
        'txtSignByPosition
        '
        Me.txtSignByPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSignByPosition.Location = New System.Drawing.Point(150, 100)
        Me.txtSignByPosition.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSignByPosition.MaxLength = 50
        Me.txtSignByPosition.Name = "txtSignByPosition"
        Me.txtSignByPosition.Size = New System.Drawing.Size(370, 22)
        Me.txtSignByPosition.TabIndex = 62
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 104)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(126, 17)
        Me.Label3.TabIndex = 65
        Me.Label3.Text = "Signed By Position"
        '
        'txtSignByName
        '
        Me.txtSignByName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSignByName.Location = New System.Drawing.Point(150, 71)
        Me.txtSignByName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSignByName.MaxLength = 50
        Me.txtSignByName.Name = "txtSignByName"
        Me.txtSignByName.Size = New System.Drawing.Size(370, 22)
        Me.txtSignByName.TabIndex = 61
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 74)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 17)
        Me.Label2.TabIndex = 64
        Me.Label2.Text = "Signed By Name"
        '
        'txtSubject
        '
        Me.txtSubject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubject.Location = New System.Drawing.Point(150, 40)
        Me.txtSubject.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSubject.MaxLength = 50
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(210, 22)
        Me.txtSubject.TabIndex = 60
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(11, 43)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(55, 17)
        Me.Label8.TabIndex = 63
        Me.Label8.Text = "Subject"
        '
        'dtpDue
        '
        Me.dtpDue.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDue.CustomFormat = "dd MMM yyyy"
        Me.dtpDue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDue.Location = New System.Drawing.Point(367, 40)
        Me.dtpDue.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpDue.Name = "dtpDue"
        Me.dtpDue.Size = New System.Drawing.Size(153, 22)
        Me.dtpDue.TabIndex = 59
        '
        'FrmTrans_Gaji_SendEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(531, 186)
        Me.Controls.Add(Me.TxtAcuan)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.TxtReadOnly)
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
        Me.Name = "FrmTrans_Gaji_SendEmail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Send Email: Salary Slip"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtAcuan As System.Windows.Forms.TextBox
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents llbSettingEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents cboEmailList As System.Windows.Forms.ComboBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSignByPosition As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSignByName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpDue As System.Windows.Forms.DateTimePicker
End Class
