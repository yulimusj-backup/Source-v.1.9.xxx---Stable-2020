<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiTagihan2
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
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpDue = New System.Windows.Forms.DateTimePicker
        Me.btnStart = New System.Windows.Forms.Button
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.lblProgress = New System.Windows.Forms.Label
        Me.chkDirectlySave = New System.Windows.Forms.CheckBox
        Me.chkDirectlyPrint = New System.Windows.Forms.CheckBox
        Me.LblCustomer = New System.Windows.Forms.Label
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtSignByName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSignByPosition = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.LblCust = New System.Windows.Forms.Label
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpEx = New System.Windows.Forms.DateTimePicker
        Me.txtTime = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.chkSendEmail = New System.Windows.Forms.CheckBox
        Me.cboEmailList = New System.Windows.Forms.ComboBox
        Me.llbSettingEmail = New System.Windows.Forms.LinkLabel
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Create collection letter for :"
        '
        'dtpDue
        '
        Me.dtpDue.CustomFormat = "dd MMM yyyy"
        Me.dtpDue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDue.Location = New System.Drawing.Point(227, 11)
        Me.dtpDue.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpDue.Name = "dtpDue"
        Me.dtpDue.Size = New System.Drawing.Size(203, 22)
        Me.dtpDue.TabIndex = 12
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(436, 11)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 25)
        Me.btnStart.TabIndex = 13
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'prgAll
        '
        Me.prgAll.Location = New System.Drawing.Point(15, 426)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(417, 23)
        Me.prgAll.TabIndex = 3
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(12, 455)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(49, 17)
        Me.lblProgress.TabIndex = 4
        Me.lblProgress.Text = "Ready"
        '
        'chkDirectlySave
        '
        Me.chkDirectlySave.AutoSize = True
        Me.chkDirectlySave.Location = New System.Drawing.Point(15, 364)
        Me.chkDirectlySave.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkDirectlySave.Name = "chkDirectlySave"
        Me.chkDirectlySave.Size = New System.Drawing.Size(190, 21)
        Me.chkDirectlySave.TabIndex = 19
        Me.chkDirectlySave.Text = "Directly save to database"
        Me.chkDirectlySave.UseVisualStyleBackColor = True
        '
        'chkDirectlyPrint
        '
        Me.chkDirectlyPrint.AutoSize = True
        Me.chkDirectlyPrint.Location = New System.Drawing.Point(211, 364)
        Me.chkDirectlyPrint.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkDirectlyPrint.Name = "chkDirectlyPrint"
        Me.chkDirectlyPrint.Size = New System.Drawing.Size(109, 21)
        Me.chkDirectlyPrint.TabIndex = 20
        Me.chkDirectlyPrint.Text = "Directly print"
        Me.chkDirectlyPrint.UseVisualStyleBackColor = True
        '
        'LblCustomer
        '
        Me.LblCustomer.AutoSize = True
        Me.LblCustomer.Location = New System.Drawing.Point(16, 102)
        Me.LblCustomer.Name = "LblCustomer"
        Me.LblCustomer.Size = New System.Drawing.Size(72, 17)
        Me.LblCustomer.TabIndex = 7
        Me.LblCustomer.Text = "Customer "
        '
        'txtSubject
        '
        Me.txtSubject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubject.Location = New System.Drawing.Point(149, 127)
        Me.txtSubject.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSubject.MaxLength = 50
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(365, 22)
        Me.txtSubject.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 130)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(55, 17)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "Subject"
        '
        'txtSignByName
        '
        Me.txtSignByName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSignByName.Location = New System.Drawing.Point(149, 158)
        Me.txtSignByName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSignByName.MaxLength = 50
        Me.txtSignByName.Name = "txtSignByName"
        Me.txtSignByName.Size = New System.Drawing.Size(365, 22)
        Me.txtSignByName.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 161)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 17)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Signed By Name"
        '
        'txtSignByPosition
        '
        Me.txtSignByPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSignByPosition.Location = New System.Drawing.Point(149, 187)
        Me.txtSignByPosition.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSignByPosition.MaxLength = 50
        Me.txtSignByPosition.Name = "txtSignByPosition"
        Me.txtSignByPosition.Size = New System.Drawing.Size(365, 22)
        Me.txtSignByPosition.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 191)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(126, 17)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Signed By Position"
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(149, 217)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(4)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(365, 139)
        Me.txtRemarks.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 222)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 17)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Remarks"
        '
        'LblCust
        '
        Me.LblCust.AutoSize = True
        Me.LblCust.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblCust.Location = New System.Drawing.Point(147, 102)
        Me.LblCust.Name = "LblCust"
        Me.LblCust.Size = New System.Drawing.Size(179, 17)
        Me.LblCust.TabIndex = 40
        Me.LblCust.Text = "The Customers not defined"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(89, 127)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(51, 22)
        Me.txtPrimary.TabIndex = 41
        Me.txtPrimary.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label5.Location = New System.Drawing.Point(16, 254)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 17)
        Me.Label5.TabIndex = 43
        Me.Label5.Text = "set numbering"
        Me.Label5.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(436, 425)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 23
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(16, 44)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(438, 17)
        Me.Label6.TabIndex = 45
        Me.Label6.Text = "Don't create collection letter if there is collection letter created after: "
        '
        'dtpEx
        '
        Me.dtpEx.CustomFormat = "dd MMM yyyy - HH:mm"
        Me.dtpEx.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEx.Location = New System.Drawing.Point(151, 65)
        Me.dtpEx.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpEx.Name = "dtpEx"
        Me.dtpEx.Size = New System.Drawing.Size(212, 22)
        Me.dtpEx.TabIndex = 0
        '
        'txtTime
        '
        Me.txtTime.Location = New System.Drawing.Point(418, 66)
        Me.txtTime.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTime.MaxLength = 5
        Me.txtTime.Name = "txtTime"
        Me.txtTime.Size = New System.Drawing.Size(69, 22)
        Me.txtTime.TabIndex = 1
        Me.txtTime.Text = "00:00"
        Me.txtTime.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(371, 70)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(39, 17)
        Me.Label7.TabIndex = 47
        Me.Label7.Text = "Time"
        Me.Label7.Visible = False
        '
        'chkSendEmail
        '
        Me.chkSendEmail.AutoSize = True
        Me.chkSendEmail.Location = New System.Drawing.Point(15, 394)
        Me.chkSendEmail.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkSendEmail.Name = "chkSendEmail"
        Me.chkSendEmail.Size = New System.Drawing.Size(101, 21)
        Me.chkSendEmail.TabIndex = 8
        Me.chkSendEmail.Text = "Send Email"
        Me.chkSendEmail.UseVisualStyleBackColor = True
        '
        'cboEmailList
        '
        Me.cboEmailList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmailList.FormattingEnabled = True
        Me.cboEmailList.Location = New System.Drawing.Point(151, 391)
        Me.cboEmailList.Margin = New System.Windows.Forms.Padding(4)
        Me.cboEmailList.Name = "cboEmailList"
        Me.cboEmailList.Size = New System.Drawing.Size(280, 24)
        Me.cboEmailList.TabIndex = 21
        '
        'llbSettingEmail
        '
        Me.llbSettingEmail.AutoSize = True
        Me.llbSettingEmail.Location = New System.Drawing.Point(441, 396)
        Me.llbSettingEmail.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbSettingEmail.Name = "llbSettingEmail"
        Me.llbSettingEmail.Size = New System.Drawing.Size(59, 17)
        Me.llbSettingEmail.TabIndex = 22
        Me.llbSettingEmail.TabStop = True
        Me.llbSettingEmail.Text = "Settings"
        Me.llbSettingEmail.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerReportsProgress = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(26, 302)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.MaxLength = 50
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(54, 22)
        Me.TxtReadOnly.TabIndex = 48
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'FrmTransaksiTagihan2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(537, 476)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.dtpEx)
        Me.Controls.Add(Me.txtTime)
        Me.Controls.Add(Me.llbSettingEmail)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cboEmailList)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.chkSendEmail)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPrimary)
        Me.Controls.Add(Me.LblCust)
        Me.Controls.Add(Me.txtRemarks)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSignByPosition)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtSignByName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSubject)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.LblCustomer)
        Me.Controls.Add(Me.chkDirectlyPrint)
        Me.Controls.Add(Me.chkDirectlySave)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.prgAll)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.dtpDue)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(545, 513)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(545, 513)
        Me.Name = "FrmTransaksiTagihan2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create Collection Letter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpDue As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents chkDirectlySave As System.Windows.Forms.CheckBox
    Friend WithEvents chkDirectlyPrint As System.Windows.Forms.CheckBox
    Friend WithEvents LblCustomer As System.Windows.Forms.Label
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtSignByName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSignByPosition As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblCust As System.Windows.Forms.Label
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpEx As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtTime As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents chkSendEmail As System.Windows.Forms.CheckBox
    Friend WithEvents cboEmailList As System.Windows.Forms.ComboBox
    Friend WithEvents llbSettingEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
End Class
