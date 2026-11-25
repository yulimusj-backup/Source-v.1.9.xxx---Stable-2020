<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmEmailSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmEmailSettings))
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtAccountCode = New System.Windows.Forms.TextBox
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkOTOrder = New System.Windows.Forms.CheckBox
        Me.chkPickUpOrder = New System.Windows.Forms.CheckBox
        Me.chkSalarySlip = New System.Windows.Forms.CheckBox
        Me.chkAttendance = New System.Windows.Forms.CheckBox
        Me.chkCollectionLetter = New System.Windows.Forms.CheckBox
        Me.txtReadonly = New System.Windows.Forms.TextBox
        Me.txtPrimaryUser = New System.Windows.Forms.TextBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.btnFindUser = New System.Windows.Forms.Button
        Me.txtUserCode = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkUseSSL = New System.Windows.Forms.CheckBox
        Me.chkShowPassword = New System.Windows.Forms.CheckBox
        Me.txtSmtpPort = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtSmtpServer = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtUsername = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtReplyTo = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtEmailFrom = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_ColLetter = New System.Windows.Forms.TabPage
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtItemCCto = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.llbTestEmail1 = New System.Windows.Forms.LinkLabel
        Me.txtItemFileNamePrefix = New System.Windows.Forms.TextBox
        Me.txtItemSubject = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtItemMessage = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TabPage_Attendance = New System.Windows.Forms.TabPage
        Me.PictureBox_Att = New System.Windows.Forms.PictureBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtItemCCto_Att = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.llbTestEmail_Att = New System.Windows.Forms.LinkLabel
        Me.txtItemFileNamePrefix_Att = New System.Windows.Forms.TextBox
        Me.txtItemSubject_Att = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtItemMessage_Att = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.TabPage_SS = New System.Windows.Forms.TabPage
        Me.PictureBox_SS = New System.Windows.Forms.PictureBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtItemCCto_SS = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.llbTestEmail_SS = New System.Windows.Forms.LinkLabel
        Me.txtItemFileNamePrefix_SS = New System.Windows.Forms.TextBox
        Me.txtItemSubject_SS = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.txtItemMessage_SS = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.TabPage_PickUp = New System.Windows.Forms.TabPage
        Me.PictureBox_PickUp = New System.Windows.Forms.PictureBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.txtItemCCto_PickUp = New System.Windows.Forms.TextBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.llbTestEmail_PickUp = New System.Windows.Forms.LinkLabel
        Me.txtItemFileNamePrefix_PickUp = New System.Windows.Forms.TextBox
        Me.txtItemSubject_PickUp = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.txtItemMessage_PickUp = New System.Windows.Forms.TextBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.TabPage_OTOrder = New System.Windows.Forms.TabPage
        Me.PictureBox_OTOrder = New System.Windows.Forms.PictureBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtItemCCto_OTOrder = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.llbTestEmail_OTOrder = New System.Windows.Forms.LinkLabel
        Me.txtItemFileNamePrefix_OTOrder = New System.Windows.Forms.TextBox
        Me.txtItemSubject_OTOrder = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtItemMessage_OTOrder = New System.Windows.Forms.TextBox
        Me.Label34 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_ColLetter.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_Attendance.SuspendLayout()
        CType(Me.PictureBox_Att, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_SS.SuspendLayout()
        CType(Me.PictureBox_SS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_PickUp.SuspendLayout()
        CType(Me.PictureBox_PickUp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_OTOrder.SuspendLayout()
        CType(Me.PictureBox_OTOrder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Account Code"
        '
        'txtAccountCode
        '
        Me.txtAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAccountCode.Location = New System.Drawing.Point(98, 12)
        Me.txtAccountCode.MaxLength = 250
        Me.txtAccountCode.Name = "txtAccountCode"
        Me.txtAccountCode.Size = New System.Drawing.Size(209, 20)
        Me.txtAccountCode.TabIndex = 0
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(98, 36)
        Me.txtDescription.MaxLength = 250
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(546, 20)
        Me.txtDescription.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Description"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkOTOrder)
        Me.GroupBox1.Controls.Add(Me.chkPickUpOrder)
        Me.GroupBox1.Controls.Add(Me.chkSalarySlip)
        Me.GroupBox1.Controls.Add(Me.chkAttendance)
        Me.GroupBox1.Controls.Add(Me.chkCollectionLetter)
        Me.GroupBox1.Controls.Add(Me.txtReadonly)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryUser)
        Me.GroupBox1.Controls.Add(Me.txtPrimary)
        Me.GroupBox1.Controls.Add(Me.btnFindUser)
        Me.GroupBox1.Controls.Add(Me.txtUserCode)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.txtAccountCode)
        Me.GroupBox1.Controls.Add(Me.txtDescription)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(9, -1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(650, 83)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'chkOTOrder
        '
        Me.chkOTOrder.AutoSize = True
        Me.chkOTOrder.Location = New System.Drawing.Point(370, 60)
        Me.chkOTOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.chkOTOrder.Name = "chkOTOrder"
        Me.chkOTOrder.Size = New System.Drawing.Size(97, 17)
        Me.chkOTOrder.TabIndex = 10
        Me.chkOTOrder.Text = "Overtime Order"
        Me.chkOTOrder.UseVisualStyleBackColor = True
        '
        'chkPickUpOrder
        '
        Me.chkPickUpOrder.AutoSize = True
        Me.chkPickUpOrder.Location = New System.Drawing.Point(277, 60)
        Me.chkPickUpOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.chkPickUpOrder.Name = "chkPickUpOrder"
        Me.chkPickUpOrder.Size = New System.Drawing.Size(93, 17)
        Me.chkPickUpOrder.TabIndex = 7
        Me.chkPickUpOrder.Text = "Pick Up Order"
        Me.chkPickUpOrder.UseVisualStyleBackColor = True
        '
        'chkSalarySlip
        '
        Me.chkSalarySlip.AutoSize = True
        Me.chkSalarySlip.Location = New System.Drawing.Point(198, 60)
        Me.chkSalarySlip.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSalarySlip.Name = "chkSalarySlip"
        Me.chkSalarySlip.Size = New System.Drawing.Size(75, 17)
        Me.chkSalarySlip.TabIndex = 6
        Me.chkSalarySlip.Text = "Salary Slip"
        Me.chkSalarySlip.UseVisualStyleBackColor = True
        '
        'chkAttendance
        '
        Me.chkAttendance.AutoSize = True
        Me.chkAttendance.Location = New System.Drawing.Point(116, 60)
        Me.chkAttendance.Margin = New System.Windows.Forms.Padding(2)
        Me.chkAttendance.Name = "chkAttendance"
        Me.chkAttendance.Size = New System.Drawing.Size(81, 17)
        Me.chkAttendance.TabIndex = 5
        Me.chkAttendance.Text = "Attendance"
        Me.chkAttendance.UseVisualStyleBackColor = True
        '
        'chkCollectionLetter
        '
        Me.chkCollectionLetter.AutoSize = True
        Me.chkCollectionLetter.Location = New System.Drawing.Point(10, 60)
        Me.chkCollectionLetter.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCollectionLetter.Name = "chkCollectionLetter"
        Me.chkCollectionLetter.Size = New System.Drawing.Size(102, 17)
        Me.chkCollectionLetter.TabIndex = 4
        Me.chkCollectionLetter.Text = "Collection Letter"
        Me.chkCollectionLetter.UseVisualStyleBackColor = True
        '
        'txtReadonly
        '
        Me.txtReadonly.Location = New System.Drawing.Point(68, 33)
        Me.txtReadonly.Name = "txtReadonly"
        Me.txtReadonly.ReadOnly = True
        Me.txtReadonly.Size = New System.Drawing.Size(28, 20)
        Me.txtReadonly.TabIndex = 9
        Me.txtReadonly.Visible = False
        '
        'txtPrimaryUser
        '
        Me.txtPrimaryUser.Location = New System.Drawing.Point(346, 12)
        Me.txtPrimaryUser.Name = "txtPrimaryUser"
        Me.txtPrimaryUser.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryUser.TabIndex = 8
        Me.txtPrimaryUser.Text = "-1"
        Me.txtPrimaryUser.Visible = False
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(68, 12)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimary.TabIndex = 7
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'btnFindUser
        '
        Me.btnFindUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindUser.Location = New System.Drawing.Point(586, 12)
        Me.btnFindUser.Name = "btnFindUser"
        Me.btnFindUser.Size = New System.Drawing.Size(56, 21)
        Me.btnFindUser.TabIndex = 2
        Me.btnFindUser.Text = "Find"
        Me.btnFindUser.UseVisualStyleBackColor = True
        '
        'txtUserCode
        '
        Me.txtUserCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUserCode.Location = New System.Drawing.Point(380, 12)
        Me.txtUserCode.MaxLength = 250
        Me.txtUserCode.Name = "txtUserCode"
        Me.txtUserCode.Size = New System.Drawing.Size(199, 20)
        Me.txtUserCode.TabIndex = 1
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(313, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(57, 13)
        Me.Label18.TabIndex = 4
        Me.Label18.Text = "User Code"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.chkUseSSL)
        Me.GroupBox2.Controls.Add(Me.chkShowPassword)
        Me.GroupBox2.Controls.Add(Me.txtSmtpPort)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.txtSmtpServer)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtPassword)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.txtUsername)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.txtReplyTo)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtEmailFrom)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 88)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(650, 140)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Account Setting"
        '
        'chkUseSSL
        '
        Me.chkUseSSL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkUseSSL.AutoSize = True
        Me.chkUseSSL.Location = New System.Drawing.Point(539, 115)
        Me.chkUseSSL.Name = "chkUseSSL"
        Me.chkUseSSL.Size = New System.Drawing.Size(68, 17)
        Me.chkUseSSL.TabIndex = 7
        Me.chkUseSSL.Text = "Use SSL"
        Me.chkUseSSL.UseVisualStyleBackColor = True
        '
        'chkShowPassword
        '
        Me.chkShowPassword.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkShowPassword.AutoSize = True
        Me.chkShowPassword.Location = New System.Drawing.Point(537, 92)
        Me.chkShowPassword.Name = "chkShowPassword"
        Me.chkShowPassword.Size = New System.Drawing.Size(102, 17)
        Me.chkShowPassword.TabIndex = 4
        Me.chkShowPassword.Text = "Show Password"
        Me.chkShowPassword.UseVisualStyleBackColor = True
        '
        'txtSmtpPort
        '
        Me.txtSmtpPort.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpPort.Location = New System.Drawing.Point(452, 115)
        Me.txtSmtpPort.MaxLength = 5
        Me.txtSmtpPort.Name = "txtSmtpPort"
        Me.txtSmtpPort.Size = New System.Drawing.Size(79, 20)
        Me.txtSmtpPort.TabIndex = 6
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(388, 118)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(59, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "SMTP Port"
        '
        'txtSmtpServer
        '
        Me.txtSmtpServer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpServer.Location = New System.Drawing.Point(98, 115)
        Me.txtSmtpServer.MaxLength = 250
        Me.txtSmtpServer.Name = "txtSmtpServer"
        Me.txtSmtpServer.Size = New System.Drawing.Size(282, 20)
        Me.txtSmtpServer.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 118)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 13)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "SMTP Server"
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPassword.Location = New System.Drawing.Point(98, 90)
        Me.txtPassword.MaxLength = 250
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(434, 20)
        Me.txtPassword.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 93)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Password"
        '
        'txtUsername
        '
        Me.txtUsername.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUsername.Location = New System.Drawing.Point(98, 65)
        Me.txtUsername.MaxLength = 250
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(547, 20)
        Me.txtUsername.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Username"
        '
        'txtReplyTo
        '
        Me.txtReplyTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReplyTo.Location = New System.Drawing.Point(98, 40)
        Me.txtReplyTo.MaxLength = 250
        Me.txtReplyTo.Name = "txtReplyTo"
        Me.txtReplyTo.Size = New System.Drawing.Size(547, 20)
        Me.txtReplyTo.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Reply to"
        '
        'txtEmailFrom
        '
        Me.txtEmailFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEmailFrom.Location = New System.Drawing.Point(98, 16)
        Me.txtEmailFrom.MaxLength = 250
        Me.txtEmailFrom.Name = "txtEmailFrom"
        Me.txtEmailFrom.Size = New System.Drawing.Size(547, 20)
        Me.txtEmailFrom.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Email from"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(336, 440)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 26)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(236, 440)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(88, 26)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 506)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(668, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 18
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(326, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(326, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 528)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(668, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 17
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(326, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(326, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(286, 472)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(88, 26)
        Me.btnEdit.TabIndex = 4
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(190, 472)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(88, 26)
        Me.btnNew.TabIndex = 3
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(94, 472)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(88, 26)
        Me.btnFind.TabIndex = 2
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(478, 472)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 26)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(382, 472)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(88, 26)
        Me.btnDelete.TabIndex = 5
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage_ColLetter)
        Me.TabControl1.Controls.Add(Me.TabPage_Attendance)
        Me.TabControl1.Controls.Add(Me.TabPage_SS)
        Me.TabControl1.Controls.Add(Me.TabPage_PickUp)
        Me.TabControl1.Controls.Add(Me.TabPage_OTOrder)
        Me.TabControl1.Location = New System.Drawing.Point(9, 234)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(650, 201)
        Me.TabControl1.TabIndex = 24
        '
        'TabPage_ColLetter
        '
        Me.TabPage_ColLetter.BackColor = System.Drawing.Color.Transparent
        Me.TabPage_ColLetter.Controls.Add(Me.PictureBox1)
        Me.TabPage_ColLetter.Controls.Add(Me.Label13)
        Me.TabPage_ColLetter.Controls.Add(Me.txtItemCCto)
        Me.TabPage_ColLetter.Controls.Add(Me.Label12)
        Me.TabPage_ColLetter.Controls.Add(Me.llbTestEmail1)
        Me.TabPage_ColLetter.Controls.Add(Me.txtItemFileNamePrefix)
        Me.TabPage_ColLetter.Controls.Add(Me.txtItemSubject)
        Me.TabPage_ColLetter.Controls.Add(Me.Label11)
        Me.TabPage_ColLetter.Controls.Add(Me.Label9)
        Me.TabPage_ColLetter.Controls.Add(Me.txtItemMessage)
        Me.TabPage_ColLetter.Controls.Add(Me.Label10)
        Me.TabPage_ColLetter.Cursor = System.Windows.Forms.Cursors.Default
        Me.TabPage_ColLetter.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_ColLetter.Name = "TabPage_ColLetter"
        Me.TabPage_ColLetter.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_ColLetter.Size = New System.Drawing.Size(642, 175)
        Me.TabPage_ColLetter.TabIndex = 0
        Me.TabPage_ColLetter.Text = "Collection Letter"
        Me.TabPage_ColLetter.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(60, 50)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 20
        Me.PictureBox1.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(73, 60)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(0, 13)
        Me.Label13.TabIndex = 19
        '
        'txtItemCCto
        '
        Me.txtItemCCto.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemCCto.Location = New System.Drawing.Point(94, 7)
        Me.txtItemCCto.MaxLength = 250
        Me.txtItemCCto.Name = "txtItemCCto"
        Me.txtItemCCto.Size = New System.Drawing.Size(542, 20)
        Me.txtItemCCto.TabIndex = 0
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(4, 10)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(33, 13)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "CC to"
        '
        'llbTestEmail1
        '
        Me.llbTestEmail1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbTestEmail1.AutoSize = True
        Me.llbTestEmail1.Location = New System.Drawing.Point(3, 158)
        Me.llbTestEmail1.Name = "llbTestEmail1"
        Me.llbTestEmail1.Size = New System.Drawing.Size(56, 13)
        Me.llbTestEmail1.TabIndex = 4
        Me.llbTestEmail1.TabStop = True
        Me.llbTestEmail1.Text = "Test Email"
        Me.llbTestEmail1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'txtItemFileNamePrefix
        '
        Me.txtItemFileNamePrefix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemFileNamePrefix.Location = New System.Drawing.Point(94, 141)
        Me.txtItemFileNamePrefix.MaxLength = 250
        Me.txtItemFileNamePrefix.Name = "txtItemFileNamePrefix"
        Me.txtItemFileNamePrefix.Size = New System.Drawing.Size(542, 20)
        Me.txtItemFileNamePrefix.TabIndex = 3
        Me.txtItemFileNamePrefix.Visible = False
        '
        'txtItemSubject
        '
        Me.txtItemSubject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemSubject.Location = New System.Drawing.Point(94, 32)
        Me.txtItemSubject.MaxLength = 250
        Me.txtItemSubject.Name = "txtItemSubject"
        Me.txtItemSubject.Size = New System.Drawing.Size(542, 20)
        Me.txtItemSubject.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 144)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 13)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "Filename Prefix"
        Me.Label11.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 35)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Subject"
        '
        'txtItemMessage
        '
        Me.txtItemMessage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemMessage.Location = New System.Drawing.Point(94, 57)
        Me.txtItemMessage.MaxLength = 250
        Me.txtItemMessage.Multiline = True
        Me.txtItemMessage.Name = "txtItemMessage"
        Me.txtItemMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtItemMessage.Size = New System.Drawing.Size(542, 116)
        Me.txtItemMessage.TabIndex = 2
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(4, 60)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(50, 13)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Message"
        '
        'TabPage_Attendance
        '
        Me.TabPage_Attendance.BackColor = System.Drawing.Color.Transparent
        Me.TabPage_Attendance.Controls.Add(Me.PictureBox_Att)
        Me.TabPage_Attendance.Controls.Add(Me.Label19)
        Me.TabPage_Attendance.Controls.Add(Me.txtItemCCto_Att)
        Me.TabPage_Attendance.Controls.Add(Me.Label20)
        Me.TabPage_Attendance.Controls.Add(Me.llbTestEmail_Att)
        Me.TabPage_Attendance.Controls.Add(Me.txtItemFileNamePrefix_Att)
        Me.TabPage_Attendance.Controls.Add(Me.txtItemSubject_Att)
        Me.TabPage_Attendance.Controls.Add(Me.Label21)
        Me.TabPage_Attendance.Controls.Add(Me.Label22)
        Me.TabPage_Attendance.Controls.Add(Me.txtItemMessage_Att)
        Me.TabPage_Attendance.Controls.Add(Me.Label23)
        Me.TabPage_Attendance.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_Attendance.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_Attendance.Name = "TabPage_Attendance"
        Me.TabPage_Attendance.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_Attendance.Size = New System.Drawing.Size(642, 175)
        Me.TabPage_Attendance.TabIndex = 2
        Me.TabPage_Attendance.Text = "Attendance Data"
        Me.TabPage_Attendance.UseVisualStyleBackColor = True
        '
        'PictureBox_Att
        '
        Me.PictureBox_Att.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox_Att.Image = CType(resources.GetObject("PictureBox_Att.Image"), System.Drawing.Image)
        Me.PictureBox_Att.Location = New System.Drawing.Point(63, 49)
        Me.PictureBox_Att.Name = "PictureBox_Att"
        Me.PictureBox_Att.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox_Att.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox_Att.TabIndex = 31
        Me.PictureBox_Att.TabStop = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(76, 58)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(0, 13)
        Me.Label19.TabIndex = 30
        '
        'txtItemCCto_Att
        '
        Me.txtItemCCto_Att.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemCCto_Att.Location = New System.Drawing.Point(97, 6)
        Me.txtItemCCto_Att.MaxLength = 250
        Me.txtItemCCto_Att.Name = "txtItemCCto_Att"
        Me.txtItemCCto_Att.Size = New System.Drawing.Size(542, 20)
        Me.txtItemCCto_Att.TabIndex = 21
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(7, 8)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(33, 13)
        Me.Label20.TabIndex = 29
        Me.Label20.Text = "CC to"
        '
        'llbTestEmail_Att
        '
        Me.llbTestEmail_Att.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbTestEmail_Att.AutoSize = True
        Me.llbTestEmail_Att.Location = New System.Drawing.Point(6, 157)
        Me.llbTestEmail_Att.Name = "llbTestEmail_Att"
        Me.llbTestEmail_Att.Size = New System.Drawing.Size(56, 13)
        Me.llbTestEmail_Att.TabIndex = 25
        Me.llbTestEmail_Att.TabStop = True
        Me.llbTestEmail_Att.Text = "Test Email"
        Me.llbTestEmail_Att.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'txtItemFileNamePrefix_Att
        '
        Me.txtItemFileNamePrefix_Att.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemFileNamePrefix_Att.Location = New System.Drawing.Point(97, 139)
        Me.txtItemFileNamePrefix_Att.MaxLength = 250
        Me.txtItemFileNamePrefix_Att.Name = "txtItemFileNamePrefix_Att"
        Me.txtItemFileNamePrefix_Att.Size = New System.Drawing.Size(542, 20)
        Me.txtItemFileNamePrefix_Att.TabIndex = 24
        Me.txtItemFileNamePrefix_Att.Visible = False
        '
        'txtItemSubject_Att
        '
        Me.txtItemSubject_Att.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemSubject_Att.Location = New System.Drawing.Point(97, 30)
        Me.txtItemSubject_Att.MaxLength = 250
        Me.txtItemSubject_Att.Name = "txtItemSubject_Att"
        Me.txtItemSubject_Att.Size = New System.Drawing.Size(542, 20)
        Me.txtItemSubject_Att.TabIndex = 22
        '
        'Label21
        '
        Me.Label21.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(7, 142)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(78, 13)
        Me.Label21.TabIndex = 28
        Me.Label21.Text = "Filename Prefix"
        Me.Label21.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(7, 33)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(43, 13)
        Me.Label22.TabIndex = 26
        Me.Label22.Text = "Subject"
        '
        'txtItemMessage_Att
        '
        Me.txtItemMessage_Att.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemMessage_Att.Location = New System.Drawing.Point(97, 55)
        Me.txtItemMessage_Att.MaxLength = 250
        Me.txtItemMessage_Att.Multiline = True
        Me.txtItemMessage_Att.Name = "txtItemMessage_Att"
        Me.txtItemMessage_Att.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtItemMessage_Att.Size = New System.Drawing.Size(542, 116)
        Me.txtItemMessage_Att.TabIndex = 23
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(7, 58)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(50, 13)
        Me.Label23.TabIndex = 27
        Me.Label23.Text = "Message"
        '
        'TabPage_SS
        '
        Me.TabPage_SS.BackColor = System.Drawing.Color.Transparent
        Me.TabPage_SS.Controls.Add(Me.PictureBox_SS)
        Me.TabPage_SS.Controls.Add(Me.Label24)
        Me.TabPage_SS.Controls.Add(Me.txtItemCCto_SS)
        Me.TabPage_SS.Controls.Add(Me.Label25)
        Me.TabPage_SS.Controls.Add(Me.llbTestEmail_SS)
        Me.TabPage_SS.Controls.Add(Me.txtItemFileNamePrefix_SS)
        Me.TabPage_SS.Controls.Add(Me.txtItemSubject_SS)
        Me.TabPage_SS.Controls.Add(Me.Label26)
        Me.TabPage_SS.Controls.Add(Me.Label27)
        Me.TabPage_SS.Controls.Add(Me.txtItemMessage_SS)
        Me.TabPage_SS.Controls.Add(Me.Label28)
        Me.TabPage_SS.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_SS.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_SS.Name = "TabPage_SS"
        Me.TabPage_SS.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_SS.Size = New System.Drawing.Size(642, 175)
        Me.TabPage_SS.TabIndex = 3
        Me.TabPage_SS.Text = "Salary Slip"
        Me.TabPage_SS.UseVisualStyleBackColor = True
        '
        'PictureBox_SS
        '
        Me.PictureBox_SS.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox_SS.Image = CType(resources.GetObject("PictureBox_SS.Image"), System.Drawing.Image)
        Me.PictureBox_SS.Location = New System.Drawing.Point(63, 49)
        Me.PictureBox_SS.Name = "PictureBox_SS"
        Me.PictureBox_SS.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox_SS.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox_SS.TabIndex = 31
        Me.PictureBox_SS.TabStop = False
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(76, 58)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(0, 13)
        Me.Label24.TabIndex = 30
        '
        'txtItemCCto_SS
        '
        Me.txtItemCCto_SS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemCCto_SS.Location = New System.Drawing.Point(97, 6)
        Me.txtItemCCto_SS.MaxLength = 250
        Me.txtItemCCto_SS.Name = "txtItemCCto_SS"
        Me.txtItemCCto_SS.Size = New System.Drawing.Size(542, 20)
        Me.txtItemCCto_SS.TabIndex = 21
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(7, 8)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(33, 13)
        Me.Label25.TabIndex = 29
        Me.Label25.Text = "CC to"
        '
        'llbTestEmail_SS
        '
        Me.llbTestEmail_SS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbTestEmail_SS.AutoSize = True
        Me.llbTestEmail_SS.Location = New System.Drawing.Point(6, 157)
        Me.llbTestEmail_SS.Name = "llbTestEmail_SS"
        Me.llbTestEmail_SS.Size = New System.Drawing.Size(56, 13)
        Me.llbTestEmail_SS.TabIndex = 25
        Me.llbTestEmail_SS.TabStop = True
        Me.llbTestEmail_SS.Text = "Test Email"
        Me.llbTestEmail_SS.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'txtItemFileNamePrefix_SS
        '
        Me.txtItemFileNamePrefix_SS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemFileNamePrefix_SS.Location = New System.Drawing.Point(97, 139)
        Me.txtItemFileNamePrefix_SS.MaxLength = 250
        Me.txtItemFileNamePrefix_SS.Name = "txtItemFileNamePrefix_SS"
        Me.txtItemFileNamePrefix_SS.Size = New System.Drawing.Size(542, 20)
        Me.txtItemFileNamePrefix_SS.TabIndex = 24
        Me.txtItemFileNamePrefix_SS.Visible = False
        '
        'txtItemSubject_SS
        '
        Me.txtItemSubject_SS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemSubject_SS.Location = New System.Drawing.Point(97, 30)
        Me.txtItemSubject_SS.MaxLength = 250
        Me.txtItemSubject_SS.Name = "txtItemSubject_SS"
        Me.txtItemSubject_SS.Size = New System.Drawing.Size(542, 20)
        Me.txtItemSubject_SS.TabIndex = 22
        '
        'Label26
        '
        Me.Label26.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(7, 142)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(78, 13)
        Me.Label26.TabIndex = 28
        Me.Label26.Text = "Filename Prefix"
        Me.Label26.Visible = False
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(7, 33)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(43, 13)
        Me.Label27.TabIndex = 26
        Me.Label27.Text = "Subject"
        '
        'txtItemMessage_SS
        '
        Me.txtItemMessage_SS.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemMessage_SS.Location = New System.Drawing.Point(97, 55)
        Me.txtItemMessage_SS.MaxLength = 250
        Me.txtItemMessage_SS.Multiline = True
        Me.txtItemMessage_SS.Name = "txtItemMessage_SS"
        Me.txtItemMessage_SS.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtItemMessage_SS.Size = New System.Drawing.Size(542, 116)
        Me.txtItemMessage_SS.TabIndex = 23
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(7, 58)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(50, 13)
        Me.Label28.TabIndex = 27
        Me.Label28.Text = "Message"
        '
        'TabPage_PickUp
        '
        Me.TabPage_PickUp.BackColor = System.Drawing.Color.Transparent
        Me.TabPage_PickUp.Controls.Add(Me.PictureBox_PickUp)
        Me.TabPage_PickUp.Controls.Add(Me.Label29)
        Me.TabPage_PickUp.Controls.Add(Me.txtItemCCto_PickUp)
        Me.TabPage_PickUp.Controls.Add(Me.Label30)
        Me.TabPage_PickUp.Controls.Add(Me.llbTestEmail_PickUp)
        Me.TabPage_PickUp.Controls.Add(Me.txtItemFileNamePrefix_PickUp)
        Me.TabPage_PickUp.Controls.Add(Me.txtItemSubject_PickUp)
        Me.TabPage_PickUp.Controls.Add(Me.Label31)
        Me.TabPage_PickUp.Controls.Add(Me.Label32)
        Me.TabPage_PickUp.Controls.Add(Me.txtItemMessage_PickUp)
        Me.TabPage_PickUp.Controls.Add(Me.Label33)
        Me.TabPage_PickUp.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_PickUp.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_PickUp.Name = "TabPage_PickUp"
        Me.TabPage_PickUp.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_PickUp.Size = New System.Drawing.Size(642, 175)
        Me.TabPage_PickUp.TabIndex = 4
        Me.TabPage_PickUp.Text = "Pick Up Order"
        Me.TabPage_PickUp.UseVisualStyleBackColor = True
        '
        'PictureBox_PickUp
        '
        Me.PictureBox_PickUp.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox_PickUp.Image = CType(resources.GetObject("PictureBox_PickUp.Image"), System.Drawing.Image)
        Me.PictureBox_PickUp.Location = New System.Drawing.Point(63, 50)
        Me.PictureBox_PickUp.Name = "PictureBox_PickUp"
        Me.PictureBox_PickUp.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox_PickUp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox_PickUp.TabIndex = 42
        Me.PictureBox_PickUp.TabStop = False
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(76, 59)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(0, 13)
        Me.Label29.TabIndex = 41
        '
        'txtItemCCto_PickUp
        '
        Me.txtItemCCto_PickUp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemCCto_PickUp.Location = New System.Drawing.Point(97, 6)
        Me.txtItemCCto_PickUp.MaxLength = 250
        Me.txtItemCCto_PickUp.Name = "txtItemCCto_PickUp"
        Me.txtItemCCto_PickUp.Size = New System.Drawing.Size(542, 20)
        Me.txtItemCCto_PickUp.TabIndex = 32
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(7, 9)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(33, 13)
        Me.Label30.TabIndex = 40
        Me.Label30.Text = "CC to"
        '
        'llbTestEmail_PickUp
        '
        Me.llbTestEmail_PickUp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbTestEmail_PickUp.AutoSize = True
        Me.llbTestEmail_PickUp.Location = New System.Drawing.Point(6, 158)
        Me.llbTestEmail_PickUp.Name = "llbTestEmail_PickUp"
        Me.llbTestEmail_PickUp.Size = New System.Drawing.Size(56, 13)
        Me.llbTestEmail_PickUp.TabIndex = 36
        Me.llbTestEmail_PickUp.TabStop = True
        Me.llbTestEmail_PickUp.Text = "Test Email"
        Me.llbTestEmail_PickUp.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'txtItemFileNamePrefix_PickUp
        '
        Me.txtItemFileNamePrefix_PickUp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemFileNamePrefix_PickUp.Location = New System.Drawing.Point(97, 140)
        Me.txtItemFileNamePrefix_PickUp.MaxLength = 250
        Me.txtItemFileNamePrefix_PickUp.Name = "txtItemFileNamePrefix_PickUp"
        Me.txtItemFileNamePrefix_PickUp.Size = New System.Drawing.Size(542, 20)
        Me.txtItemFileNamePrefix_PickUp.TabIndex = 35
        Me.txtItemFileNamePrefix_PickUp.Visible = False
        '
        'txtItemSubject_PickUp
        '
        Me.txtItemSubject_PickUp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemSubject_PickUp.Location = New System.Drawing.Point(97, 31)
        Me.txtItemSubject_PickUp.MaxLength = 250
        Me.txtItemSubject_PickUp.Name = "txtItemSubject_PickUp"
        Me.txtItemSubject_PickUp.Size = New System.Drawing.Size(542, 20)
        Me.txtItemSubject_PickUp.TabIndex = 33
        '
        'Label31
        '
        Me.Label31.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(7, 143)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(78, 13)
        Me.Label31.TabIndex = 39
        Me.Label31.Text = "Filename Prefix"
        Me.Label31.Visible = False
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(7, 34)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(43, 13)
        Me.Label32.TabIndex = 37
        Me.Label32.Text = "Subject"
        '
        'txtItemMessage_PickUp
        '
        Me.txtItemMessage_PickUp.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemMessage_PickUp.Location = New System.Drawing.Point(97, 56)
        Me.txtItemMessage_PickUp.MaxLength = 250
        Me.txtItemMessage_PickUp.Multiline = True
        Me.txtItemMessage_PickUp.Name = "txtItemMessage_PickUp"
        Me.txtItemMessage_PickUp.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtItemMessage_PickUp.Size = New System.Drawing.Size(542, 116)
        Me.txtItemMessage_PickUp.TabIndex = 34
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(7, 59)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(50, 13)
        Me.Label33.TabIndex = 38
        Me.Label33.Text = "Message"
        '
        'TabPage_OTOrder
        '
        Me.TabPage_OTOrder.BackColor = System.Drawing.Color.Transparent
        Me.TabPage_OTOrder.Controls.Add(Me.PictureBox_OTOrder)
        Me.TabPage_OTOrder.Controls.Add(Me.Label14)
        Me.TabPage_OTOrder.Controls.Add(Me.txtItemCCto_OTOrder)
        Me.TabPage_OTOrder.Controls.Add(Me.Label15)
        Me.TabPage_OTOrder.Controls.Add(Me.llbTestEmail_OTOrder)
        Me.TabPage_OTOrder.Controls.Add(Me.txtItemFileNamePrefix_OTOrder)
        Me.TabPage_OTOrder.Controls.Add(Me.txtItemSubject_OTOrder)
        Me.TabPage_OTOrder.Controls.Add(Me.Label16)
        Me.TabPage_OTOrder.Controls.Add(Me.Label17)
        Me.TabPage_OTOrder.Controls.Add(Me.txtItemMessage_OTOrder)
        Me.TabPage_OTOrder.Controls.Add(Me.Label34)
        Me.TabPage_OTOrder.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_OTOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_OTOrder.Name = "TabPage_OTOrder"
        Me.TabPage_OTOrder.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_OTOrder.Size = New System.Drawing.Size(642, 175)
        Me.TabPage_OTOrder.TabIndex = 5
        Me.TabPage_OTOrder.Text = "Overtime Order"
        Me.TabPage_OTOrder.UseVisualStyleBackColor = True
        '
        'PictureBox_OTOrder
        '
        Me.PictureBox_OTOrder.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox_OTOrder.Image = CType(resources.GetObject("PictureBox_OTOrder.Image"), System.Drawing.Image)
        Me.PictureBox_OTOrder.Location = New System.Drawing.Point(63, 50)
        Me.PictureBox_OTOrder.Name = "PictureBox_OTOrder"
        Me.PictureBox_OTOrder.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox_OTOrder.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox_OTOrder.TabIndex = 53
        Me.PictureBox_OTOrder.TabStop = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(76, 59)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(0, 13)
        Me.Label14.TabIndex = 52
        '
        'txtItemCCto_OTOrder
        '
        Me.txtItemCCto_OTOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemCCto_OTOrder.Location = New System.Drawing.Point(97, 6)
        Me.txtItemCCto_OTOrder.MaxLength = 250
        Me.txtItemCCto_OTOrder.Name = "txtItemCCto_OTOrder"
        Me.txtItemCCto_OTOrder.Size = New System.Drawing.Size(542, 20)
        Me.txtItemCCto_OTOrder.TabIndex = 43
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(7, 9)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(33, 13)
        Me.Label15.TabIndex = 51
        Me.Label15.Text = "CC to"
        '
        'llbTestEmail_OTOrder
        '
        Me.llbTestEmail_OTOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbTestEmail_OTOrder.AutoSize = True
        Me.llbTestEmail_OTOrder.Location = New System.Drawing.Point(6, 158)
        Me.llbTestEmail_OTOrder.Name = "llbTestEmail_OTOrder"
        Me.llbTestEmail_OTOrder.Size = New System.Drawing.Size(56, 13)
        Me.llbTestEmail_OTOrder.TabIndex = 47
        Me.llbTestEmail_OTOrder.TabStop = True
        Me.llbTestEmail_OTOrder.Text = "Test Email"
        Me.llbTestEmail_OTOrder.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'txtItemFileNamePrefix_OTOrder
        '
        Me.txtItemFileNamePrefix_OTOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemFileNamePrefix_OTOrder.Location = New System.Drawing.Point(97, 140)
        Me.txtItemFileNamePrefix_OTOrder.MaxLength = 250
        Me.txtItemFileNamePrefix_OTOrder.Name = "txtItemFileNamePrefix_OTOrder"
        Me.txtItemFileNamePrefix_OTOrder.Size = New System.Drawing.Size(542, 20)
        Me.txtItemFileNamePrefix_OTOrder.TabIndex = 46
        Me.txtItemFileNamePrefix_OTOrder.Visible = False
        '
        'txtItemSubject_OTOrder
        '
        Me.txtItemSubject_OTOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemSubject_OTOrder.Location = New System.Drawing.Point(97, 31)
        Me.txtItemSubject_OTOrder.MaxLength = 250
        Me.txtItemSubject_OTOrder.Name = "txtItemSubject_OTOrder"
        Me.txtItemSubject_OTOrder.Size = New System.Drawing.Size(542, 20)
        Me.txtItemSubject_OTOrder.TabIndex = 44
        '
        'Label16
        '
        Me.Label16.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(7, 143)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(78, 13)
        Me.Label16.TabIndex = 50
        Me.Label16.Text = "Filename Prefix"
        Me.Label16.Visible = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(7, 34)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(43, 13)
        Me.Label17.TabIndex = 48
        Me.Label17.Text = "Subject"
        '
        'txtItemMessage_OTOrder
        '
        Me.txtItemMessage_OTOrder.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtItemMessage_OTOrder.Location = New System.Drawing.Point(97, 56)
        Me.txtItemMessage_OTOrder.MaxLength = 250
        Me.txtItemMessage_OTOrder.Multiline = True
        Me.txtItemMessage_OTOrder.Name = "txtItemMessage_OTOrder"
        Me.txtItemMessage_OTOrder.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtItemMessage_OTOrder.Size = New System.Drawing.Size(542, 116)
        Me.txtItemMessage_OTOrder.TabIndex = 45
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(7, 59)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(50, 13)
        Me.Label34.TabIndex = 49
        Me.Label34.Text = "Message"
        '
        'ToolTip1
        '
        Me.ToolTip1.ToolTipTitle = "Message Tips :"
        '
        'FrmEmailSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(668, 550)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.MinimumSize = New System.Drawing.Size(678, 550)
        Me.Name = "FrmEmailSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Email Account Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_ColLetter.ResumeLayout(False)
        Me.TabPage_ColLetter.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_Attendance.ResumeLayout(False)
        Me.TabPage_Attendance.PerformLayout()
        CType(Me.PictureBox_Att, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_SS.ResumeLayout(False)
        Me.TabPage_SS.PerformLayout()
        CType(Me.PictureBox_SS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_PickUp.ResumeLayout(False)
        Me.TabPage_PickUp.PerformLayout()
        CType(Me.PictureBox_PickUp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_OTOrder.ResumeLayout(False)
        Me.TabPage_OTOrder.PerformLayout()
        CType(Me.PictureBox_OTOrder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtReplyTo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtEmailFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSmtpServer As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSmtpPort As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_ColLetter As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtItemCCto As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents llbTestEmail1 As System.Windows.Forms.LinkLabel
    Friend WithEvents txtItemFileNamePrefix As System.Windows.Forms.TextBox
    Friend WithEvents txtItemSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtItemMessage As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnFindUser As System.Windows.Forms.Button
    Friend WithEvents txtUserCode As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryUser As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents chkShowPassword As System.Windows.Forms.CheckBox
    Friend WithEvents txtReadonly As System.Windows.Forms.TextBox
    Friend WithEvents chkUseSSL As System.Windows.Forms.CheckBox
    Friend WithEvents chkSalarySlip As System.Windows.Forms.CheckBox
    Friend WithEvents chkAttendance As System.Windows.Forms.CheckBox
    Friend WithEvents chkCollectionLetter As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage_Attendance As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox_Att As System.Windows.Forms.PictureBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtItemCCto_Att As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents llbTestEmail_Att As System.Windows.Forms.LinkLabel
    Friend WithEvents txtItemFileNamePrefix_Att As System.Windows.Forms.TextBox
    Friend WithEvents txtItemSubject_Att As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtItemMessage_Att As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TabPage_SS As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox_SS As System.Windows.Forms.PictureBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtItemCCto_SS As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents llbTestEmail_SS As System.Windows.Forms.LinkLabel
    Friend WithEvents txtItemFileNamePrefix_SS As System.Windows.Forms.TextBox
    Friend WithEvents txtItemSubject_SS As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents txtItemMessage_SS As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents chkPickUpOrder As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage_PickUp As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox_PickUp As System.Windows.Forms.PictureBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents txtItemCCto_PickUp As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents llbTestEmail_PickUp As System.Windows.Forms.LinkLabel
    Friend WithEvents txtItemFileNamePrefix_PickUp As System.Windows.Forms.TextBox
    Friend WithEvents txtItemSubject_PickUp As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents txtItemMessage_PickUp As System.Windows.Forms.TextBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents chkOTOrder As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage_OTOrder As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox_OTOrder As System.Windows.Forms.PictureBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtItemCCto_OTOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents llbTestEmail_OTOrder As System.Windows.Forms.LinkLabel
    Friend WithEvents txtItemFileNamePrefix_OTOrder As System.Windows.Forms.TextBox
    Friend WithEvents txtItemSubject_OTOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtItemMessage_OTOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
End Class
