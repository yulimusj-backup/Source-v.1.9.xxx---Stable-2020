<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterBackUpDB
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
        Me.btnDelete = New System.Windows.Forms.Button
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.chkReminderClosing = New System.Windows.Forms.CheckBox
        Me.chlShutdown = New System.Windows.Forms.CheckBox
        Me.btnFindLocation2 = New System.Windows.Forms.Button
        Me.btnFindLocation1 = New System.Windows.Forms.Button
        Me.TxtBackUpTime = New System.Windows.Forms.TextBox
        Me.btnFindUserID = New System.Windows.Forms.Button
        Me.TxtPrimaryUserID = New System.Windows.Forms.TextBox
        Me.TxtUserID = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboDayOfWeek = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboInterval = New System.Windows.Forms.ComboBox
        Me.TxtLocation2 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtLocation1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnNew = New System.Windows.Forms.Button
        Me.FrmDia = New System.Windows.Forms.FolderBrowserDialog
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(368, 297)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(296, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 368)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(612, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 45
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(296, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(296, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 393)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(612, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 46
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(296, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(319, 249)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(483, 297)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(253, 297)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(104, 23)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(24, 297)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.chkReminderClosing)
        Me.gboDetail.Controls.Add(Me.chlShutdown)
        Me.gboDetail.Controls.Add(Me.btnFindLocation2)
        Me.gboDetail.Controls.Add(Me.btnFindLocation1)
        Me.gboDetail.Controls.Add(Me.TxtBackUpTime)
        Me.gboDetail.Controls.Add(Me.btnFindUserID)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryUserID)
        Me.gboDetail.Controls.Add(Me.TxtUserID)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.cboDayOfWeek)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.cboInterval)
        Me.gboDetail.Controls.Add(Me.TxtLocation2)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.TxtLocation1)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Controls.Add(Me.TxtDesc)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Location = New System.Drawing.Point(12, 1)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(589, 288)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'chkReminderClosing
        '
        Me.chkReminderClosing.AutoSize = True
        Me.chkReminderClosing.Location = New System.Drawing.Point(169, 215)
        Me.chkReminderClosing.Name = "chkReminderClosing"
        Me.chkReminderClosing.Size = New System.Drawing.Size(146, 21)
        Me.chkReminderClosing.TabIndex = 43
        Me.chkReminderClosing.Text = "Remind me on exit"
        Me.chkReminderClosing.UseVisualStyleBackColor = True
        '
        'chlShutdown
        '
        Me.chlShutdown.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chlShutdown.AutoSize = True
        Me.chlShutdown.Location = New System.Drawing.Point(387, 184)
        Me.chlShutdown.Name = "chlShutdown"
        Me.chlShutdown.Size = New System.Drawing.Size(191, 21)
        Me.chlShutdown.TabIndex = 11
        Me.chlShutdown.Text = "Shutdown Computer After"
        Me.chlShutdown.UseVisualStyleBackColor = True
        '
        'btnFindLocation2
        '
        Me.btnFindLocation2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLocation2.Location = New System.Drawing.Point(526, 119)
        Me.btnFindLocation2.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindLocation2.Name = "btnFindLocation2"
        Me.btnFindLocation2.Size = New System.Drawing.Size(52, 25)
        Me.btnFindLocation2.TabIndex = 7
        Me.btnFindLocation2.TabStop = False
        Me.btnFindLocation2.Text = "...."
        Me.btnFindLocation2.UseVisualStyleBackColor = True
        '
        'btnFindLocation1
        '
        Me.btnFindLocation1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLocation1.Location = New System.Drawing.Point(526, 87)
        Me.btnFindLocation1.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindLocation1.Name = "btnFindLocation1"
        Me.btnFindLocation1.Size = New System.Drawing.Size(52, 25)
        Me.btnFindLocation1.TabIndex = 5
        Me.btnFindLocation1.TabStop = False
        Me.btnFindLocation1.Text = "...."
        Me.btnFindLocation1.UseVisualStyleBackColor = True
        '
        'TxtBackUpTime
        '
        Me.TxtBackUpTime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBackUpTime.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBackUpTime.Location = New System.Drawing.Point(169, 183)
        Me.TxtBackUpTime.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtBackUpTime.MaxLength = 50
        Me.TxtBackUpTime.Name = "TxtBackUpTime"
        Me.TxtBackUpTime.Size = New System.Drawing.Size(196, 22)
        Me.TxtBackUpTime.TabIndex = 10
        Me.TxtBackUpTime.Text = "00.00"
        Me.TxtBackUpTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnFindUserID
        '
        Me.btnFindUserID.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindUserID.Location = New System.Drawing.Point(526, 23)
        Me.btnFindUserID.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindUserID.Name = "btnFindUserID"
        Me.btnFindUserID.Size = New System.Drawing.Size(52, 25)
        Me.btnFindUserID.TabIndex = 2
        Me.btnFindUserID.TabStop = False
        Me.btnFindUserID.Text = "...."
        Me.btnFindUserID.UseVisualStyleBackColor = True
        '
        'TxtPrimaryUserID
        '
        Me.TxtPrimaryUserID.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryUserID.Location = New System.Drawing.Point(342, 24)
        Me.TxtPrimaryUserID.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryUserID.Name = "TxtPrimaryUserID"
        Me.TxtPrimaryUserID.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimaryUserID.TabIndex = 42
        Me.TxtPrimaryUserID.TabStop = False
        Me.TxtPrimaryUserID.Text = "-1"
        Me.TxtPrimaryUserID.Visible = False
        '
        'TxtUserID
        '
        Me.TxtUserID.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUserID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUserID.Location = New System.Drawing.Point(342, 23)
        Me.TxtUserID.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtUserID.MaxLength = 50
        Me.TxtUserID.Name = "TxtUserID"
        Me.TxtUserID.Size = New System.Drawing.Size(176, 22)
        Me.TxtUserID.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(279, 23)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 17)
        Me.Label7.TabIndex = 40
        Me.Label7.Text = "User ID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 186)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 17)
        Me.Label6.TabIndex = 39
        Me.Label6.Text = "Time"
        '
        'cboDayOfWeek
        '
        Me.cboDayOfWeek.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDayOfWeek.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDayOfWeek.FormattingEnabled = True
        Me.cboDayOfWeek.Items.AddRange(New Object() {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Everyday"})
        Me.cboDayOfWeek.Location = New System.Drawing.Point(387, 151)
        Me.cboDayOfWeek.Name = "cboDayOfWeek"
        Me.cboDayOfWeek.Size = New System.Drawing.Size(191, 24)
        Me.cboDayOfWeek.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 154)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(111, 17)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Back Up Interval"
        '
        'cboInterval
        '
        Me.cboInterval.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboInterval.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboInterval.FormattingEnabled = True
        Me.cboInterval.Items.AddRange(New Object() {"Daily", "Weekly"})
        Me.cboInterval.Location = New System.Drawing.Point(169, 151)
        Me.cboInterval.Name = "cboInterval"
        Me.cboInterval.Size = New System.Drawing.Size(196, 24)
        Me.cboInterval.TabIndex = 8
        '
        'TxtLocation2
        '
        Me.TxtLocation2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocation2.Location = New System.Drawing.Point(169, 119)
        Me.TxtLocation2.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtLocation2.MaxLength = 100
        Me.TxtLocation2.Name = "TxtLocation2"
        Me.TxtLocation2.Size = New System.Drawing.Size(349, 22)
        Me.TxtLocation2.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 122)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 17)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "Location 2"
        '
        'TxtLocation1
        '
        Me.TxtLocation1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocation1.Location = New System.Drawing.Point(169, 87)
        Me.TxtLocation1.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtLocation1.MaxLength = 100
        Me.TxtLocation1.Name = "TxtLocation1"
        Me.TxtLocation1.Size = New System.Drawing.Size(349, 22)
        Me.TxtLocation1.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 90)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 17)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Location 1"
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(164, 249)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(169, 55)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDesc.MaxLength = 100
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(409, 22)
        Me.TxtDesc.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 55)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 17)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Description"
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(169, 23)
        Me.TxtCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCode.MaxLength = 50
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(102, 22)
        Me.TxtCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Back Up Code"
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(139, 297)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'FrmMasterBackUpDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(612, 418)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(630, 430)
        Me.Name = "FrmMasterBackUpDB"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Back Up Plan"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtLocation2 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtLocation1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents FrmDia As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboDayOfWeek As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboInterval As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryUserID As System.Windows.Forms.TextBox
    Friend WithEvents TxtUserID As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnFindUserID As System.Windows.Forms.Button
    Friend WithEvents TxtBackUpTime As System.Windows.Forms.TextBox
    Friend WithEvents btnFindLocation2 As System.Windows.Forms.Button
    Friend WithEvents btnFindLocation1 As System.Windows.Forms.Button
    Friend WithEvents chlShutdown As System.Windows.Forms.CheckBox
    Friend WithEvents chkReminderClosing As System.Windows.Forms.CheckBox
End Class
