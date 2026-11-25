<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterAkun_Detail
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.cboCurrency = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.LblIDR = New System.Windows.Forms.Label
        Me.TxtBeginningBalanceOriginalCur = New System.Windows.Forms.TextBox
        Me.LblOriginalCur = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboAccountType = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryAccount = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.TxtKodeAkun1 = New System.Windows.Forms.TextBox
        Me.chkHeaderAccount = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtBeginningBalance = New System.Windows.Forms.TextBox
        Me.chkActiveAccount = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtAccountNo = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.cboCurrency)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.LblIDR)
        Me.gboDetail.Controls.Add(Me.TxtBeginningBalanceOriginalCur)
        Me.gboDetail.Controls.Add(Me.LblOriginalCur)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.cboAccountType)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryAccount)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.TxtKodeAkun1)
        Me.gboDetail.Controls.Add(Me.chkHeaderAccount)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.TxtBeginningBalance)
        Me.gboDetail.Controls.Add(Me.chkActiveAccount)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtDesc)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtAccountNo)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(7, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(535, 259)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'cboCurrency
        '
        Me.cboCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrency.FormattingEnabled = True
        Me.cboCurrency.Location = New System.Drawing.Point(112, 98)
        Me.cboCurrency.Name = "cboCurrency"
        Me.cboCurrency.Size = New System.Drawing.Size(414, 21)
        Me.cboCurrency.TabIndex = 5
        '
        'TxtPrimaryCurrency
        '
        Me.TxtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCurrency.Location = New System.Drawing.Point(72, 98)
        Me.TxtPrimaryCurrency.MaxLength = 190
        Me.TxtPrimaryCurrency.Name = "TxtPrimaryCurrency"
        Me.TxtPrimaryCurrency.Size = New System.Drawing.Size(34, 20)
        Me.TxtPrimaryCurrency.TabIndex = 36
        Me.TxtPrimaryCurrency.TabStop = False
        Me.TxtPrimaryCurrency.Text = "-1"
        Me.TxtPrimaryCurrency.Visible = False
        '
        'LblIDR
        '
        Me.LblIDR.AutoSize = True
        Me.LblIDR.Location = New System.Drawing.Point(111, 127)
        Me.LblIDR.Name = "LblIDR"
        Me.LblIDR.Size = New System.Drawing.Size(26, 13)
        Me.LblIDR.TabIndex = 35
        Me.LblIDR.Text = "IDR"
        '
        'TxtBeginningBalanceOriginalCur
        '
        Me.TxtBeginningBalanceOriginalCur.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBeginningBalanceOriginalCur.Location = New System.Drawing.Point(374, 124)
        Me.TxtBeginningBalanceOriginalCur.MaxLength = 18
        Me.TxtBeginningBalanceOriginalCur.Name = "TxtBeginningBalanceOriginalCur"
        Me.TxtBeginningBalanceOriginalCur.Size = New System.Drawing.Size(152, 20)
        Me.TxtBeginningBalanceOriginalCur.TabIndex = 7
        Me.TxtBeginningBalanceOriginalCur.Text = "0"
        Me.TxtBeginningBalanceOriginalCur.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblOriginalCur
        '
        Me.LblOriginalCur.AutoSize = True
        Me.LblOriginalCur.Location = New System.Drawing.Point(267, 127)
        Me.LblOriginalCur.Name = "LblOriginalCur"
        Me.LblOriginalCur.Size = New System.Drawing.Size(99, 13)
        Me.LblOriginalCur.TabIndex = 33
        Me.LblOriginalCur.Text = "In Original Currency"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 101)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Currency"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 74)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 13)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Account Type"
        '
        'cboAccountType
        '
        Me.cboAccountType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountType.FormattingEnabled = True
        Me.cboAccountType.Location = New System.Drawing.Point(112, 71)
        Me.cboAccountType.Name = "cboAccountType"
        Me.cboAccountType.Size = New System.Drawing.Size(152, 21)
        Me.cboAccountType.TabIndex = 2
        '
        'TxtPrimaryAccount
        '
        Me.TxtPrimaryAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryAccount.Location = New System.Drawing.Point(72, 19)
        Me.TxtPrimaryAccount.MaxLength = 190
        Me.TxtPrimaryAccount.Name = "TxtPrimaryAccount"
        Me.TxtPrimaryAccount.Size = New System.Drawing.Size(34, 20)
        Me.TxtPrimaryAccount.TabIndex = 27
        Me.TxtPrimaryAccount.TabStop = False
        Me.TxtPrimaryAccount.Text = "-1"
        Me.TxtPrimaryAccount.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 153)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "Account Nature"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(112, 150)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(414, 72)
        Me.TxtRemarks.TabIndex = 8
        '
        'TxtKodeAkun1
        '
        Me.TxtKodeAkun1.Location = New System.Drawing.Point(112, 19)
        Me.TxtKodeAkun1.Name = "TxtKodeAkun1"
        Me.TxtKodeAkun1.ReadOnly = True
        Me.TxtKodeAkun1.Size = New System.Drawing.Size(36, 20)
        Me.TxtKodeAkun1.TabIndex = 0
        Me.TxtKodeAkun1.TabStop = False
        '
        'chkHeaderAccount
        '
        Me.chkHeaderAccount.AutoSize = True
        Me.chkHeaderAccount.Location = New System.Drawing.Point(375, 73)
        Me.chkHeaderAccount.Name = "chkHeaderAccount"
        Me.chkHeaderAccount.Size = New System.Drawing.Size(104, 17)
        Me.chkHeaderAccount.TabIndex = 4
        Me.chkHeaderAccount.Text = "Header Account"
        Me.chkHeaderAccount.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 127)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Beginning Balance"
        '
        'TxtBeginningBalance
        '
        Me.TxtBeginningBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBeginningBalance.Location = New System.Drawing.Point(143, 124)
        Me.TxtBeginningBalance.MaxLength = 18
        Me.TxtBeginningBalance.Name = "TxtBeginningBalance"
        Me.TxtBeginningBalance.Size = New System.Drawing.Size(121, 20)
        Me.TxtBeginningBalance.TabIndex = 6
        Me.TxtBeginningBalance.Text = "0"
        Me.TxtBeginningBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkActiveAccount
        '
        Me.chkActiveAccount.AutoSize = True
        Me.chkActiveAccount.Location = New System.Drawing.Point(270, 73)
        Me.chkActiveAccount.Name = "chkActiveAccount"
        Me.chkActiveAccount.Size = New System.Drawing.Size(99, 17)
        Me.chkActiveAccount.TabIndex = 3
        Me.chkActiveAccount.Text = "Active Account"
        Me.chkActiveAccount.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Description"
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(112, 45)
        Me.TxtDesc.MaxLength = 150
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(414, 20)
        Me.TxtDesc.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Account No"
        '
        'TxtAccountNo
        '
        Me.TxtAccountNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountNo.Location = New System.Drawing.Point(154, 19)
        Me.TxtAccountNo.MaxLength = 15
        Me.TxtAccountNo.Name = "TxtAccountNo"
        Me.TxtAccountNo.Size = New System.Drawing.Size(372, 20)
        Me.TxtAccountNo.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(270, 227)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(154, 227)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 286)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(545, 22)
        Me.StatusStrip1.TabIndex = 24
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(265, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(265, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 264)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(545, 22)
        Me.StatusStrip2.TabIndex = 25
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(265, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(265, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'FrmMasterAkun_Detail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(545, 308)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(561, 315)
        Me.Name = "FrmMasterAkun_Detail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Account Detail"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents chkHeaderAccount As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtBeginningBalance As System.Windows.Forms.TextBox
    Friend WithEvents chkActiveAccount As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtAccountNo As System.Windows.Forms.TextBox
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtKodeAkun1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccount As System.Windows.Forms.TextBox
    Friend WithEvents TxtBeginningBalanceOriginalCur As System.Windows.Forms.TextBox
    Friend WithEvents LblOriginalCur As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboAccountType As System.Windows.Forms.ComboBox
    Friend WithEvents LblIDR As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents cboCurrency As System.Windows.Forms.ComboBox
End Class
