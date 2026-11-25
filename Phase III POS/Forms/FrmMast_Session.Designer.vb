<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMast_Session
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
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpSessionThrough = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.dtpSessionPeriod = New System.Windows.Forms.DateTimePicker
        Me.TxtBranchName = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.TxtCashierName = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.BtnFindBranchCode = New System.Windows.Forms.Button
        Me.TxtPrimaryBranch = New System.Windows.Forms.TextBox
        Me.TxtBranchCode = New System.Windows.Forms.TextBox
        Me.btnFindCashier = New System.Windows.Forms.Button
        Me.TxtPrimaryCashier = New System.Windows.Forms.TextBox
        Me.TxtCashierCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtReferenceName = New System.Windows.Forms.TextBox
        Me.TxtCash_Beginning = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtCashierCashAccountDesc = New System.Windows.Forms.TextBox
        Me.btnFindCashierCashAccount = New System.Windows.Forms.Button
        Me.TxtBalanceCashAccountDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCashierCashAccount = New System.Windows.Forms.TextBox
        Me.btnFindBalanceCashAccount = New System.Windows.Forms.Button
        Me.TxtCashierCashAccountCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryBalanceCashAccount = New System.Windows.Forms.TextBox
        Me.TxtBalanceCashAccountCode = New System.Windows.Forms.TextBox
        Me.TxtBeginingCashAccountDesc = New System.Windows.Forms.TextBox
        Me.btnFindBeginingCashAccount = New System.Windows.Forms.Button
        Me.TxtPrimaryBeginingCashAccount = New System.Windows.Forms.TextBox
        Me.TxtBeginingCashAccountCode = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(216, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 39
        Me.Label3.Text = "Branch"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(245, 33)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(43, 13)
        Me.Label15.TabIndex = 58
        Me.Label15.Text = "through"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 36
        Me.Label2.Text = "Cashier Code"
        '
        'DtpSessionThrough
        '
        Me.DtpSessionThrough.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpSessionThrough.CustomFormat = "HH:mm"
        Me.DtpSessionThrough.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpSessionThrough.Location = New System.Drawing.Point(297, 29)
        Me.DtpSessionThrough.Name = "DtpSessionThrough"
        Me.DtpSessionThrough.ShowUpDown = True
        Me.DtpSessionThrough.Size = New System.Drawing.Size(79, 20)
        Me.DtpSessionThrough.TabIndex = 2
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(5, 33)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(77, 13)
        Me.Label14.TabIndex = 56
        Me.Label14.Text = "Session Period"
        '
        'dtpSessionPeriod
        '
        Me.dtpSessionPeriod.CustomFormat = "HH:mm"
        Me.dtpSessionPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSessionPeriod.Location = New System.Drawing.Point(99, 29)
        Me.dtpSessionPeriod.Name = "dtpSessionPeriod"
        Me.dtpSessionPeriod.ShowUpDown = True
        Me.dtpSessionPeriod.Size = New System.Drawing.Size(140, 20)
        Me.dtpSessionPeriod.TabIndex = 1
        '
        'TxtBranchName
        '
        Me.TxtBranchName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBranchName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBranchName.Location = New System.Drawing.Point(248, 81)
        Me.TxtBranchName.MaxLength = 150
        Me.TxtBranchName.Name = "TxtBranchName"
        Me.TxtBranchName.ReadOnly = True
        Me.TxtBranchName.Size = New System.Drawing.Size(128, 20)
        Me.TxtBranchName.TabIndex = 44
        Me.TxtBranchName.TabStop = False
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(216, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(311, 153)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(501, 191)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(427, 153)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(99, 109)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(277, 20)
        Me.TxtRemarks.TabIndex = 5
        '
        'TxtCashierName
        '
        Me.TxtCashierName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCashierName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCashierName.Location = New System.Drawing.Point(248, 55)
        Me.TxtCashierName.MaxLength = 150
        Me.TxtCashierName.Name = "TxtCashierName"
        Me.TxtCashierName.ReadOnly = True
        Me.TxtCashierName.Size = New System.Drawing.Size(128, 20)
        Me.TxtCashierName.TabIndex = 43
        Me.TxtCashierName.TabStop = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(153, 191)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(110, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(130, 107)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(77, 20)
        Me.TxtReadOnly.TabIndex = 6
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 220)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(880, 22)
        Me.StatusStrip2.TabIndex = 53
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(216, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(216, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 109)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 35
        Me.Label7.Text = "Remarks"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(18, 3)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimary.TabIndex = 23
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(385, 191)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(110, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(617, 191)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(110, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(269, 191)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(110, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(848, 185)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label15)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpSessionThrough)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpSessionPeriod)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtBranchName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCashierName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindBranchCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryBranch)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtBranchCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCashier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryCashier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCashierCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtReferenceName)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCash_Beginning)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCashierCashAccountDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindCashierCashAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtBalanceCashAccountDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryCashierCashAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindBalanceCashAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCashierCashAccountCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryBalanceCashAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtBalanceCashAccountCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtBeginingCashAccountDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindBeginingCashAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryBeginingCashAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtBeginingCashAccountCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Size = New System.Drawing.Size(836, 134)
        Me.SplitContainer1.SplitterDistance = 379
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'BtnFindBranchCode
        '
        Me.BtnFindBranchCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindBranchCode.Location = New System.Drawing.Point(188, 81)
        Me.BtnFindBranchCode.Name = "BtnFindBranchCode"
        Me.BtnFindBranchCode.Size = New System.Drawing.Size(51, 20)
        Me.BtnFindBranchCode.TabIndex = 41
        Me.BtnFindBranchCode.TabStop = False
        Me.BtnFindBranchCode.Text = "Find"
        Me.BtnFindBranchCode.UseVisualStyleBackColor = True
        '
        'TxtPrimaryBranch
        '
        Me.TxtPrimaryBranch.Location = New System.Drawing.Point(18, 81)
        Me.TxtPrimaryBranch.Name = "TxtPrimaryBranch"
        Me.TxtPrimaryBranch.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryBranch.TabIndex = 42
        Me.TxtPrimaryBranch.TabStop = False
        Me.TxtPrimaryBranch.Text = "-1"
        Me.TxtPrimaryBranch.Visible = False
        '
        'TxtBranchCode
        '
        Me.TxtBranchCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBranchCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBranchCode.Location = New System.Drawing.Point(99, 81)
        Me.TxtBranchCode.MaxLength = 150
        Me.TxtBranchCode.Name = "TxtBranchCode"
        Me.TxtBranchCode.Size = New System.Drawing.Size(83, 20)
        Me.TxtBranchCode.TabIndex = 4
        '
        'btnFindCashier
        '
        Me.btnFindCashier.Location = New System.Drawing.Point(188, 55)
        Me.btnFindCashier.Name = "btnFindCashier"
        Me.btnFindCashier.Size = New System.Drawing.Size(51, 20)
        Me.btnFindCashier.TabIndex = 37
        Me.btnFindCashier.TabStop = False
        Me.btnFindCashier.Text = "Find"
        Me.btnFindCashier.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCashier
        '
        Me.TxtPrimaryCashier.Location = New System.Drawing.Point(18, 55)
        Me.TxtPrimaryCashier.Name = "TxtPrimaryCashier"
        Me.TxtPrimaryCashier.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCashier.TabIndex = 38
        Me.TxtPrimaryCashier.TabStop = False
        Me.TxtPrimaryCashier.Text = "-1"
        Me.TxtPrimaryCashier.Visible = False
        '
        'TxtCashierCode
        '
        Me.TxtCashierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCashierCode.Location = New System.Drawing.Point(99, 55)
        Me.TxtCashierCode.MaxLength = 150
        Me.TxtCashierCode.Name = "TxtCashierCode"
        Me.TxtCashierCode.Size = New System.Drawing.Size(83, 20)
        Me.TxtCashierCode.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Reference Name"
        '
        'TxtReferenceName
        '
        Me.TxtReferenceName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReferenceName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReferenceName.Location = New System.Drawing.Point(99, 3)
        Me.TxtReferenceName.MaxLength = 150
        Me.TxtReferenceName.Name = "TxtReferenceName"
        Me.TxtReferenceName.Size = New System.Drawing.Size(277, 20)
        Me.TxtReferenceName.TabIndex = 0
        '
        'TxtCash_Beginning
        '
        Me.TxtCash_Beginning.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCash_Beginning.Location = New System.Drawing.Point(130, 3)
        Me.TxtCash_Beginning.MaxLength = 12
        Me.TxtCash_Beginning.Name = "TxtCash_Beginning"
        Me.TxtCash_Beginning.Size = New System.Drawing.Size(320, 20)
        Me.TxtCash_Beginning.TabIndex = 0
        Me.TxtCash_Beginning.Text = "0"
        Me.TxtCash_Beginning.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 58)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 13)
        Me.Label13.TabIndex = 101
        Me.Label13.Text = "Cashier Cash Account"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 84)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(116, 13)
        Me.Label12.TabIndex = 96
        Me.Label12.Text = "Balance Cash Account"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(124, 13)
        Me.Label11.TabIndex = 91
        Me.Label11.Text = "Beginning Cash Account"
        '
        'TxtCashierCashAccountDesc
        '
        Me.TxtCashierCashAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCashierCashAccountDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCashierCashAccountDesc.Location = New System.Drawing.Point(273, 55)
        Me.TxtCashierCashAccountDesc.MaxLength = 150
        Me.TxtCashierCashAccountDesc.Name = "TxtCashierCashAccountDesc"
        Me.TxtCashierCashAccountDesc.ReadOnly = True
        Me.TxtCashierCashAccountDesc.Size = New System.Drawing.Size(177, 20)
        Me.TxtCashierCashAccountDesc.TabIndex = 105
        Me.TxtCashierCashAccountDesc.TabStop = False
        '
        'btnFindCashierCashAccount
        '
        Me.btnFindCashierCashAccount.Location = New System.Drawing.Point(214, 55)
        Me.btnFindCashierCashAccount.Name = "btnFindCashierCashAccount"
        Me.btnFindCashierCashAccount.Size = New System.Drawing.Size(51, 20)
        Me.btnFindCashierCashAccount.TabIndex = 103
        Me.btnFindCashierCashAccount.TabStop = False
        Me.btnFindCashierCashAccount.Text = "Find"
        Me.btnFindCashierCashAccount.UseVisualStyleBackColor = True
        '
        'TxtBalanceCashAccountDesc
        '
        Me.TxtBalanceCashAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalanceCashAccountDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBalanceCashAccountDesc.Location = New System.Drawing.Point(273, 81)
        Me.TxtBalanceCashAccountDesc.MaxLength = 150
        Me.TxtBalanceCashAccountDesc.Name = "TxtBalanceCashAccountDesc"
        Me.TxtBalanceCashAccountDesc.ReadOnly = True
        Me.TxtBalanceCashAccountDesc.Size = New System.Drawing.Size(177, 20)
        Me.TxtBalanceCashAccountDesc.TabIndex = 100
        Me.TxtBalanceCashAccountDesc.TabStop = False
        '
        'TxtPrimaryCashierCashAccount
        '
        Me.TxtPrimaryCashierCashAccount.Location = New System.Drawing.Point(78, 55)
        Me.TxtPrimaryCashierCashAccount.Name = "TxtPrimaryCashierCashAccount"
        Me.TxtPrimaryCashierCashAccount.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCashierCashAccount.TabIndex = 104
        Me.TxtPrimaryCashierCashAccount.TabStop = False
        Me.TxtPrimaryCashierCashAccount.Text = "-1"
        Me.TxtPrimaryCashierCashAccount.Visible = False
        '
        'btnFindBalanceCashAccount
        '
        Me.btnFindBalanceCashAccount.Location = New System.Drawing.Point(214, 81)
        Me.btnFindBalanceCashAccount.Name = "btnFindBalanceCashAccount"
        Me.btnFindBalanceCashAccount.Size = New System.Drawing.Size(51, 20)
        Me.btnFindBalanceCashAccount.TabIndex = 98
        Me.btnFindBalanceCashAccount.TabStop = False
        Me.btnFindBalanceCashAccount.Text = "Find"
        Me.btnFindBalanceCashAccount.UseVisualStyleBackColor = True
        '
        'TxtCashierCashAccountCode
        '
        Me.TxtCashierCashAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCashierCashAccountCode.Location = New System.Drawing.Point(130, 55)
        Me.TxtCashierCashAccountCode.MaxLength = 150
        Me.TxtCashierCashAccountCode.Name = "TxtCashierCashAccountCode"
        Me.TxtCashierCashAccountCode.Size = New System.Drawing.Size(78, 20)
        Me.TxtCashierCashAccountCode.TabIndex = 2
        '
        'TxtPrimaryBalanceCashAccount
        '
        Me.TxtPrimaryBalanceCashAccount.Location = New System.Drawing.Point(78, 81)
        Me.TxtPrimaryBalanceCashAccount.Name = "TxtPrimaryBalanceCashAccount"
        Me.TxtPrimaryBalanceCashAccount.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryBalanceCashAccount.TabIndex = 99
        Me.TxtPrimaryBalanceCashAccount.TabStop = False
        Me.TxtPrimaryBalanceCashAccount.Text = "-1"
        Me.TxtPrimaryBalanceCashAccount.Visible = False
        '
        'TxtBalanceCashAccountCode
        '
        Me.TxtBalanceCashAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBalanceCashAccountCode.Location = New System.Drawing.Point(130, 81)
        Me.TxtBalanceCashAccountCode.MaxLength = 150
        Me.TxtBalanceCashAccountCode.Name = "TxtBalanceCashAccountCode"
        Me.TxtBalanceCashAccountCode.Size = New System.Drawing.Size(78, 20)
        Me.TxtBalanceCashAccountCode.TabIndex = 3
        '
        'TxtBeginingCashAccountDesc
        '
        Me.TxtBeginingCashAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBeginingCashAccountDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBeginingCashAccountDesc.Location = New System.Drawing.Point(273, 29)
        Me.TxtBeginingCashAccountDesc.MaxLength = 150
        Me.TxtBeginingCashAccountDesc.Name = "TxtBeginingCashAccountDesc"
        Me.TxtBeginingCashAccountDesc.ReadOnly = True
        Me.TxtBeginingCashAccountDesc.Size = New System.Drawing.Size(177, 20)
        Me.TxtBeginingCashAccountDesc.TabIndex = 95
        Me.TxtBeginingCashAccountDesc.TabStop = False
        '
        'btnFindBeginingCashAccount
        '
        Me.btnFindBeginingCashAccount.Location = New System.Drawing.Point(214, 29)
        Me.btnFindBeginingCashAccount.Name = "btnFindBeginingCashAccount"
        Me.btnFindBeginingCashAccount.Size = New System.Drawing.Size(51, 20)
        Me.btnFindBeginingCashAccount.TabIndex = 93
        Me.btnFindBeginingCashAccount.TabStop = False
        Me.btnFindBeginingCashAccount.Text = "Find"
        Me.btnFindBeginingCashAccount.UseVisualStyleBackColor = True
        '
        'TxtPrimaryBeginingCashAccount
        '
        Me.TxtPrimaryBeginingCashAccount.Location = New System.Drawing.Point(78, 29)
        Me.TxtPrimaryBeginingCashAccount.Name = "TxtPrimaryBeginingCashAccount"
        Me.TxtPrimaryBeginingCashAccount.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryBeginingCashAccount.TabIndex = 94
        Me.TxtPrimaryBeginingCashAccount.TabStop = False
        Me.TxtPrimaryBeginingCashAccount.Text = "-1"
        Me.TxtPrimaryBeginingCashAccount.Visible = False
        '
        'TxtBeginingCashAccountCode
        '
        Me.TxtBeginingCashAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBeginingCashAccountCode.Location = New System.Drawing.Point(130, 29)
        Me.TxtBeginingCashAccountCode.MaxLength = 150
        Me.TxtBeginingCashAccountCode.Name = "TxtBeginingCashAccountCode"
        Me.TxtBeginingCashAccountCode.Size = New System.Drawing.Size(78, 20)
        Me.TxtBeginingCashAccountCode.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 3)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 13)
        Me.Label5.TabIndex = 90
        Me.Label5.Text = "Beginning Cash"
        '
        'FrmMast_Session
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(880, 242)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(896, 280)
        Me.Name = "FrmMast_Session"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Session Reference Data"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpSessionThrough As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dtpSessionPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtBranchName As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtCashierName As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents BtnFindBranchCode As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryBranch As System.Windows.Forms.TextBox
    Friend WithEvents TxtBranchCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCashier As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCashier As System.Windows.Forms.TextBox
    Friend WithEvents TxtCashierCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtReferenceName As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtCashierCashAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCashierCashAccount As System.Windows.Forms.Button
    Friend WithEvents TxtBalanceCashAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCashierCashAccount As System.Windows.Forms.TextBox
    Friend WithEvents btnFindBalanceCashAccount As System.Windows.Forms.Button
    Friend WithEvents TxtCashierCashAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryBalanceCashAccount As System.Windows.Forms.TextBox
    Friend WithEvents TxtBalanceCashAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtBeginingCashAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindBeginingCashAccount As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryBeginingCashAccount As System.Windows.Forms.TextBox
    Friend WithEvents TxtBeginingCashAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCash_Beginning As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
