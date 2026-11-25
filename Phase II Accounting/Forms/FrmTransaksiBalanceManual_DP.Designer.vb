<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiBalanceManual_DP
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryAccDP = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtPrimaryAcc = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtPrimaryCurr = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.LblAccount = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.LblAccountDP = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.lblExcRate = New System.Windows.Forms.Label
        Me.DtpDueDate = New System.Windows.Forms.DateTimePicker
        Me.cboCurrency = New System.Windows.Forms.ComboBox
        Me.TxtExchangeRate = New System.Windows.Forms.TextBox
        Me.cboAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryDPAccount = New System.Windows.Forms.ComboBox
        Me.CboDPAccount = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryDocAcuan = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPihakAcuan = New System.Windows.Forms.TextBox
        Me.TxtPrimaryTransaksi = New System.Windows.Forms.TextBox
        Me.LblDPNo = New System.Windows.Forms.Label
        Me.TxtAOPDP_No = New System.Windows.Forms.TextBox
        Me.TxtSuppInvNo = New System.Windows.Forms.TextBox
        Me.TxtPihakAcuanDesc = New System.Windows.Forms.TextBox
        Me.TxtPihakAcuanCode = New System.Windows.Forms.TextBox
        Me.BtnFindDocAcuan = New System.Windows.Forms.Button
        Me.TxtDocAcuanNo = New System.Windows.Forms.TextBox
        Me.LblInvNo = New System.Windows.Forms.Label
        Me.LblPihakAcuan = New System.Windows.Forms.Label
        Me.LblPONo = New System.Windows.Forms.Label
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Panel1)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryDocAcuan)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryPihakAcuan)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryTransaksi)
        Me.gboDetail.Controls.Add(Me.LblDPNo)
        Me.gboDetail.Controls.Add(Me.TxtAOPDP_No)
        Me.gboDetail.Controls.Add(Me.TxtSuppInvNo)
        Me.gboDetail.Controls.Add(Me.TxtPihakAcuanDesc)
        Me.gboDetail.Controls.Add(Me.TxtPihakAcuanCode)
        Me.gboDetail.Controls.Add(Me.BtnFindDocAcuan)
        Me.gboDetail.Controls.Add(Me.TxtDocAcuanNo)
        Me.gboDetail.Controls.Add(Me.LblInvNo)
        Me.gboDetail.Controls.Add(Me.LblPihakAcuan)
        Me.gboDetail.Controls.Add(Me.LblPONo)
        Me.gboDetail.Location = New System.Drawing.Point(5, -1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(530, 307)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Panel1.Controls.Add(Me.TxtRemarks)
        Me.Panel1.Controls.Add(Me.TxtBalance)
        Me.Panel1.Controls.Add(Me.DtpDate)
        Me.Panel1.Controls.Add(Me.TxtPrimaryAccDP)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.TxtPrimaryAcc)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.TxtPrimaryCurr)
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnOK)
        Me.Panel1.Controls.Add(Me.LblAccount)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.LblAccountDP)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.lblExcRate)
        Me.Panel1.Controls.Add(Me.DtpDueDate)
        Me.Panel1.Controls.Add(Me.cboCurrency)
        Me.Panel1.Controls.Add(Me.TxtExchangeRate)
        Me.Panel1.Controls.Add(Me.cboAccount)
        Me.Panel1.Controls.Add(Me.cboPrimaryAccount)
        Me.Panel1.Controls.Add(Me.cboPrimaryDPAccount)
        Me.Panel1.Controls.Add(Me.CboDPAccount)
        Me.Panel1.Location = New System.Drawing.Point(0, 113)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(530, 188)
        Me.Panel1.TabIndex = 0
        '
        'TxtRemarks
        '
        Me.TxtRemarks.AcceptsTab = True
        Me.TxtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(132, 106)
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(392, 20)
        Me.TxtRemarks.TabIndex = 7
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.Location = New System.Drawing.Point(229, 52)
        Me.TxtBalance.MaxLength = 17
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.Size = New System.Drawing.Size(121, 20)
        Me.TxtBalance.TabIndex = 3
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(132, 0)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(392, 20)
        Me.DtpDate.TabIndex = 0
        '
        'TxtPrimaryAccDP
        '
        Me.TxtPrimaryAccDP.Location = New System.Drawing.Point(100, 108)
        Me.TxtPrimaryAccDP.Name = "TxtPrimaryAccDP"
        Me.TxtPrimaryAccDP.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryAccDP.TabIndex = 63
        Me.TxtPrimaryAccDP.TabStop = False
        Me.TxtPrimaryAccDP.Text = "-1"
        Me.TxtPrimaryAccDP.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Date"
        '
        'TxtPrimaryAcc
        '
        Me.TxtPrimaryAcc.Location = New System.Drawing.Point(100, 81)
        Me.TxtPrimaryAcc.Name = "TxtPrimaryAcc"
        Me.TxtPrimaryAcc.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryAcc.TabIndex = 60
        Me.TxtPrimaryAcc.TabStop = False
        Me.TxtPrimaryAcc.Text = "-1"
        Me.TxtPrimaryAcc.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Due Date"
        '
        'TxtPrimaryCurr
        '
        Me.TxtPrimaryCurr.Location = New System.Drawing.Point(100, 53)
        Me.TxtPrimaryCurr.Name = "TxtPrimaryCurr"
        Me.TxtPrimaryCurr.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryCurr.TabIndex = 59
        Me.TxtPrimaryCurr.TabStop = False
        Me.TxtPrimaryCurr.Text = "-1"
        Me.TxtPrimaryCurr.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(264, 153)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(148, 153)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'LblAccount
        '
        Me.LblAccount.AutoSize = True
        Me.LblAccount.Location = New System.Drawing.Point(9, 84)
        Me.LblAccount.Name = "LblAccount"
        Me.LblAccount.Size = New System.Drawing.Size(64, 13)
        Me.LblAccount.TabIndex = 6
        Me.LblAccount.Text = "AP Account"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 109)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 58
        Me.Label6.Text = "Remarks"
        '
        'LblAccountDP
        '
        Me.LblAccountDP.AutoSize = True
        Me.LblAccountDP.Location = New System.Drawing.Point(9, 111)
        Me.LblAccountDP.Name = "LblAccountDP"
        Me.LblAccountDP.Size = New System.Drawing.Size(65, 13)
        Me.LblAccountDP.TabIndex = 7
        Me.LblAccountDP.Text = "DP Account"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 56)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Balance"
        '
        'lblExcRate
        '
        Me.lblExcRate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblExcRate.AutoSize = True
        Me.lblExcRate.Location = New System.Drawing.Point(356, 55)
        Me.lblExcRate.Name = "lblExcRate"
        Me.lblExcRate.Size = New System.Drawing.Size(81, 13)
        Me.lblExcRate.TabIndex = 9
        Me.lblExcRate.Text = "Exchange Rate"
        '
        'DtpDueDate
        '
        Me.DtpDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDueDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDueDate.Location = New System.Drawing.Point(132, 26)
        Me.DtpDueDate.Name = "DtpDueDate"
        Me.DtpDueDate.Size = New System.Drawing.Size(392, 20)
        Me.DtpDueDate.TabIndex = 1
        '
        'cboCurrency
        '
        Me.cboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrency.FormattingEnabled = True
        Me.cboCurrency.Location = New System.Drawing.Point(132, 52)
        Me.cboCurrency.Name = "cboCurrency"
        Me.cboCurrency.Size = New System.Drawing.Size(91, 21)
        Me.cboCurrency.TabIndex = 2
        '
        'TxtExchangeRate
        '
        Me.TxtExchangeRate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtExchangeRate.Location = New System.Drawing.Point(443, 52)
        Me.TxtExchangeRate.Name = "TxtExchangeRate"
        Me.TxtExchangeRate.ReadOnly = True
        Me.TxtExchangeRate.Size = New System.Drawing.Size(81, 20)
        Me.TxtExchangeRate.TabIndex = 4
        '
        'cboAccount
        '
        Me.cboAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(132, 79)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(392, 21)
        Me.cboAccount.TabIndex = 5
        '
        'cboPrimaryAccount
        '
        Me.cboPrimaryAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryAccount.FormattingEnabled = True
        Me.cboPrimaryAccount.Location = New System.Drawing.Point(132, 79)
        Me.cboPrimaryAccount.Name = "cboPrimaryAccount"
        Me.cboPrimaryAccount.Size = New System.Drawing.Size(91, 21)
        Me.cboPrimaryAccount.TabIndex = 64
        Me.cboPrimaryAccount.TabStop = False
        '
        'cboPrimaryDPAccount
        '
        Me.cboPrimaryDPAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryDPAccount.FormattingEnabled = True
        Me.cboPrimaryDPAccount.Location = New System.Drawing.Point(132, 106)
        Me.cboPrimaryDPAccount.Name = "cboPrimaryDPAccount"
        Me.cboPrimaryDPAccount.Size = New System.Drawing.Size(91, 21)
        Me.cboPrimaryDPAccount.TabIndex = 65
        Me.cboPrimaryDPAccount.TabStop = False
        '
        'CboDPAccount
        '
        Me.CboDPAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboDPAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboDPAccount.FormattingEnabled = True
        Me.CboDPAccount.Location = New System.Drawing.Point(132, 106)
        Me.CboDPAccount.Name = "CboDPAccount"
        Me.CboDPAccount.Size = New System.Drawing.Size(392, 21)
        Me.CboDPAccount.TabIndex = 6
        '
        'TxtPrimaryDocAcuan
        '
        Me.TxtPrimaryDocAcuan.Location = New System.Drawing.Point(100, 38)
        Me.TxtPrimaryDocAcuan.Name = "TxtPrimaryDocAcuan"
        Me.TxtPrimaryDocAcuan.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryDocAcuan.TabIndex = 62
        Me.TxtPrimaryDocAcuan.TabStop = False
        Me.TxtPrimaryDocAcuan.Text = "-1"
        Me.TxtPrimaryDocAcuan.Visible = False
        '
        'TxtPrimaryPihakAcuan
        '
        Me.TxtPrimaryPihakAcuan.Location = New System.Drawing.Point(100, 66)
        Me.TxtPrimaryPihakAcuan.Name = "TxtPrimaryPihakAcuan"
        Me.TxtPrimaryPihakAcuan.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryPihakAcuan.TabIndex = 61
        Me.TxtPrimaryPihakAcuan.TabStop = False
        Me.TxtPrimaryPihakAcuan.Text = "-1"
        Me.TxtPrimaryPihakAcuan.Visible = False
        '
        'TxtPrimaryTransaksi
        '
        Me.TxtPrimaryTransaksi.Location = New System.Drawing.Point(100, 13)
        Me.TxtPrimaryTransaksi.Name = "TxtPrimaryTransaksi"
        Me.TxtPrimaryTransaksi.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryTransaksi.TabIndex = 56
        Me.TxtPrimaryTransaksi.TabStop = False
        Me.TxtPrimaryTransaksi.Text = "-1"
        Me.TxtPrimaryTransaksi.Visible = False
        '
        'LblDPNo
        '
        Me.LblDPNo.AutoSize = True
        Me.LblDPNo.Location = New System.Drawing.Point(9, 16)
        Me.LblDPNo.Name = "LblDPNo"
        Me.LblDPNo.Size = New System.Drawing.Size(64, 13)
        Me.LblDPNo.TabIndex = 55
        Me.LblDPNo.Text = "AOP DP No"
        '
        'TxtAOPDP_No
        '
        Me.TxtAOPDP_No.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAOPDP_No.Location = New System.Drawing.Point(132, 13)
        Me.TxtAOPDP_No.Name = "TxtAOPDP_No"
        Me.TxtAOPDP_No.Size = New System.Drawing.Size(392, 20)
        Me.TxtAOPDP_No.TabIndex = 0
        '
        'TxtSuppInvNo
        '
        Me.TxtSuppInvNo.AcceptsTab = True
        Me.TxtSuppInvNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSuppInvNo.Location = New System.Drawing.Point(132, 90)
        Me.TxtSuppInvNo.Name = "TxtSuppInvNo"
        Me.TxtSuppInvNo.Size = New System.Drawing.Size(392, 20)
        Me.TxtSuppInvNo.TabIndex = 2
        '
        'TxtPihakAcuanDesc
        '
        Me.TxtPihakAcuanDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPihakAcuanDesc.Location = New System.Drawing.Point(229, 64)
        Me.TxtPihakAcuanDesc.Name = "TxtPihakAcuanDesc"
        Me.TxtPihakAcuanDesc.ReadOnly = True
        Me.TxtPihakAcuanDesc.Size = New System.Drawing.Size(295, 20)
        Me.TxtPihakAcuanDesc.TabIndex = 13
        Me.TxtPihakAcuanDesc.TabStop = False
        '
        'TxtPihakAcuanCode
        '
        Me.TxtPihakAcuanCode.Location = New System.Drawing.Point(132, 64)
        Me.TxtPihakAcuanCode.Name = "TxtPihakAcuanCode"
        Me.TxtPihakAcuanCode.ReadOnly = True
        Me.TxtPihakAcuanCode.Size = New System.Drawing.Size(91, 20)
        Me.TxtPihakAcuanCode.TabIndex = 12
        Me.TxtPihakAcuanCode.TabStop = False
        '
        'BtnFindDocAcuan
        '
        Me.BtnFindDocAcuan.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindDocAcuan.Location = New System.Drawing.Point(481, 36)
        Me.BtnFindDocAcuan.Name = "BtnFindDocAcuan"
        Me.BtnFindDocAcuan.Size = New System.Drawing.Size(43, 23)
        Me.BtnFindDocAcuan.TabIndex = 11
        Me.BtnFindDocAcuan.TabStop = False
        Me.BtnFindDocAcuan.Text = "Find"
        Me.BtnFindDocAcuan.UseVisualStyleBackColor = True
        '
        'TxtDocAcuanNo
        '
        Me.TxtDocAcuanNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocAcuanNo.Location = New System.Drawing.Point(132, 38)
        Me.TxtDocAcuanNo.Name = "TxtDocAcuanNo"
        Me.TxtDocAcuanNo.Size = New System.Drawing.Size(343, 20)
        Me.TxtDocAcuanNo.TabIndex = 1
        '
        'LblInvNo
        '
        Me.LblInvNo.AutoSize = True
        Me.LblInvNo.Location = New System.Drawing.Point(9, 95)
        Me.LblInvNo.Name = "LblInvNo"
        Me.LblInvNo.Size = New System.Drawing.Size(100, 13)
        Me.LblInvNo.TabIndex = 2
        Me.LblInvNo.Text = "Supplier Invoice No"
        '
        'LblPihakAcuan
        '
        Me.LblPihakAcuan.AutoSize = True
        Me.LblPihakAcuan.Location = New System.Drawing.Point(9, 69)
        Me.LblPihakAcuan.Name = "LblPihakAcuan"
        Me.LblPihakAcuan.Size = New System.Drawing.Size(45, 13)
        Me.LblPihakAcuan.TabIndex = 1
        Me.LblPihakAcuan.Text = "Supplier"
        '
        'LblPONo
        '
        Me.LblPONo.AutoSize = True
        Me.LblPONo.Location = New System.Drawing.Point(9, 42)
        Me.LblPONo.Name = "LblPONo"
        Me.LblPONo.Size = New System.Drawing.Size(98, 13)
        Me.LblPONo.TabIndex = 0
        Me.LblPONo.Text = "Purchase Order No"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 366)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(549, 22)
        Me.StatusStrip2.TabIndex = 55
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(291, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(291, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 388)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(549, 22)
        Me.StatusStrip1.TabIndex = 54
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(291, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(291, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(424, 315)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(91, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(325, 315)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(91, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(226, 315)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(91, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(127, 315)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(91, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(28, 315)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(91, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'FrmTransaksiBalanceManual_DP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(549, 410)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimumSize = New System.Drawing.Size(557, 437)
        Me.Name = "FrmTransaksiBalanceManual_DP"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Historical Purchase Down Payment"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents LblPihakAcuan As System.Windows.Forms.Label
    Friend WithEvents LblPONo As System.Windows.Forms.Label
    Friend WithEvents lblExcRate As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents LblAccountDP As System.Windows.Forms.Label
    Friend WithEvents LblAccount As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblInvNo As System.Windows.Forms.Label
    Friend WithEvents CboDPAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboAccount As System.Windows.Forms.ComboBox
    Friend WithEvents TxtExchangeRate As System.Windows.Forms.TextBox
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents cboCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents DtpDueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtSuppInvNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtPihakAcuanDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPihakAcuanCode As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindDocAcuan As System.Windows.Forms.Button
    Friend WithEvents TxtDocAcuanNo As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents LblDPNo As System.Windows.Forms.Label
    Friend WithEvents TxtAOPDP_No As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCurr As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryPihakAcuan As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryDocAcuan As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccDP As System.Windows.Forms.TextBox
    Friend WithEvents cboPrimaryDPAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryAccount As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
