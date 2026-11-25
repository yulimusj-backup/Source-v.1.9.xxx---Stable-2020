<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiAOPDP
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
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnEdit = New System.Windows.Forms.Button
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtTaxExchangeRate = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtExhangeRateToIDR = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtExchangeRateToUSD = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtTaxAmount = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.chkTax = New System.Windows.Forms.CheckBox
        Me.txtSubTotal = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.dtpSupplierInvoiceDate = New System.Windows.Forms.DateTimePicker
        Me.txtPrimaryDPAccount = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.cboDPAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryDPAccount = New System.Windows.Forms.ComboBox
        Me.lblCurrTax = New System.Windows.Forms.Label
        Me.txtPrimaryCurrTax = New System.Windows.Forms.TextBox
        Me.cboCurrTax = New System.Windows.Forms.ComboBox
        Me.cboPrimaryCurrTax = New System.Windows.Forms.ComboBox
        Me.txtPrimaryAccount = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cboAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryAccount = New System.Windows.Forms.ComboBox
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak = New System.Windows.Forms.TextBox
        Me.cboPPNDipungut = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtNPWPNo = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtTaxInvoiceNumber = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSupplierInvoiceNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtPaymentsTerms = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDutyStamp = New System.Windows.Forms.TextBox
        Me.btnFindPO = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtPrimaryPO = New System.Windows.Forms.TextBox
        Me.txtPONo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryAOP = New System.Windows.Forms.TextBox
        Me.TxtAOPNo = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(323, 267)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(572, 267)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(237, 267)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(489, 267)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'txtTax
        '
        Me.txtTax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTax.Location = New System.Drawing.Point(538, 46)
        Me.txtTax.MaxLength = 50
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.Size = New System.Drawing.Size(109, 20)
        Me.txtTax.TabIndex = 16
        Me.txtTax.Text = "0"
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(538, 121)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(324, 20)
        Me.TxtTotal.TabIndex = 22
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(447, 22)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(53, 13)
        Me.Label15.TabIndex = 32
        Me.Label15.Text = "Sub Total"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(447, 125)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(31, 13)
        Me.Label24.TabIndex = 40
        Me.Label24.Text = "Total"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(538, 146)
        Me.TxtRemark.MaxLength = 50
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(324, 20)
        Me.TxtRemark.TabIndex = 23
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(447, 150)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 42
        Me.Label6.Text = "Remarks"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(406, 267)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(447, 73)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(25, 13)
        Me.Label18.TabIndex = 32
        Me.Label18.Text = "Tax"
        '
        'txtTaxExchangeRate
        '
        Me.txtTaxExchangeRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxExchangeRate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxExchangeRate.Location = New System.Drawing.Point(782, 46)
        Me.txtTaxExchangeRate.MaxLength = 50
        Me.txtTaxExchangeRate.Name = "txtTaxExchangeRate"
        Me.txtTaxExchangeRate.ReadOnly = True
        Me.txtTaxExchangeRate.Size = New System.Drawing.Size(80, 20)
        Me.txtTaxExchangeRate.TabIndex = 17
        Me.txtTaxExchangeRate.Text = "1"
        Me.txtTaxExchangeRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(674, 50)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(102, 13)
        Me.Label17.TabIndex = 30
        Me.Label17.Text = "Tax Exchange Rate"
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCurrency.Location = New System.Drawing.Point(180, 95)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(32, 20)
        Me.btnFindCurrency.TabIndex = 6
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "..."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        Me.btnFindCurrency.Visible = False
        '
        'txtExhangeRateToIDR
        '
        Me.txtExhangeRateToIDR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExhangeRateToIDR.Location = New System.Drawing.Point(677, 302)
        Me.txtExhangeRateToIDR.MaxLength = 50
        Me.txtExhangeRateToIDR.Name = "txtExhangeRateToIDR"
        Me.txtExhangeRateToIDR.Size = New System.Drawing.Size(293, 20)
        Me.txtExhangeRateToIDR.TabIndex = 0
        Me.txtExhangeRateToIDR.TabStop = False
        Me.txtExhangeRateToIDR.Text = "1"
        Me.txtExhangeRateToIDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExhangeRateToIDR.Visible = False
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(550, 307)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(119, 13)
        Me.Label23.TabIndex = 225
        Me.Label23.Text = "Exchange Rate To IDR"
        Me.Label23.Visible = False
        '
        'txtExchangeRateToUSD
        '
        Me.txtExchangeRateToUSD.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExchangeRateToUSD.Location = New System.Drawing.Point(677, 276)
        Me.txtExchangeRateToUSD.MaxLength = 50
        Me.txtExchangeRateToUSD.Name = "txtExchangeRateToUSD"
        Me.txtExchangeRateToUSD.Size = New System.Drawing.Size(293, 20)
        Me.txtExchangeRateToUSD.TabIndex = 41
        Me.txtExchangeRateToUSD.TabStop = False
        Me.txtExchangeRateToUSD.Text = "1"
        Me.txtExchangeRateToUSD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExchangeRateToUSD.Visible = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(550, 281)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(123, 13)
        Me.Label21.TabIndex = 223
        Me.Label21.Text = "Exchange Rate To USD"
        Me.Label21.Visible = False
        '
        'txtTaxAmount
        '
        Me.txtTaxAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount.Location = New System.Drawing.Point(656, 69)
        Me.txtTaxAmount.MaxLength = 50
        Me.txtTaxAmount.Name = "txtTaxAmount"
        Me.txtTaxAmount.ReadOnly = True
        Me.txtTaxAmount.Size = New System.Drawing.Size(206, 20)
        Me.txtTaxAmount.TabIndex = 20
        Me.txtTaxAmount.TabStop = False
        Me.txtTaxAmount.Text = "0"
        Me.txtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(653, 50)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(15, 13)
        Me.Label25.TabIndex = 29
        Me.Label25.Text = "%"
        '
        'chkTax
        '
        Me.chkTax.AutoSize = True
        Me.chkTax.Enabled = False
        Me.chkTax.Location = New System.Drawing.Point(447, 48)
        Me.chkTax.Name = "chkTax"
        Me.chkTax.Size = New System.Drawing.Size(85, 17)
        Me.chkTax.TabIndex = 27
        Me.chkTax.Text = "Tax Exclude"
        Me.chkTax.UseVisualStyleBackColor = True
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSubTotal.Location = New System.Drawing.Point(538, 18)
        Me.txtSubTotal.MaxLength = 50
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.Size = New System.Drawing.Size(324, 20)
        Me.txtSubTotal.TabIndex = 15
        Me.txtSubTotal.Text = "0"
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 297)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(893, 22)
        Me.StatusStrip1.TabIndex = 47
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Label14)
        Me.gboDetail.Controls.Add(Me.dtpSupplierInvoiceDate)
        Me.gboDetail.Controls.Add(Me.txtPrimaryDPAccount)
        Me.gboDetail.Controls.Add(Me.Label12)
        Me.gboDetail.Controls.Add(Me.cboDPAccount)
        Me.gboDetail.Controls.Add(Me.cboPrimaryDPAccount)
        Me.gboDetail.Controls.Add(Me.lblCurrTax)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCurrTax)
        Me.gboDetail.Controls.Add(Me.cboCurrTax)
        Me.gboDetail.Controls.Add(Me.cboPrimaryCurrTax)
        Me.gboDetail.Controls.Add(Me.txtPrimaryAccount)
        Me.gboDetail.Controls.Add(Me.Label10)
        Me.gboDetail.Controls.Add(Me.cboAccount)
        Me.gboDetail.Controls.Add(Me.cboPrimaryAccount)
        Me.gboDetail.Controls.Add(Me.Txt_E_Faktur_KodeJenisTransaksiPajak)
        Me.gboDetail.Controls.Add(Me.cboPPNDipungut)
        Me.gboDetail.Controls.Add(Me.Label13)
        Me.gboDetail.Controls.Add(Me.TxtNPWPNo)
        Me.gboDetail.Controls.Add(Me.Label11)
        Me.gboDetail.Controls.Add(Me.TxtTaxInvoiceNumber)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.txtSupplierInvoiceNo)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.txtCurrencyCode)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.Label9)
        Me.gboDetail.Controls.Add(Me.TxtPaymentsTerms)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.txtDutyStamp)
        Me.gboDetail.Controls.Add(Me.btnFindPO)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.txtPrimaryPO)
        Me.gboDetail.Controls.Add(Me.txtPONo)
        Me.gboDetail.Controls.Add(Me.txtExhangeRateToIDR)
        Me.gboDetail.Controls.Add(Me.Label23)
        Me.gboDetail.Controls.Add(Me.Label18)
        Me.gboDetail.Controls.Add(Me.txtTaxExchangeRate)
        Me.gboDetail.Controls.Add(Me.Label17)
        Me.gboDetail.Controls.Add(Me.btnFindCurrency)
        Me.gboDetail.Controls.Add(Me.txtTaxAmount)
        Me.gboDetail.Controls.Add(Me.Label25)
        Me.gboDetail.Controls.Add(Me.chkTax)
        Me.gboDetail.Controls.Add(Me.txtTax)
        Me.gboDetail.Controls.Add(Me.Label24)
        Me.gboDetail.Controls.Add(Me.TxtTotal)
        Me.gboDetail.Controls.Add(Me.Label15)
        Me.gboDetail.Controls.Add(Me.txtSubTotal)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtPrimarySupplier)
        Me.gboDetail.Controls.Add(Me.txtSupplierCode)
        Me.gboDetail.Controls.Add(Me.TxtSupplierName)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.Label22)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryAOP)
        Me.gboDetail.Controls.Add(Me.TxtAOPNo)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.Label21)
        Me.gboDetail.Controls.Add(Me.txtExchangeRateToUSD)
        Me.gboDetail.Location = New System.Drawing.Point(9, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(874, 261)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(256, 124)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(30, 13)
        Me.Label14.TabIndex = 226
        Me.Label14.Text = "Date"
        '
        'dtpSupplierInvoiceDate
        '
        Me.dtpSupplierInvoiceDate.CustomFormat = "dd MMM yyyy"
        Me.dtpSupplierInvoiceDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSupplierInvoiceDate.Location = New System.Drawing.Point(302, 119)
        Me.dtpSupplierInvoiceDate.Name = "dtpSupplierInvoiceDate"
        Me.dtpSupplierInvoiceDate.Size = New System.Drawing.Size(128, 20)
        Me.dtpSupplierInvoiceDate.TabIndex = 10
        '
        'txtPrimaryDPAccount
        '
        Me.txtPrimaryDPAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrimaryDPAccount.Location = New System.Drawing.Point(88, 202)
        Me.txtPrimaryDPAccount.MaxLength = 50
        Me.txtPrimaryDPAccount.Name = "txtPrimaryDPAccount"
        Me.txtPrimaryDPAccount.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryDPAccount.TabIndex = 23
        Me.txtPrimaryDPAccount.TabStop = False
        Me.txtPrimaryDPAccount.Text = "-1"
        Me.txtPrimaryDPAccount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryDPAccount.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(11, 206)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(65, 13)
        Me.Label12.TabIndex = 50
        Me.Label12.Text = "DP Account"
        '
        'cboDPAccount
        '
        Me.cboDPAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDPAccount.FormattingEnabled = True
        Me.cboDPAccount.Location = New System.Drawing.Point(115, 202)
        Me.cboDPAccount.Name = "cboDPAccount"
        Me.cboDPAccount.Size = New System.Drawing.Size(316, 21)
        Me.cboDPAccount.TabIndex = 14
        '
        'cboPrimaryDPAccount
        '
        Me.cboPrimaryDPAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryDPAccount.FormattingEnabled = True
        Me.cboPrimaryDPAccount.Location = New System.Drawing.Point(115, 202)
        Me.cboPrimaryDPAccount.Name = "cboPrimaryDPAccount"
        Me.cboPrimaryDPAccount.Size = New System.Drawing.Size(60, 21)
        Me.cboPrimaryDPAccount.TabIndex = 27
        '
        'lblCurrTax
        '
        Me.lblCurrTax.AutoSize = True
        Me.lblCurrTax.Location = New System.Drawing.Point(508, 99)
        Me.lblCurrTax.Name = "lblCurrTax"
        Me.lblCurrTax.Size = New System.Drawing.Size(26, 13)
        Me.lblCurrTax.TabIndex = 38
        Me.lblCurrTax.Text = "IDR"
        '
        'txtPrimaryCurrTax
        '
        Me.txtPrimaryCurrTax.Location = New System.Drawing.Point(445, 69)
        Me.txtPrimaryCurrTax.Name = "txtPrimaryCurrTax"
        Me.txtPrimaryCurrTax.Size = New System.Drawing.Size(15, 20)
        Me.txtPrimaryCurrTax.TabIndex = 32
        Me.txtPrimaryCurrTax.TabStop = False
        Me.txtPrimaryCurrTax.Text = "-1"
        Me.txtPrimaryCurrTax.Visible = False
        '
        'cboCurrTax
        '
        Me.cboCurrTax.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrTax.FormattingEnabled = True
        Me.cboCurrTax.Location = New System.Drawing.Point(480, 69)
        Me.cboCurrTax.Name = "cboCurrTax"
        Me.cboCurrTax.Size = New System.Drawing.Size(54, 21)
        Me.cboCurrTax.TabIndex = 18
        '
        'cboPrimaryCurrTax
        '
        Me.cboPrimaryCurrTax.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryCurrTax.FormattingEnabled = True
        Me.cboPrimaryCurrTax.Location = New System.Drawing.Point(486, 69)
        Me.cboPrimaryCurrTax.Name = "cboPrimaryCurrTax"
        Me.cboPrimaryCurrTax.Size = New System.Drawing.Size(42, 21)
        Me.cboPrimaryCurrTax.TabIndex = 34
        '
        'txtPrimaryAccount
        '
        Me.txtPrimaryAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrimaryAccount.Location = New System.Drawing.Point(87, 175)
        Me.txtPrimaryAccount.MaxLength = 50
        Me.txtPrimaryAccount.Name = "txtPrimaryAccount"
        Me.txtPrimaryAccount.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryAccount.TabIndex = 21
        Me.txtPrimaryAccount.TabStop = False
        Me.txtPrimaryAccount.Text = "-1"
        Me.txtPrimaryAccount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryAccount.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(10, 179)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 13)
        Me.Label10.TabIndex = 49
        Me.Label10.Text = "AP Account"
        '
        'cboAccount
        '
        Me.cboAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(114, 175)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(316, 21)
        Me.cboAccount.TabIndex = 13
        '
        'cboPrimaryAccount
        '
        Me.cboPrimaryAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryAccount.FormattingEnabled = True
        Me.cboPrimaryAccount.Location = New System.Drawing.Point(114, 175)
        Me.cboPrimaryAccount.Name = "cboPrimaryAccount"
        Me.cboPrimaryAccount.Size = New System.Drawing.Size(60, 21)
        Me.cboPrimaryAccount.TabIndex = 27
        '
        'Txt_E_Faktur_KodeJenisTransaksiPajak
        '
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Location = New System.Drawing.Point(82, 146)
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Name = "Txt_E_Faktur_KodeJenisTransaksiPajak"
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Size = New System.Drawing.Size(25, 20)
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.TabIndex = 17
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.TabStop = False
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Visible = False
        '
        'cboPPNDipungut
        '
        Me.cboPPNDipungut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPPNDipungut.FormattingEnabled = True
        Me.cboPPNDipungut.Items.AddRange(New Object() {"Collect Tax", "Don't Collect Tax"})
        Me.cboPPNDipungut.Location = New System.Drawing.Point(538, 69)
        Me.cboPPNDipungut.Name = "cboPPNDipungut"
        Me.cboPPNDipungut.Size = New System.Drawing.Size(111, 21)
        Me.cboPPNDipungut.TabIndex = 19
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(256, 150)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 13)
        Me.Label13.TabIndex = 19
        Me.Label13.Text = "NPWP"
        '
        'TxtNPWPNo
        '
        Me.TxtNPWPNo.Location = New System.Drawing.Point(302, 146)
        Me.TxtNPWPNo.MaxLength = 100
        Me.TxtNPWPNo.Name = "TxtNPWPNo"
        Me.TxtNPWPNo.Size = New System.Drawing.Size(128, 20)
        Me.TxtNPWPNo.TabIndex = 12
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(10, 150)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 13)
        Me.Label11.TabIndex = 48
        Me.Label11.Text = "Tax Invoice No"
        '
        'TxtTaxInvoiceNumber
        '
        Me.TxtTaxInvoiceNumber.Location = New System.Drawing.Point(114, 146)
        Me.TxtTaxInvoiceNumber.MaxLength = 100
        Me.TxtTaxInvoiceNumber.Name = "TxtTaxInvoiceNumber"
        Me.TxtTaxInvoiceNumber.Size = New System.Drawing.Size(137, 20)
        Me.TxtTaxInvoiceNumber.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 125)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 47
        Me.Label3.Text = "Supplier Invoice No"
        '
        'txtSupplierInvoiceNo
        '
        Me.txtSupplierInvoiceNo.Location = New System.Drawing.Point(114, 121)
        Me.txtSupplierInvoiceNo.MaxLength = 100
        Me.txtSupplierInvoiceNo.Name = "txtSupplierInvoiceNo"
        Me.txtSupplierInvoiceNo.Size = New System.Drawing.Size(137, 20)
        Me.txtSupplierInvoiceNo.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(271, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Date"
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrencyCode.Location = New System.Drawing.Point(114, 95)
        Me.txtCurrencyCode.MaxLength = 50
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.ReadOnly = True
        Me.txtCurrencyCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtCurrencyCode.Size = New System.Drawing.Size(101, 20)
        Me.txtCurrencyCode.TabIndex = 10
        Me.txtCurrencyCode.TabStop = False
        Me.txtCurrencyCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReadOnly.Location = New System.Drawing.Point(307, 95)
        Me.TxtReadOnly.MaxLength = 50
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(60, 20)
        Me.TxtReadOnly.TabIndex = 7
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(399, 99)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(31, 13)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "Days"
        '
        'TxtPaymentsTerms
        '
        Me.TxtPaymentsTerms.Location = New System.Drawing.Point(307, 95)
        Me.TxtPaymentsTerms.MaxLength = 50
        Me.TxtPaymentsTerms.Name = "TxtPaymentsTerms"
        Me.TxtPaymentsTerms.Size = New System.Drawing.Size(85, 20)
        Me.TxtPaymentsTerms.TabIndex = 8
        Me.TxtPaymentsTerms.Text = "0"
        Me.TxtPaymentsTerms.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(221, 99)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Payment Terms"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(447, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Stamp Duty"
        '
        'txtDutyStamp
        '
        Me.txtDutyStamp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDutyStamp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDutyStamp.Location = New System.Drawing.Point(538, 95)
        Me.txtDutyStamp.MaxLength = 50
        Me.txtDutyStamp.Name = "txtDutyStamp"
        Me.txtDutyStamp.Size = New System.Drawing.Size(324, 20)
        Me.txtDutyStamp.TabIndex = 21
        Me.txtDutyStamp.Text = "0"
        Me.txtDutyStamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnFindPO
        '
        Me.btnFindPO.Location = New System.Drawing.Point(398, 46)
        Me.btnFindPO.Name = "btnFindPO"
        Me.btnFindPO.Size = New System.Drawing.Size(32, 20)
        Me.btnFindPO.TabIndex = 3
        Me.btnFindPO.TabStop = False
        Me.btnFindPO.Text = "...."
        Me.btnFindPO.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 50)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(98, 13)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "Purchase Order No"
        '
        'txtPrimaryPO
        '
        Me.txtPrimaryPO.Location = New System.Drawing.Point(147, 46)
        Me.txtPrimaryPO.Name = "txtPrimaryPO"
        Me.txtPrimaryPO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryPO.TabIndex = 5
        Me.txtPrimaryPO.TabStop = False
        Me.txtPrimaryPO.Visible = False
        '
        'txtPONo
        '
        Me.txtPONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPONo.Location = New System.Drawing.Point(114, 46)
        Me.txtPONo.MaxLength = 50
        Me.txtPONo.Name = "txtPONo"
        Me.txtPONo.Size = New System.Drawing.Size(278, 20)
        Me.txtPONo.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 73)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 45
        Me.Label1.Text = "Supplier"
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(147, 69)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySupplier.TabIndex = 8
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Visible = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(114, 69)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.ReadOnly = True
        Me.txtSupplierCode.Size = New System.Drawing.Size(101, 20)
        Me.txtSupplierCode.TabIndex = 4
        Me.txtSupplierCode.TabStop = False
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Location = New System.Drawing.Point(224, 69)
        Me.TxtSupplierName.MaxLength = 150
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(206, 20)
        Me.TxtSupplierName.TabIndex = 5
        Me.TxtSupplierName.TabStop = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCurrency.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(145, 95)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(48, 20)
        Me.txtPrimaryCurrency.TabIndex = 11
        Me.txtPrimaryCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCurrency.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(10, 99)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(77, 13)
        Me.Label22.TabIndex = 46
        Me.Label22.Text = "Currency Code"
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(307, 18)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(123, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtPrimaryAOP
        '
        Me.TxtPrimaryAOP.Location = New System.Drawing.Point(151, 18)
        Me.TxtPrimaryAOP.Name = "TxtPrimaryAOP"
        Me.TxtPrimaryAOP.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryAOP.TabIndex = 1
        Me.TxtPrimaryAOP.TabStop = False
        Me.TxtPrimaryAOP.Visible = False
        '
        'TxtAOPNo
        '
        Me.TxtAOPNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAOPNo.Location = New System.Drawing.Point(114, 18)
        Me.TxtAOPNo.MaxLength = 50
        Me.TxtAOPNo.Name = "TxtAOPNo"
        Me.TxtAOPNo.ReadOnly = True
        Me.TxtAOPNo.Size = New System.Drawing.Size(151, 20)
        Me.TxtAOPNo.TabIndex = 0
        Me.TxtAOPNo.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "AOP DP No"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(443, 229)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 25
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(327, 229)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 24
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiAOPDP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(893, 319)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimumSize = New System.Drawing.Size(909, 357)
        Me.Name = "FrmTransaksiAOPDP"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Approval Of Purchase - Down Payment"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents txtTax As System.Windows.Forms.TextBox
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtTaxExchangeRate As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtExhangeRateToIDR As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtExchangeRateToUSD As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkTax As System.Windows.Forms.CheckBox
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryAOP As System.Windows.Forms.TextBox
    Friend WithEvents TxtAOPNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnFindPO As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents txtPONo As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDutyStamp As System.Windows.Forms.TextBox
    Friend WithEvents TxtPaymentsTerms As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWPNo As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtTaxInvoiceNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierInvoiceNo As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboPPNDipungut As System.Windows.Forms.ComboBox
    Friend WithEvents Txt_E_Faktur_KodeJenisTransaksiPajak As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cboAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryAccount As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryCurrTax As System.Windows.Forms.TextBox
    Friend WithEvents cboCurrTax As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryCurrTax As System.Windows.Forms.ComboBox
    Friend WithEvents lblCurrTax As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryDPAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cboDPAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryDPAccount As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dtpSupplierInvoiceDate As System.Windows.Forms.DateTimePicker
End Class
