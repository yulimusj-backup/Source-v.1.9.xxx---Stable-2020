<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesReturnBalanceManual
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
        Me.txtPrimaryAccount = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.cboAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryAR = New System.Windows.Forms.ComboBox
        Me.TxtExchangeRate = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.cboCurrency = New System.Windows.Forms.ComboBox
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Dtpduedate = New System.Windows.Forms.DateTimePicker
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtPaid = New System.Windows.Forms.TextBox
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtpInvoiceDate = New System.Windows.Forms.DateTimePicker
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtCustName = New System.Windows.Forms.TextBox
        Me.TxtPrimaryTransaksi = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.btnFindCust = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtInvoiceNumber = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCustCode = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtSales = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtVAT = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtDutystamp = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.txtPrimaryAccount)
        Me.gboDetail.Controls.Add(Me.Label13)
        Me.gboDetail.Controls.Add(Me.cboAccount)
        Me.gboDetail.Controls.Add(Me.cboPrimaryAR)
        Me.gboDetail.Controls.Add(Me.TxtExchangeRate)
        Me.gboDetail.Controls.Add(Me.Label14)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.cboCurrency)
        Me.gboDetail.Controls.Add(Me.TxtTotal)
        Me.gboDetail.Controls.Add(Me.txtReadOnly)
        Me.gboDetail.Controls.Add(Me.Label12)
        Me.gboDetail.Controls.Add(Me.Dtpduedate)
        Me.gboDetail.Controls.Add(Me.Label9)
        Me.gboDetail.Controls.Add(Me.TxtPaid)
        Me.gboDetail.Controls.Add(Me.TxtBalance)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.dtpInvoiceDate)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.TxtCustName)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryTransaksi)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryCustomer)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.btnFindCust)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtInvoiceNumber)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtCustCode)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(7, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(523, 284)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'txtPrimaryAccount
        '
        Me.txtPrimaryAccount.Location = New System.Drawing.Point(83, 180)
        Me.txtPrimaryAccount.Name = "txtPrimaryAccount"
        Me.txtPrimaryAccount.Size = New System.Drawing.Size(23, 20)
        Me.txtPrimaryAccount.TabIndex = 13
        Me.txtPrimaryAccount.TabStop = False
        Me.txtPrimaryAccount.Text = "-1"
        Me.txtPrimaryAccount.Visible = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 180)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(65, 13)
        Me.Label13.TabIndex = 28
        Me.Label13.Text = "AR Account"
        '
        'cboAccount
        '
        Me.cboAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(112, 177)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(402, 21)
        Me.cboAccount.TabIndex = 14
        '
        'cboPrimaryAR
        '
        Me.cboPrimaryAR.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryAR.FormattingEnabled = True
        Me.cboPrimaryAR.Location = New System.Drawing.Point(112, 177)
        Me.cboPrimaryAR.Name = "cboPrimaryAR"
        Me.cboPrimaryAR.Size = New System.Drawing.Size(66, 21)
        Me.cboPrimaryAR.TabIndex = 15
        '
        'TxtExchangeRate
        '
        Me.TxtExchangeRate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtExchangeRate.Location = New System.Drawing.Point(423, 149)
        Me.TxtExchangeRate.Name = "TxtExchangeRate"
        Me.TxtExchangeRate.Size = New System.Drawing.Size(91, 20)
        Me.TxtExchangeRate.TabIndex = 12
        Me.TxtExchangeRate.TabStop = False
        Me.TxtExchangeRate.Text = "1.0000"
        Me.TxtExchangeRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(336, 152)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(81, 13)
        Me.Label14.TabIndex = 11
        Me.Label14.Text = "Exchange Rate"
        '
        'TxtPrimaryCurrency
        '
        Me.TxtPrimaryCurrency.Location = New System.Drawing.Point(83, 150)
        Me.TxtPrimaryCurrency.Name = "TxtPrimaryCurrency"
        Me.TxtPrimaryCurrency.Size = New System.Drawing.Size(23, 20)
        Me.TxtPrimaryCurrency.TabIndex = 8
        Me.TxtPrimaryCurrency.TabStop = False
        Me.TxtPrimaryCurrency.Text = "-1"
        Me.TxtPrimaryCurrency.Visible = False
        '
        'cboCurrency
        '
        Me.cboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrency.FormattingEnabled = True
        Me.cboCurrency.Location = New System.Drawing.Point(112, 149)
        Me.cboCurrency.Name = "cboCurrency"
        Me.cboCurrency.Size = New System.Drawing.Size(68, 21)
        Me.cboCurrency.TabIndex = 9
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.Location = New System.Drawing.Point(10, 231)
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(132, 20)
        Me.TxtTotal.TabIndex = 19
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtTotal.Visible = False
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(83, 205)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 16
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(9, 129)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 13)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "Due Date"
        '
        'Dtpduedate
        '
        Me.Dtpduedate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Dtpduedate.CustomFormat = "dd MMM yyyy"
        Me.Dtpduedate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Dtpduedate.Location = New System.Drawing.Point(112, 123)
        Me.Dtpduedate.Name = "Dtpduedate"
        Me.Dtpduedate.Size = New System.Drawing.Size(402, 20)
        Me.Dtpduedate.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 152)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 13)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Balance"
        '
        'TxtPaid
        '
        Me.TxtPaid.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPaid.Location = New System.Drawing.Point(398, 205)
        Me.TxtPaid.MaxLength = 12
        Me.TxtPaid.Name = "TxtPaid"
        Me.TxtPaid.Size = New System.Drawing.Size(116, 20)
        Me.TxtPaid.TabIndex = 18
        Me.TxtPaid.TabStop = False
        Me.TxtPaid.Text = "0"
        Me.TxtPaid.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPaid.Visible = False
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.Location = New System.Drawing.Point(186, 149)
        Me.TxtBalance.MaxLength = 17
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.Size = New System.Drawing.Size(144, 20)
        Me.TxtBalance.TabIndex = 10
        Me.TxtBalance.Text = "0"
        Me.TxtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 103)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 13)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "SR Date"
        '
        'dtpInvoiceDate
        '
        Me.dtpInvoiceDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpInvoiceDate.CustomFormat = "dd MMM yyyy"
        Me.dtpInvoiceDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInvoiceDate.Location = New System.Drawing.Point(112, 97)
        Me.dtpInvoiceDate.Name = "dtpInvoiceDate"
        Me.dtpInvoiceDate.Size = New System.Drawing.Size(402, 20)
        Me.dtpInvoiceDate.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(82, 13)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Customer Name"
        '
        'TxtCustName
        '
        Me.TxtCustName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustName.Location = New System.Drawing.Point(112, 45)
        Me.TxtCustName.Name = "TxtCustName"
        Me.TxtCustName.ReadOnly = True
        Me.TxtCustName.Size = New System.Drawing.Size(402, 20)
        Me.TxtCustName.TabIndex = 4
        Me.TxtCustName.TabStop = False
        '
        'TxtPrimaryTransaksi
        '
        Me.TxtPrimaryTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryTransaksi.Location = New System.Drawing.Point(83, 45)
        Me.TxtPrimaryTransaksi.Name = "TxtPrimaryTransaksi"
        Me.TxtPrimaryTransaksi.Size = New System.Drawing.Size(23, 20)
        Me.TxtPrimaryTransaksi.TabIndex = 3
        Me.TxtPrimaryTransaksi.TabStop = False
        Me.TxtPrimaryTransaksi.Text = "-1"
        Me.TxtPrimaryTransaksi.Visible = False
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(84, 19)
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(23, 20)
        Me.TxtPrimaryCustomer.TabIndex = 0
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Text = "-1"
        Me.TxtPrimaryCustomer.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 208)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Remarks"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(112, 205)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(402, 20)
        Me.TxtRemarks.TabIndex = 17
        '
        'btnFindCust
        '
        Me.btnFindCust.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCust.Location = New System.Drawing.Point(472, 19)
        Me.btnFindCust.Name = "btnFindCust"
        Me.btnFindCust.Size = New System.Drawing.Size(42, 23)
        Me.btnFindCust.TabIndex = 2
        Me.btnFindCust.TabStop = False
        Me.btnFindCust.Text = "Find"
        Me.btnFindCust.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "SR Number"
        '
        'TxtInvoiceNumber
        '
        Me.TxtInvoiceNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInvoiceNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInvoiceNumber.Location = New System.Drawing.Point(112, 71)
        Me.TxtInvoiceNumber.MaxLength = 50
        Me.TxtInvoiceNumber.Name = "TxtInvoiceNumber"
        Me.TxtInvoiceNumber.Size = New System.Drawing.Size(402, 20)
        Me.TxtInvoiceNumber.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Customer Code"
        '
        'TxtCustCode
        '
        Me.TxtCustCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustCode.Location = New System.Drawing.Point(112, 19)
        Me.TxtCustCode.MaxLength = 50
        Me.TxtCustCode.Name = "TxtCustCode"
        Me.TxtCustCode.Size = New System.Drawing.Size(354, 20)
        Me.TxtCustCode.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(264, 252)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 21
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(148, 252)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 20
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtSales)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtVAT)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TxtDutystamp)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Location = New System.Drawing.Point(417, 290)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(299, 74)
        Me.GroupBox1.TabIndex = 202
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        Me.GroupBox1.Visible = False
        '
        'TxtSales
        '
        Me.TxtSales.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSales.Location = New System.Drawing.Point(116, 19)
        Me.TxtSales.MaxLength = 12
        Me.TxtSales.Name = "TxtSales"
        Me.TxtSales.Size = New System.Drawing.Size(501, 20)
        Me.TxtSales.TabIndex = 5
        Me.TxtSales.TabStop = False
        Me.TxtSales.Text = "0"
        Me.TxtSales.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Sales Return"
        '
        'TxtVAT
        '
        Me.TxtVAT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtVAT.Location = New System.Drawing.Point(116, 45)
        Me.TxtVAT.MaxLength = 12
        Me.TxtVAT.Name = "TxtVAT"
        Me.TxtVAT.Size = New System.Drawing.Size(501, 20)
        Me.TxtVAT.TabIndex = 6
        Me.TxtVAT.TabStop = False
        Me.TxtVAT.Text = "0"
        Me.TxtVAT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 13)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "VAT Return"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(13, 74)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 196
        Me.Label11.Text = "Duty stamp"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 100)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(31, 13)
        Me.Label6.TabIndex = 30
        Me.Label6.Text = "Total"
        '
        'TxtDutystamp
        '
        Me.TxtDutystamp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDutystamp.Location = New System.Drawing.Point(116, 71)
        Me.TxtDutystamp.MaxLength = 12
        Me.TxtDutystamp.Name = "TxtDutystamp"
        Me.TxtDutystamp.Size = New System.Drawing.Size(501, 20)
        Me.TxtDutystamp.TabIndex = 7
        Me.TxtDutystamp.TabStop = False
        Me.TxtDutystamp.Text = "0"
        Me.TxtDutystamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(13, 126)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(28, 13)
        Me.Label10.TabIndex = 38
        Me.Label10.Text = "Paid"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(419, 290)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(91, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(320, 290)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(91, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(221, 290)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(91, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(122, 290)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(91, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(23, 290)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(91, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
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
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 342)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(540, 22)
        Me.StatusStrip1.TabIndex = 35
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 320)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(540, 22)
        Me.StatusStrip2.TabIndex = 37
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
        'FrmTransaksiSalesReturnBalanceManual
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(540, 364)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(548, 353)
        Me.Name = "FrmTransaksiSalesReturnBalanceManual"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Historical Sales Return"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtPaid As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpInvoiceDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtCustName As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtVAT As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCust As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSales As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtInvoiceNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Dtpduedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtDutystamp As System.Windows.Forms.TextBox
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents cboCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents TxtExchangeRate As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryAR As System.Windows.Forms.ComboBox
End Class
