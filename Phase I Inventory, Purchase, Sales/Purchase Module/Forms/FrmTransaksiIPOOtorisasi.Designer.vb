<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiIPOOtorisasi
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUseSecondary = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Label27 = New System.Windows.Forms.Label
        Me.txtCustPoNo = New System.Windows.Forms.TextBox
        Me.dtpETD = New System.Windows.Forms.DateTimePicker
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtDeliveryTime = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtPrimaryBrand = New System.Windows.Forms.TextBox
        Me.TxtBrand = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtCustomerContactName = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtCustName = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtCustCode = New System.Windows.Forms.TextBox
        Me.txtPrimarySO = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtSalesOrderNo = New System.Windows.Forms.TextBox
        Me.txtTaxAmount = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.chkTax = New System.Windows.Forms.CheckBox
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtSubTotal = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtPhoneNo = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtContactPerson = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.llbAuthorizationDetail = New System.Windows.Forms.LinkLabel
        Me.txtAuthorizationStatus = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryIPO = New System.Windows.Forms.TextBox
        Me.TxtIPONo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.llbUncheckAllAuthorize = New System.Windows.Forms.LinkLabel
        Me.llbCheckAllAuthorize = New System.Windows.Forms.LinkLabel
        Me.chkBelumOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.llbUncheckAllCancel = New System.Windows.Forms.LinkLabel
        Me.llbCheckAllCancel = New System.Windows.Forms.LinkLabel
        Me.chkSudahOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbExit = New System.Windows.Forms.ToolStripButton
        Me.tsbAuthorize = New System.Windows.Forms.ToolStripButton
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.Label27)
        Me.gboDetail.Controls.Add(Me.txtCustPoNo)
        Me.gboDetail.Controls.Add(Me.dtpETD)
        Me.gboDetail.Controls.Add(Me.Label26)
        Me.gboDetail.Controls.Add(Me.Label20)
        Me.gboDetail.Controls.Add(Me.txtDeliveryTime)
        Me.gboDetail.Controls.Add(Me.Label19)
        Me.gboDetail.Controls.Add(Me.txtPrimaryBrand)
        Me.gboDetail.Controls.Add(Me.TxtBrand)
        Me.gboDetail.Controls.Add(Me.Label18)
        Me.gboDetail.Controls.Add(Me.txtCustomerContactName)
        Me.gboDetail.Controls.Add(Me.Label17)
        Me.gboDetail.Controls.Add(Me.TxtCustName)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.txtCustCode)
        Me.gboDetail.Controls.Add(Me.txtPrimarySO)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.txtSalesOrderNo)
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
        Me.gboDetail.Controls.Add(Me.Label23)
        Me.gboDetail.Controls.Add(Me.txtPhoneNo)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.txtContactPerson)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtPrimarySupplier)
        Me.gboDetail.Controls.Add(Me.txtSupplierCode)
        Me.gboDetail.Controls.Add(Me.TxtSupplierName)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.Label22)
        Me.gboDetail.Controls.Add(Me.txtCurrencyCode)
        Me.gboDetail.Controls.Add(Me.llbAuthorizationDetail)
        Me.gboDetail.Controls.Add(Me.txtAuthorizationStatus)
        Me.gboDetail.Controls.Add(Me.Label21)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryIPO)
        Me.gboDetail.Controls.Add(Me.TxtIPONo)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Location = New System.Drawing.Point(12, 149)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(860, 459)
        Me.gboDetail.TabIndex = 2
        Me.gboDetail.TabStop = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColDiscount, Me.ColNettPrice, Me.ColAmount, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryUnits, Me.ColUseSecondary})
        Me.dgvItem.Location = New System.Drawing.Point(7, 254)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(840, 199)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 237
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.Frozen = True
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.ReadOnly = True
        Me.ColDescription.Width = 85
        '
        'ColQty
        '
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.Width = 56
        '
        'ColUnitPrice
        '
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.Width = 78
        '
        'ColDiscount
        '
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.ReadOnly = True
        Me.ColDiscount.Width = 91
        '
        'ColNettPrice
        '
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.Width = 79
        '
        'ColAmount
        '
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.Width = 68
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.ReadOnly = True
        Me.ColRemark.Width = 74
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.ReadOnly = True
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'ColUseSecondary
        '
        Me.ColUseSecondary.HeaderText = "Column1"
        Me.ColUseSecondary.Name = "ColUseSecondary"
        Me.ColUseSecondary.ReadOnly = True
        Me.ColUseSecondary.Visible = False
        Me.ColUseSecondary.Width = 54
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(231, 179)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(66, 13)
        Me.Label27.TabIndex = 236
        Me.Label27.Text = "Cust. PO No"
        '
        'txtCustPoNo
        '
        Me.txtCustPoNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustPoNo.Location = New System.Drawing.Point(303, 176)
        Me.txtCustPoNo.MaxLength = 50
        Me.txtCustPoNo.Name = "txtCustPoNo"
        Me.txtCustPoNo.ReadOnly = True
        Me.txtCustPoNo.Size = New System.Drawing.Size(125, 20)
        Me.txtCustPoNo.TabIndex = 196
        Me.txtCustPoNo.TabStop = False
        '
        'dtpETD
        '
        Me.dtpETD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpETD.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpETD.Enabled = False
        Me.dtpETD.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpETD.Location = New System.Drawing.Point(703, 44)
        Me.dtpETD.Name = "dtpETD"
        Me.dtpETD.Size = New System.Drawing.Size(143, 20)
        Me.dtpETD.TabIndex = 201
        '
        'Label26
        '
        Me.Label26.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(631, 48)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(66, 13)
        Me.Label26.TabIndex = 235
        Me.Label26.Text = "days     ETD"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(450, 49)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(71, 13)
        Me.Label20.TabIndex = 234
        Me.Label20.Text = "Delivery Time"
        '
        'txtDeliveryTime
        '
        Me.txtDeliveryTime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDeliveryTime.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDeliveryTime.Location = New System.Drawing.Point(565, 46)
        Me.txtDeliveryTime.MaxLength = 50
        Me.txtDeliveryTime.Name = "txtDeliveryTime"
        Me.txtDeliveryTime.ReadOnly = True
        Me.txtDeliveryTime.Size = New System.Drawing.Size(60, 20)
        Me.txtDeliveryTime.TabIndex = 200
        Me.txtDeliveryTime.Text = "0"
        Me.txtDeliveryTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(8, 50)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(35, 13)
        Me.Label19.TabIndex = 232
        Me.Label19.Text = "Brand"
        '
        'txtPrimaryBrand
        '
        Me.txtPrimaryBrand.Location = New System.Drawing.Point(145, 45)
        Me.txtPrimaryBrand.Name = "txtPrimaryBrand"
        Me.txtPrimaryBrand.ReadOnly = True
        Me.txtPrimaryBrand.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryBrand.TabIndex = 233
        Me.txtPrimaryBrand.TabStop = False
        Me.txtPrimaryBrand.Visible = False
        '
        'TxtBrand
        '
        Me.TxtBrand.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBrand.Location = New System.Drawing.Point(109, 45)
        Me.TxtBrand.MaxLength = 50
        Me.TxtBrand.Name = "TxtBrand"
        Me.TxtBrand.ReadOnly = True
        Me.TxtBrand.Size = New System.Drawing.Size(319, 20)
        Me.TxtBrand.TabIndex = 186
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(8, 231)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(91, 13)
        Me.Label18.TabIndex = 231
        Me.Label18.Text = "Customer Contact"
        '
        'txtCustomerContactName
        '
        Me.txtCustomerContactName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerContactName.Location = New System.Drawing.Point(109, 228)
        Me.txtCustomerContactName.MaxLength = 150
        Me.txtCustomerContactName.Name = "txtCustomerContactName"
        Me.txtCustomerContactName.ReadOnly = True
        Me.txtCustomerContactName.Size = New System.Drawing.Size(319, 20)
        Me.txtCustomerContactName.TabIndex = 198
        Me.txtCustomerContactName.TabStop = False
        Me.txtCustomerContactName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(8, 205)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(82, 13)
        Me.Label17.TabIndex = 230
        Me.Label17.Text = "Customer Name"
        '
        'TxtCustName
        '
        Me.TxtCustName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustName.Location = New System.Drawing.Point(109, 202)
        Me.TxtCustName.MaxLength = 150
        Me.TxtCustName.Name = "TxtCustName"
        Me.TxtCustName.ReadOnly = True
        Me.TxtCustName.Size = New System.Drawing.Size(319, 20)
        Me.TxtCustName.TabIndex = 197
        Me.TxtCustName.TabStop = False
        Me.TxtCustName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 179)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(79, 13)
        Me.Label5.TabIndex = 229
        Me.Label5.Text = "Customer Code"
        '
        'txtCustCode
        '
        Me.txtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustCode.Location = New System.Drawing.Point(109, 176)
        Me.txtCustCode.MaxLength = 50
        Me.txtCustCode.Name = "txtCustCode"
        Me.txtCustCode.ReadOnly = True
        Me.txtCustCode.Size = New System.Drawing.Size(116, 20)
        Me.txtCustCode.TabIndex = 195
        Me.txtCustCode.TabStop = False
        Me.txtCustCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPrimarySO
        '
        Me.txtPrimarySO.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimarySO.Location = New System.Drawing.Point(109, 150)
        Me.txtPrimarySO.Name = "txtPrimarySO"
        Me.txtPrimarySO.ReadOnly = True
        Me.txtPrimarySO.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimarySO.Size = New System.Drawing.Size(39, 20)
        Me.txtPrimarySO.TabIndex = 228
        Me.txtPrimarySO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimarySO.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 154)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(79, 13)
        Me.Label8.TabIndex = 227
        Me.Label8.Text = "Sales Order No"
        '
        'txtSalesOrderNo
        '
        Me.txtSalesOrderNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSalesOrderNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSalesOrderNo.Location = New System.Drawing.Point(109, 150)
        Me.txtSalesOrderNo.MaxLength = 50
        Me.txtSalesOrderNo.Name = "txtSalesOrderNo"
        Me.txtSalesOrderNo.ReadOnly = True
        Me.txtSalesOrderNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtSalesOrderNo.Size = New System.Drawing.Size(319, 20)
        Me.txtSalesOrderNo.TabIndex = 193
        Me.txtSalesOrderNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTaxAmount
        '
        Me.txtTaxAmount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount.Location = New System.Drawing.Point(636, 121)
        Me.txtTaxAmount.MaxLength = 50
        Me.txtTaxAmount.Name = "txtTaxAmount"
        Me.txtTaxAmount.ReadOnly = True
        Me.txtTaxAmount.Size = New System.Drawing.Size(210, 20)
        Me.txtTaxAmount.TabIndex = 205
        Me.txtTaxAmount.TabStop = False
        Me.txtTaxAmount.Text = "0"
        Me.txtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(617, 124)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(15, 13)
        Me.Label25.TabIndex = 226
        Me.Label25.Text = "%"
        '
        'chkTax
        '
        Me.chkTax.AutoSize = True
        Me.chkTax.Location = New System.Drawing.Point(453, 123)
        Me.chkTax.Name = "chkTax"
        Me.chkTax.Size = New System.Drawing.Size(85, 17)
        Me.chkTax.TabIndex = 208
        Me.chkTax.Text = "Tax Exclude"
        Me.chkTax.UseVisualStyleBackColor = True
        '
        'txtTax
        '
        Me.txtTax.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTax.Location = New System.Drawing.Point(565, 121)
        Me.txtTax.MaxLength = 50
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.Size = New System.Drawing.Size(46, 20)
        Me.txtTax.TabIndex = 210
        Me.txtTax.Text = "0"
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(450, 149)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(31, 13)
        Me.Label24.TabIndex = 225
        Me.Label24.Text = "Total"
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(565, 146)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(281, 20)
        Me.TxtTotal.TabIndex = 206
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(450, 99)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(53, 13)
        Me.Label15.TabIndex = 224
        Me.Label15.Text = "Sub Total"
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSubTotal.Location = New System.Drawing.Point(565, 96)
        Me.txtSubTotal.MaxLength = 50
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.Size = New System.Drawing.Size(281, 20)
        Me.txtSubTotal.TabIndex = 204
        Me.txtSubTotal.TabStop = False
        Me.txtSubTotal.Text = "0"
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(565, 198)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.ReadOnly = True
        Me.TxtRemark.Size = New System.Drawing.Size(279, 50)
        Me.TxtRemark.TabIndex = 211
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(450, 205)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 223
        Me.Label6.Text = "Remarks"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(243, 126)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(55, 13)
        Me.Label23.TabIndex = 222
        Me.Label23.Text = "Phone No"
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPhoneNo.Location = New System.Drawing.Point(303, 123)
        Me.txtPhoneNo.MaxLength = 50
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.ReadOnly = True
        Me.txtPhoneNo.Size = New System.Drawing.Size(125, 20)
        Me.txtPhoneNo.TabIndex = 192
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 127)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(85, 13)
        Me.Label7.TabIndex = 221
        Me.Label7.Text = "Supplier Contact"
        '
        'txtContactPerson
        '
        Me.txtContactPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtContactPerson.Location = New System.Drawing.Point(109, 124)
        Me.txtContactPerson.MaxLength = 50
        Me.txtContactPerson.Name = "txtContactPerson"
        Me.txtContactPerson.ReadOnly = True
        Me.txtContactPerson.Size = New System.Drawing.Size(116, 20)
        Me.txtContactPerson.TabIndex = 191
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 220
        Me.Label2.Text = "Supplier Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 76)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 218
        Me.Label1.Text = "Supplier Code"
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(145, 71)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.ReadOnly = True
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySupplier.TabIndex = 219
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Visible = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(109, 71)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.ReadOnly = True
        Me.txtSupplierCode.Size = New System.Drawing.Size(319, 20)
        Me.txtSupplierCode.TabIndex = 188
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Location = New System.Drawing.Point(109, 98)
        Me.TxtSupplierName.MaxLength = 150
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(319, 20)
        Me.TxtSupplierName.TabIndex = 190
        Me.TxtSupplierName.TabStop = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCurrency.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(565, 70)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.ReadOnly = True
        Me.txtPrimaryCurrency.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(22, 20)
        Me.txtPrimaryCurrency.TabIndex = 217
        Me.txtPrimaryCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCurrency.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(450, 74)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(77, 13)
        Me.Label22.TabIndex = 216
        Me.Label22.Text = "Currency Code"
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrencyCode.Location = New System.Drawing.Point(565, 70)
        Me.txtCurrencyCode.MaxLength = 50
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.ReadOnly = True
        Me.txtCurrencyCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtCurrencyCode.Size = New System.Drawing.Size(281, 20)
        Me.txtCurrencyCode.TabIndex = 202
        Me.txtCurrencyCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'llbAuthorizationDetail
        '
        Me.llbAuthorizationDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAuthorizationDetail.AutoSize = True
        Me.llbAuthorizationDetail.Location = New System.Drawing.Point(793, 175)
        Me.llbAuthorizationDetail.Name = "llbAuthorizationDetail"
        Me.llbAuthorizationDetail.Size = New System.Drawing.Size(57, 13)
        Me.llbAuthorizationDetail.TabIndex = 209
        Me.llbAuthorizationDetail.TabStop = True
        Me.llbAuthorizationDetail.Text = "view detail"
        '
        'txtAuthorizationStatus
        '
        Me.txtAuthorizationStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAuthorizationStatus.Location = New System.Drawing.Point(565, 172)
        Me.txtAuthorizationStatus.MaxLength = 50
        Me.txtAuthorizationStatus.Name = "txtAuthorizationStatus"
        Me.txtAuthorizationStatus.ReadOnly = True
        Me.txtAuthorizationStatus.Size = New System.Drawing.Size(222, 20)
        Me.txtAuthorizationStatus.TabIndex = 207
        Me.txtAuthorizationStatus.TabStop = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(450, 175)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(101, 13)
        Me.Label21.TabIndex = 215
        Me.Label21.Text = "Authorization Status"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(565, 19)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(281, 20)
        Me.dtpDate.TabIndex = 199
        '
        'TxtPrimaryIPO
        '
        Me.TxtPrimaryIPO.Location = New System.Drawing.Point(149, 19)
        Me.TxtPrimaryIPO.Name = "TxtPrimaryIPO"
        Me.TxtPrimaryIPO.ReadOnly = True
        Me.TxtPrimaryIPO.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryIPO.TabIndex = 214
        Me.TxtPrimaryIPO.TabStop = False
        Me.TxtPrimaryIPO.Visible = False
        '
        'TxtIPONo
        '
        Me.TxtIPONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtIPONo.Location = New System.Drawing.Point(109, 19)
        Me.TxtIPONo.MaxLength = 50
        Me.TxtIPONo.Name = "TxtIPONo"
        Me.TxtIPONo.ReadOnly = True
        Me.TxtIPONo.Size = New System.Drawing.Size(319, 20)
        Me.TxtIPONo.TabIndex = 185
        Me.TxtIPONo.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(450, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 13)
        Me.Label3.TabIndex = 213
        Me.Label3.Text = "Internal PO Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 13)
        Me.Label4.TabIndex = 212
        Me.Label4.Text = "Internal PO No"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(860, 131)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.llbUncheckAllAuthorize)
        Me.TabPage1.Controls.Add(Me.llbCheckAllAuthorize)
        Me.TabPage1.Controls.Add(Me.chkBelumOtorisasi)
        Me.TabPage1.Controls.Add(Me.btnExit)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(852, 105)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Authorize"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'llbUncheckAllAuthorize
        '
        Me.llbUncheckAllAuthorize.AutoSize = True
        Me.llbUncheckAllAuthorize.Location = New System.Drawing.Point(66, 89)
        Me.llbUncheckAllAuthorize.Name = "llbUncheckAllAuthorize"
        Me.llbUncheckAllAuthorize.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAllAuthorize.TabIndex = 47
        Me.llbUncheckAllAuthorize.TabStop = True
        Me.llbUncheckAllAuthorize.Text = "Uncheck All"
        '
        'llbCheckAllAuthorize
        '
        Me.llbCheckAllAuthorize.AutoSize = True
        Me.llbCheckAllAuthorize.Location = New System.Drawing.Point(6, 89)
        Me.llbCheckAllAuthorize.Name = "llbCheckAllAuthorize"
        Me.llbCheckAllAuthorize.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAllAuthorize.TabIndex = 46
        Me.llbCheckAllAuthorize.TabStop = True
        Me.llbCheckAllAuthorize.Text = "Check All"
        '
        'chkBelumOtorisasi
        '
        Me.chkBelumOtorisasi.ColumnWidth = 180
        Me.chkBelumOtorisasi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkBelumOtorisasi.FormattingEnabled = True
        Me.chkBelumOtorisasi.Location = New System.Drawing.Point(3, 3)
        Me.chkBelumOtorisasi.MultiColumn = True
        Me.chkBelumOtorisasi.Name = "chkBelumOtorisasi"
        Me.chkBelumOtorisasi.Size = New System.Drawing.Size(846, 94)
        Me.chkBelumOtorisasi.TabIndex = 0
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(701, 36)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(62, 29)
        Me.btnExit.TabIndex = 43
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.llbUncheckAllCancel)
        Me.TabPage2.Controls.Add(Me.llbCheckAllCancel)
        Me.TabPage2.Controls.Add(Me.chkSudahOtorisasi)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(852, 105)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Cancel Authorize"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'llbUncheckAllCancel
        '
        Me.llbUncheckAllCancel.AutoSize = True
        Me.llbUncheckAllCancel.Location = New System.Drawing.Point(66, 89)
        Me.llbUncheckAllCancel.Name = "llbUncheckAllCancel"
        Me.llbUncheckAllCancel.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAllCancel.TabIndex = 5
        Me.llbUncheckAllCancel.TabStop = True
        Me.llbUncheckAllCancel.Text = "Uncheck All"
        '
        'llbCheckAllCancel
        '
        Me.llbCheckAllCancel.AutoSize = True
        Me.llbCheckAllCancel.Location = New System.Drawing.Point(6, 89)
        Me.llbCheckAllCancel.Name = "llbCheckAllCancel"
        Me.llbCheckAllCancel.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAllCancel.TabIndex = 4
        Me.llbCheckAllCancel.TabStop = True
        Me.llbCheckAllCancel.Text = "Check All"
        '
        'chkSudahOtorisasi
        '
        Me.chkSudahOtorisasi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkSudahOtorisasi.FormattingEnabled = True
        Me.chkSudahOtorisasi.Location = New System.Drawing.Point(3, 3)
        Me.chkSudahOtorisasi.Name = "chkSudahOtorisasi"
        Me.chkSudahOtorisasi.Size = New System.Drawing.Size(846, 94)
        Me.chkSudahOtorisasi.TabIndex = 1
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.tsbAuthorize})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 633)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip1.Size = New System.Drawing.Size(884, 30)
        Me.ToolStrip1.TabIndex = 3
        '
        'tsbExit
        '
        Me.tsbExit.Image = Global.Purchase_Module.My.Resources.Resources.cancel1
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(59, 27)
        Me.tsbExit.Text = "E&xit"
        Me.tsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbAuthorize
        '
        Me.tsbAuthorize.Image = Global.Purchase_Module.My.Resources.Resources.lock_open1
        Me.tsbAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAuthorize.Name = "tsbAuthorize"
        Me.tsbAuthorize.Size = New System.Drawing.Size(92, 27)
        Me.tsbAuthorize.Text = "Authorize"
        Me.tsbAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 611)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip2.TabIndex = 53
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'FrmTransaksiIPOOtorisasi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(884, 663)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 699)
        Me.Name = "FrmTransaksiIPOOtorisasi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Internal Purchase Order Authorization Manager"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents chkBelumOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents chkSudahOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents llbUncheckAllAuthorize As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAllAuthorize As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAllCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAllCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents txtCustPoNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpETD As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtDeliveryTime As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryBrand As System.Windows.Forms.TextBox
    Friend WithEvents TxtBrand As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerContactName As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtCustName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySO As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtSalesOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents txtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkTax As System.Windows.Forms.CheckBox
    Friend WithEvents txtTax As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtContactPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents llbAuthorizationDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents txtAuthorizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryIPO As System.Windows.Forms.TextBox
    Friend WithEvents TxtIPONo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUseSecondary As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
