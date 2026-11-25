<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiQuotation
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryQuotation = New System.Windows.Forms.TextBox
        Me.TxtQuoNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupplierCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceByMultiUnit_Str = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryCont = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDown = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDownPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUpPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUseSecondary = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDownFromPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUpFromPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceByMultiUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.lblUnitsMulti = New System.Windows.Forms.Label
        Me.txtPrimaryUnitsMulti = New System.Windows.Forms.TextBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.txtAuthorizationStatus = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.llbAuthorizationDetail = New System.Windows.Forms.LinkLabel
        Me.txtPaymentTerm = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtPaymentDueDate = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.cboPaymentDueDateReference = New System.Windows.Forms.ComboBox
        Me.TxtInquiryNo = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtDPFromPercent = New System.Windows.Forms.TextBox
        Me.TxtDPAmount = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.chkUseDP = New System.Windows.Forms.CheckBox
        Me.TxtDPPercent = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtTaxAmount = New System.Windows.Forms.TextBox
        Me.txtSubTotal = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.chkTax = New System.Windows.Forms.CheckBox
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtNamaSigned = New System.Windows.Forms.TextBox
        Me.TxtJabatanSigned = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.RboAddressShipTo = New System.Windows.Forms.RadioButton
        Me.LlbViewAddressShipTo = New System.Windows.Forms.LinkLabel
        Me.TxtIdentifierAddressShipTo = New System.Windows.Forms.TextBox
        Me.CboAddressShipTo = New System.Windows.Forms.ComboBox
        Me.LlbViewAddressCustomerShipTo = New System.Windows.Forms.LinkLabel
        Me.TxtIdentifierAddressCustomerShipTo = New System.Windows.Forms.TextBox
        Me.CboAddressCustomerShipTo = New System.Windows.Forms.ComboBox
        Me.BtnFindCustomerShipTo = New System.Windows.Forms.Button
        Me.TxtPrimaryCustomerShipTo = New System.Windows.Forms.TextBox
        Me.TxtCustomerDescShipTo = New System.Windows.Forms.TextBox
        Me.RboCustomerShipTo = New System.Windows.Forms.RadioButton
        Me.TxtCustomerCodeShipTo = New System.Windows.Forms.TextBox
        Me.txtCustomerAddIdentifier = New System.Windows.Forms.TextBox
        Me.gboTerm = New System.Windows.Forms.GroupBox
        Me.txtTermsNCondition = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtNoNPWP = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtContactPerson = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.llbAddress = New System.Windows.Forms.LinkLabel
        Me.txtPhoneNo = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.cboCustomerAddress = New System.Windows.Forms.ComboBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.llbUpdateAllDeliveryDate = New System.Windows.Forms.LinkLabel
        Me.CboUnit = New System.Windows.Forms.ComboBox
        Me.TxtQtyperBatchMulti = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.llbDetailStock = New System.Windows.Forms.LinkLabel
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.chkUseCustCodeNo = New System.Windows.Forms.CheckBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.dtpdeliverydate = New System.Windows.Forms.DateTimePicker
        Me.btnFindCustomerProductCode = New System.Windows.Forms.Button
        Me.txtCustomerProductCodeNo = New System.Windows.Forms.TextBox
        Me.txtCustomerProductDesc = New System.Windows.Forms.TextBox
        Me.txtPriceByMultiUnit = New System.Windows.Forms.TextBox
        Me.LblUnitPrice = New System.Windows.Forms.LinkLabel
        Me.cboPriceByMultiUnit = New System.Windows.Forms.ComboBox
        Me.TxtDiscount = New System.Windows.Forms.TextBox
        Me.TxtUnitPrice = New System.Windows.Forms.TextBox
        Me.TxtDiscFromPerct = New System.Windows.Forms.TextBox
        Me.BtnInput = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtDiscAmount = New System.Windows.Forms.TextBox
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtNettPrice = New System.Windows.Forms.TextBox
        Me.txtAmount = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.chkUseSecondary = New System.Windows.Forms.CheckBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.lblUnitsDefault = New System.Windows.Forms.Label
        Me.txtPrimaryUnitsDefault = New System.Windows.Forms.TextBox
        Me.TxtQtyperBatchDefault = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.gboTerm.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Location = New System.Drawing.Point(385, 2)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(32, 20)
        Me.btnFindCustomer.TabIndex = 1
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "...."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Customer Code"
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(112, 2)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCustomer.TabIndex = 133
        Me.txtPrimaryCustomer.TabStop = False
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(112, 2)
        Me.txtCustomerCode.MaxLength = 50
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(267, 20)
        Me.txtCustomerCode.TabIndex = 0
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Location = New System.Drawing.Point(112, 26)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(305, 20)
        Me.TxtCustomerName.TabIndex = 2
        Me.TxtCustomerName.TabStop = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(133, 26)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(275, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtPrimaryQuotation
        '
        Me.TxtPrimaryQuotation.Location = New System.Drawing.Point(133, 2)
        Me.TxtPrimaryQuotation.Name = "TxtPrimaryQuotation"
        Me.TxtPrimaryQuotation.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryQuotation.TabIndex = 1
        Me.TxtPrimaryQuotation.TabStop = False
        Me.TxtPrimaryQuotation.Visible = False
        '
        'TxtQuoNo
        '
        Me.TxtQuoNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQuoNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtQuoNo.Location = New System.Drawing.Point(133, 2)
        Me.TxtQuoNo.MaxLength = 50
        Me.TxtQuoNo.Name = "TxtQuoNo"
        Me.TxtQuoNo.ReadOnly = True
        Me.TxtQuoNo.Size = New System.Drawing.Size(275, 20)
        Me.TxtQuoNo.TabIndex = 0
        Me.TxtQuoNo.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 30)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Quotation No"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(449, 474)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(333, 474)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 541)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(894, 22)
        Me.StatusStrip1.TabIndex = 39
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
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
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.lblUnitsMulti)
        Me.gboDetail.Controls.Add(Me.txtPrimaryUnitsMulti)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.lblUnitsDefault)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.txtPrimaryUnitsDefault)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtQtyperBatchDefault)
        Me.gboDetail.Location = New System.Drawing.Point(4, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(887, 506)
        Me.gboDetail.TabIndex = 6
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColSupplierCodeNo, Me.ColSuppDesc, Me.ColQtyMulti, Me.ColUnitsMulti, Me.ColRatioDesc, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColPriceByMultiUnit_Str, Me.ColDiscount, Me.ColDiscAmount, Me.ColNettPrice, Me.ColAmount, Me.ColDeliveryCont, Me.ColDeliveryDate, Me.ColRemark, Me.ColToleranceDown, Me.ColToleranceDownPerc, Me.ColToleranceUp, Me.ColToleranceUpPerc, Me.ColPrimaryCodeNo, Me.ColPrimaryUnits, Me.ColUseSecondary, Me.ColDiscFromPerct, Me.ColIdentifier, Me.ColPrimaryUnitsMulti, Me.ColRatioDefault, Me.ColRatioReference, Me.ColToleranceDownFromPerc, Me.ColToleranceUpFromPerc, Me.ColPriceByMultiUnit})
        Me.dgvItem.Location = New System.Drawing.Point(4, 319)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(878, 149)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 2
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.Frozen = True
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColSupplierCodeNo
        '
        Me.ColSupplierCodeNo.HeaderText = "Customer's Code No"
        Me.ColSupplierCodeNo.Name = "ColSupplierCodeNo"
        Me.ColSupplierCodeNo.Width = 128
        '
        'ColSuppDesc
        '
        Me.ColSuppDesc.HeaderText = "Customer's Description"
        Me.ColSuppDesc.Name = "ColSuppDesc"
        Me.ColSuppDesc.Width = 139
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColUnitsMulti
        '
        Me.ColUnitsMulti.HeaderText = "Unit"
        Me.ColUnitsMulti.Name = "ColUnitsMulti"
        Me.ColUnitsMulti.Width = 51
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRatioDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRatioDesc.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.Width = 57
        '
        'ColQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQty.HeaderText = "Qty Default"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 93
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.Width = 78
        '
        'ColPriceByMultiUnit_Str
        '
        Me.ColPriceByMultiUnit_Str.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColPriceByMultiUnit_Str.HeaderText = ""
        Me.ColPriceByMultiUnit_Str.Name = "ColPriceByMultiUnit_Str"
        Me.ColPriceByMultiUnit_Str.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColPriceByMultiUnit_Str.Width = 5
        '
        'ColDiscount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.Width = 91
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColDiscAmount.HeaderText = "Disc Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.Width = 92
        '
        'ColNettPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.Width = 79
        '
        'ColAmount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColDeliveryCont
        '
        Me.ColDeliveryCont.HeaderText = "Delivery Date Control"
        Me.ColDeliveryCont.Name = "ColDeliveryCont"
        Me.ColDeliveryCont.Width = 113
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDeliveryDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.Width = 96
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColToleranceDown
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceDown.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColToleranceDown.HeaderText = "Tolerance Down"
        Me.ColToleranceDown.Name = "ColToleranceDown"
        Me.ColToleranceDown.Width = 111
        '
        'ColToleranceDownPerc
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceDownPerc.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColToleranceDownPerc.HeaderText = "Tolerance Down (%)"
        Me.ColToleranceDownPerc.Name = "ColToleranceDownPerc"
        Me.ColToleranceDownPerc.Width = 128
        '
        'ColToleranceUp
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceUp.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColToleranceUp.HeaderText = "Tolerance Up"
        Me.ColToleranceUp.Name = "ColToleranceUp"
        Me.ColToleranceUp.Width = 97
        '
        'ColToleranceUpPerc
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceUpPerc.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColToleranceUpPerc.HeaderText = "Tolerance Up (%)"
        Me.ColToleranceUpPerc.Name = "ColToleranceUpPerc"
        Me.ColToleranceUpPerc.Width = 114
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 132
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 113
        '
        'ColUseSecondary
        '
        Me.ColUseSecondary.HeaderText = "ColUseSecondary"
        Me.ColUseSecondary.Name = "ColUseSecondary"
        Me.ColUseSecondary.Visible = False
        Me.ColUseSecondary.Width = 127
        '
        'ColDiscFromPerct
        '
        Me.ColDiscFromPerct.HeaderText = "Disc From Perct"
        Me.ColDiscFromPerct.Name = "ColDiscFromPerct"
        Me.ColDiscFromPerct.ReadOnly = True
        Me.ColDiscFromPerct.Visible = False
        Me.ColDiscFromPerct.Width = 133
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.ReadOnly = True
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 87
        '
        'ColPrimaryUnitsMulti
        '
        Me.ColPrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColPrimaryUnitsMulti.Name = "ColPrimaryUnitsMulti"
        Me.ColPrimaryUnitsMulti.Visible = False
        Me.ColPrimaryUnitsMulti.Width = 150
        '
        'ColRatioDefault
        '
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 115
        '
        'ColRatioReference
        '
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 136
        '
        'ColToleranceDownFromPerc
        '
        Me.ColToleranceDownFromPerc.HeaderText = "Tolerance Down From Perc"
        Me.ColToleranceDownFromPerc.Name = "ColToleranceDownFromPerc"
        Me.ColToleranceDownFromPerc.Visible = False
        Me.ColToleranceDownFromPerc.Width = 205
        '
        'ColToleranceUpFromPerc
        '
        Me.ColToleranceUpFromPerc.HeaderText = "Tolerance Up From Perc"
        Me.ColToleranceUpFromPerc.Name = "ColToleranceUpFromPerc"
        Me.ColToleranceUpFromPerc.Visible = False
        Me.ColToleranceUpFromPerc.Width = 188
        '
        'ColPriceByMultiUnit
        '
        Me.ColPriceByMultiUnit.HeaderText = "Price By Multi Unit"
        Me.ColPriceByMultiUnit.Name = "ColPriceByMultiUnit"
        Me.ColPriceByMultiUnit.Visible = False
        Me.ColPriceByMultiUnit.Width = 147
        '
        'lblUnitsMulti
        '
        Me.lblUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsMulti.AutoSize = True
        Me.lblUnitsMulti.Location = New System.Drawing.Point(47, 433)
        Me.lblUnitsMulti.Name = "lblUnitsMulti"
        Me.lblUnitsMulti.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsMulti.TabIndex = 88
        Me.lblUnitsMulti.Text = "Units"
        '
        'txtPrimaryUnitsMulti
        '
        Me.txtPrimaryUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryUnitsMulti.Location = New System.Drawing.Point(13, 429)
        Me.txtPrimaryUnitsMulti.Name = "txtPrimaryUnitsMulti"
        Me.txtPrimaryUnitsMulti.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryUnitsMulti.TabIndex = 89
        Me.txtPrimaryUnitsMulti.TabStop = False
        Me.txtPrimaryUnitsMulti.Text = "-1"
        Me.txtPrimaryUnitsMulti.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(4, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(878, 188)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.SplitContainer2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(870, 162)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data 1"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryQuotation)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtAuthorizationStatus)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label29)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.llbAuthorizationDetail)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPaymentTerm)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPaymentDueDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label18)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label20)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label17)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboPaymentDueDateReference)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtInquiryNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtQuoNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtReadOnly)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtDPFromPercent)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtDPAmount)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label22)
        Me.SplitContainer2.Panel2.Controls.Add(Me.chkUseDP)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtDPPercent)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtCurrencyCode)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnFindCurrency)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtTaxAmount)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtSubTotal)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label25)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.chkTax)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtTotal)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtTax)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label24)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtPrimaryCurrency)
        Me.SplitContainer2.Size = New System.Drawing.Size(864, 156)
        Me.SplitContainer2.SplitterDistance = 415
        Me.SplitContainer2.TabIndex = 7
        Me.SplitContainer2.TabStop = False
        '
        'txtAuthorizationStatus
        '
        Me.txtAuthorizationStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAuthorizationStatus.Location = New System.Drawing.Point(133, 51)
        Me.txtAuthorizationStatus.MaxLength = 150
        Me.txtAuthorizationStatus.Name = "txtAuthorizationStatus"
        Me.txtAuthorizationStatus.ReadOnly = True
        Me.txtAuthorizationStatus.Size = New System.Drawing.Size(239, 20)
        Me.txtAuthorizationStatus.TabIndex = 2
        Me.txtAuthorizationStatus.TabStop = False
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(9, 55)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(101, 13)
        Me.Label29.TabIndex = 165
        Me.Label29.Text = "Authorization Status"
        '
        'llbAuthorizationDetail
        '
        Me.llbAuthorizationDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAuthorizationDetail.AutoSize = True
        Me.llbAuthorizationDetail.Location = New System.Drawing.Point(379, 55)
        Me.llbAuthorizationDetail.Name = "llbAuthorizationDetail"
        Me.llbAuthorizationDetail.Size = New System.Drawing.Size(29, 13)
        Me.llbAuthorizationDetail.TabIndex = 2
        Me.llbAuthorizationDetail.TabStop = True
        Me.llbAuthorizationDetail.Text = "view"
        '
        'txtPaymentTerm
        '
        Me.txtPaymentTerm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentTerm.Location = New System.Drawing.Point(133, 100)
        Me.txtPaymentTerm.MaxLength = 50
        Me.txtPaymentTerm.Name = "txtPaymentTerm"
        Me.txtPaymentTerm.Size = New System.Drawing.Size(275, 20)
        Me.txtPaymentTerm.TabIndex = 5
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 13)
        Me.Label8.TabIndex = 178
        Me.Label8.Text = "Payment Term"
        '
        'txtPaymentDueDate
        '
        Me.txtPaymentDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentDueDate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentDueDate.Location = New System.Drawing.Point(133, 125)
        Me.txtPaymentDueDate.MaxLength = 4
        Me.txtPaymentDueDate.Name = "txtPaymentDueDate"
        Me.txtPaymentDueDate.Size = New System.Drawing.Size(47, 20)
        Me.txtPaymentDueDate.TabIndex = 6
        Me.txtPaymentDueDate.Text = "0"
        Me.txtPaymentDueDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(9, 129)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(97, 13)
        Me.Label18.TabIndex = 179
        Me.Label18.Text = "Payment Due Date"
        '
        'Label20
        '
        Me.Label20.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(187, 129)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(55, 13)
        Me.Label20.TabIndex = 8
        Me.Label20.Text = "days from "
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(9, 80)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(109, 13)
        Me.Label17.TabIndex = 183
        Me.Label17.Text = "Customer's Inquiry No"
        '
        'cboPaymentDueDateReference
        '
        Me.cboPaymentDueDateReference.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboPaymentDueDateReference.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaymentDueDateReference.FormattingEnabled = True
        Me.cboPaymentDueDateReference.Items.AddRange(New Object() {"Purchase Order Date", "Invoice Date"})
        Me.cboPaymentDueDateReference.Location = New System.Drawing.Point(249, 125)
        Me.cboPaymentDueDateReference.Name = "cboPaymentDueDateReference"
        Me.cboPaymentDueDateReference.Size = New System.Drawing.Size(159, 21)
        Me.cboPaymentDueDateReference.TabIndex = 7
        '
        'TxtInquiryNo
        '
        Me.TxtInquiryNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInquiryNo.Location = New System.Drawing.Point(133, 76)
        Me.TxtInquiryNo.MaxLength = 50
        Me.TxtInquiryNo.Name = "TxtInquiryNo"
        Me.TxtInquiryNo.Size = New System.Drawing.Size(275, 20)
        Me.TxtInquiryNo.TabIndex = 4
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(290, 51)
        Me.TxtReadOnly.MaxLength = 150
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(59, 20)
        Me.TxtReadOnly.TabIndex = 194
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtDPFromPercent
        '
        Me.TxtDPFromPercent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDPFromPercent.Cursor = System.Windows.Forms.Cursors.Default
        Me.TxtDPFromPercent.Location = New System.Drawing.Point(69, 79)
        Me.TxtDPFromPercent.Name = "TxtDPFromPercent"
        Me.TxtDPFromPercent.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TxtDPFromPercent.Size = New System.Drawing.Size(36, 20)
        Me.TxtDPFromPercent.TabIndex = 8
        Me.TxtDPFromPercent.TabStop = False
        Me.TxtDPFromPercent.Text = "1"
        Me.TxtDPFromPercent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtDPFromPercent.Visible = False
        '
        'TxtDPAmount
        '
        Me.TxtDPAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDPAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDPAmount.Location = New System.Drawing.Point(197, 77)
        Me.TxtDPAmount.MaxLength = 20
        Me.TxtDPAmount.Name = "TxtDPAmount"
        Me.TxtDPAmount.Size = New System.Drawing.Size(240, 20)
        Me.TxtDPAmount.TabIndex = 9
        Me.TxtDPAmount.Text = "0"
        Me.TxtDPAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(178, 81)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(15, 13)
        Me.Label22.TabIndex = 197
        Me.Label22.Text = "%"
        '
        'chkUseDP
        '
        Me.chkUseDP.AutoSize = True
        Me.chkUseDP.Location = New System.Drawing.Point(11, 79)
        Me.chkUseDP.Name = "chkUseDP"
        Me.chkUseDP.Size = New System.Drawing.Size(63, 17)
        Me.chkUseDP.TabIndex = 7
        Me.chkUseDP.Text = "Use DP"
        Me.chkUseDP.UseVisualStyleBackColor = True
        '
        'TxtDPPercent
        '
        Me.TxtDPPercent.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDPPercent.Location = New System.Drawing.Point(109, 77)
        Me.TxtDPPercent.MaxLength = 9
        Me.TxtDPPercent.Name = "TxtDPPercent"
        Me.TxtDPPercent.Size = New System.Drawing.Size(63, 20)
        Me.TxtDPPercent.TabIndex = 8
        Me.TxtDPPercent.Text = "0"
        Me.TxtDPPercent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 129)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 13)
        Me.Label6.TabIndex = 193
        Me.Label6.Text = "Signed By"
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrencyCode.Location = New System.Drawing.Point(109, 3)
        Me.txtCurrencyCode.MaxLength = 50
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtCurrencyCode.Size = New System.Drawing.Size(64, 20)
        Me.txtCurrencyCode.TabIndex = 0
        Me.txtCurrencyCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(11, 104)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(49, 13)
        Me.Label11.TabIndex = 191
        Me.Label11.Text = "Remarks"
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(109, 100)
        Me.txtRemarks.MaxLength = 50
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(329, 20)
        Me.txtRemarks.TabIndex = 10
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(178, 2)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(32, 20)
        Me.btnFindCurrency.TabIndex = 1
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "...."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'txtTaxAmount
        '
        Me.txtTaxAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount.Location = New System.Drawing.Point(197, 28)
        Me.txtTaxAmount.MaxLength = 20
        Me.txtTaxAmount.Name = "txtTaxAmount"
        Me.txtTaxAmount.ReadOnly = True
        Me.txtTaxAmount.Size = New System.Drawing.Size(240, 20)
        Me.txtTaxAmount.TabIndex = 5
        Me.txtTaxAmount.TabStop = False
        Me.txtTaxAmount.Text = "0"
        Me.txtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSubTotal.Location = New System.Drawing.Point(215, 3)
        Me.txtSubTotal.MaxLength = 20
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.Size = New System.Drawing.Size(222, 20)
        Me.txtSubTotal.TabIndex = 2
        Me.txtSubTotal.TabStop = False
        Me.txtSubTotal.Text = "0"
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(178, 32)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(15, 13)
        Me.Label25.TabIndex = 4
        Me.Label25.Text = "%"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(11, 7)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(53, 13)
        Me.Label15.TabIndex = 184
        Me.Label15.Text = "Sub Total"
        '
        'chkTax
        '
        Me.chkTax.AutoSize = True
        Me.chkTax.Location = New System.Drawing.Point(11, 30)
        Me.chkTax.Name = "chkTax"
        Me.chkTax.Size = New System.Drawing.Size(44, 17)
        Me.chkTax.TabIndex = 3
        Me.chkTax.Text = "Tax"
        Me.chkTax.UseVisualStyleBackColor = True
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(109, 54)
        Me.TxtTotal.MaxLength = 20
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(329, 20)
        Me.TxtTotal.TabIndex = 6
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTax
        '
        Me.txtTax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTax.Location = New System.Drawing.Point(109, 28)
        Me.txtTax.MaxLength = 9
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.Size = New System.Drawing.Size(63, 20)
        Me.txtTax.TabIndex = 4
        Me.txtTax.Text = "0"
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(11, 57)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(31, 13)
        Me.Label24.TabIndex = 185
        Me.Label24.Text = "Total"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(109, 117)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNamaSigned)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtJabatanSigned)
        Me.SplitContainer1.Size = New System.Drawing.Size(329, 36)
        Me.SplitContainer1.SplitterDistance = 154
        Me.SplitContainer1.TabIndex = 11
        Me.SplitContainer1.TabStop = False
        '
        'TxtNamaSigned
        '
        Me.TxtNamaSigned.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNamaSigned.Location = New System.Drawing.Point(0, 8)
        Me.TxtNamaSigned.MaxLength = 50
        Me.TxtNamaSigned.Name = "TxtNamaSigned"
        Me.TxtNamaSigned.Size = New System.Drawing.Size(151, 20)
        Me.TxtNamaSigned.TabIndex = 0
        '
        'TxtJabatanSigned
        '
        Me.TxtJabatanSigned.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJabatanSigned.Location = New System.Drawing.Point(4, 8)
        Me.TxtJabatanSigned.MaxLength = 50
        Me.TxtJabatanSigned.Name = "TxtJabatanSigned"
        Me.TxtJabatanSigned.Size = New System.Drawing.Size(163, 20)
        Me.TxtJabatanSigned.TabIndex = 0
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCurrency.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(69, 3)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(36, 20)
        Me.txtPrimaryCurrency.TabIndex = 11
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCurrency.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.RboAddressShipTo)
        Me.TabPage2.Controls.Add(Me.LlbViewAddressShipTo)
        Me.TabPage2.Controls.Add(Me.TxtIdentifierAddressShipTo)
        Me.TabPage2.Controls.Add(Me.CboAddressShipTo)
        Me.TabPage2.Controls.Add(Me.LlbViewAddressCustomerShipTo)
        Me.TabPage2.Controls.Add(Me.TxtIdentifierAddressCustomerShipTo)
        Me.TabPage2.Controls.Add(Me.CboAddressCustomerShipTo)
        Me.TabPage2.Controls.Add(Me.BtnFindCustomerShipTo)
        Me.TabPage2.Controls.Add(Me.TxtPrimaryCustomerShipTo)
        Me.TabPage2.Controls.Add(Me.TxtCustomerDescShipTo)
        Me.TabPage2.Controls.Add(Me.RboCustomerShipTo)
        Me.TabPage2.Controls.Add(Me.TxtCustomerCodeShipTo)
        Me.TabPage2.Controls.Add(Me.txtCustomerAddIdentifier)
        Me.TabPage2.Controls.Add(Me.gboTerm)
        Me.TabPage2.Controls.Add(Me.TxtCustomerName)
        Me.TabPage2.Controls.Add(Me.txtPrimaryCustomer)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.txtNoNPWP)
        Me.TabPage2.Controls.Add(Me.btnFindCustomer)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.txtContactPerson)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.llbAddress)
        Me.TabPage2.Controls.Add(Me.txtPhoneNo)
        Me.TabPage2.Controls.Add(Me.Label23)
        Me.TabPage2.Controls.Add(Me.cboCustomerAddress)
        Me.TabPage2.Controls.Add(Me.txtCustomerCode)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(870, 162)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Data 2"
        '
        'RboAddressShipTo
        '
        Me.RboAddressShipTo.AutoSize = True
        Me.RboAddressShipTo.Checked = True
        Me.RboAddressShipTo.Location = New System.Drawing.Point(440, 4)
        Me.RboAddressShipTo.Name = "RboAddressShipTo"
        Me.RboAddressShipTo.Size = New System.Drawing.Size(99, 17)
        Me.RboAddressShipTo.TabIndex = 209
        Me.RboAddressShipTo.TabStop = True
        Me.RboAddressShipTo.Text = "Ship to Address"
        Me.RboAddressShipTo.UseVisualStyleBackColor = True
        '
        'LlbViewAddressShipTo
        '
        Me.LlbViewAddressShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbViewAddressShipTo.AutoSize = True
        Me.LlbViewAddressShipTo.Location = New System.Drawing.Point(833, 6)
        Me.LlbViewAddressShipTo.Name = "LlbViewAddressShipTo"
        Me.LlbViewAddressShipTo.Size = New System.Drawing.Size(29, 13)
        Me.LlbViewAddressShipTo.TabIndex = 211
        Me.LlbViewAddressShipTo.TabStop = True
        Me.LlbViewAddressShipTo.Text = "view"
        '
        'TxtIdentifierAddressShipTo
        '
        Me.TxtIdentifierAddressShipTo.Location = New System.Drawing.Point(546, 2)
        Me.TxtIdentifierAddressShipTo.Name = "TxtIdentifierAddressShipTo"
        Me.TxtIdentifierAddressShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtIdentifierAddressShipTo.TabIndex = 212
        Me.TxtIdentifierAddressShipTo.TabStop = False
        Me.TxtIdentifierAddressShipTo.Text = "-1"
        Me.TxtIdentifierAddressShipTo.Visible = False
        '
        'CboAddressShipTo
        '
        Me.CboAddressShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAddressShipTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAddressShipTo.FormattingEnabled = True
        Me.CboAddressShipTo.Location = New System.Drawing.Point(546, 2)
        Me.CboAddressShipTo.Name = "CboAddressShipTo"
        Me.CboAddressShipTo.Size = New System.Drawing.Size(282, 21)
        Me.CboAddressShipTo.TabIndex = 210
        '
        'LlbViewAddressCustomerShipTo
        '
        Me.LlbViewAddressCustomerShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbViewAddressCustomerShipTo.AutoSize = True
        Me.LlbViewAddressCustomerShipTo.Location = New System.Drawing.Point(832, 57)
        Me.LlbViewAddressCustomerShipTo.Name = "LlbViewAddressCustomerShipTo"
        Me.LlbViewAddressCustomerShipTo.Size = New System.Drawing.Size(29, 13)
        Me.LlbViewAddressCustomerShipTo.TabIndex = 215
        Me.LlbViewAddressCustomerShipTo.TabStop = True
        Me.LlbViewAddressCustomerShipTo.Text = "view"
        '
        'TxtIdentifierAddressCustomerShipTo
        '
        Me.TxtIdentifierAddressCustomerShipTo.Location = New System.Drawing.Point(545, 54)
        Me.TxtIdentifierAddressCustomerShipTo.Name = "TxtIdentifierAddressCustomerShipTo"
        Me.TxtIdentifierAddressCustomerShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtIdentifierAddressCustomerShipTo.TabIndex = 208
        Me.TxtIdentifierAddressCustomerShipTo.TabStop = False
        Me.TxtIdentifierAddressCustomerShipTo.Text = "-1"
        Me.TxtIdentifierAddressCustomerShipTo.Visible = False
        '
        'CboAddressCustomerShipTo
        '
        Me.CboAddressCustomerShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAddressCustomerShipTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAddressCustomerShipTo.FormattingEnabled = True
        Me.CboAddressCustomerShipTo.Location = New System.Drawing.Point(545, 54)
        Me.CboAddressCustomerShipTo.Name = "CboAddressCustomerShipTo"
        Me.CboAddressCustomerShipTo.Size = New System.Drawing.Size(282, 21)
        Me.CboAddressCustomerShipTo.TabIndex = 214
        '
        'BtnFindCustomerShipTo
        '
        Me.BtnFindCustomerShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCustomerShipTo.Location = New System.Drawing.Point(617, 29)
        Me.BtnFindCustomerShipTo.Name = "BtnFindCustomerShipTo"
        Me.BtnFindCustomerShipTo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCustomerShipTo.TabIndex = 203
        Me.BtnFindCustomerShipTo.TabStop = False
        Me.BtnFindCustomerShipTo.Text = "...."
        Me.BtnFindCustomerShipTo.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCustomerShipTo
        '
        Me.TxtPrimaryCustomerShipTo.Location = New System.Drawing.Point(545, 29)
        Me.TxtPrimaryCustomerShipTo.Name = "TxtPrimaryCustomerShipTo"
        Me.TxtPrimaryCustomerShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCustomerShipTo.TabIndex = 207
        Me.TxtPrimaryCustomerShipTo.TabStop = False
        Me.TxtPrimaryCustomerShipTo.Text = "-1"
        Me.TxtPrimaryCustomerShipTo.Visible = False
        '
        'TxtCustomerDescShipTo
        '
        Me.TxtCustomerDescShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerDescShipTo.Location = New System.Drawing.Point(656, 29)
        Me.TxtCustomerDescShipTo.MaxLength = 150
        Me.TxtCustomerDescShipTo.Name = "TxtCustomerDescShipTo"
        Me.TxtCustomerDescShipTo.ReadOnly = True
        Me.TxtCustomerDescShipTo.Size = New System.Drawing.Size(205, 20)
        Me.TxtCustomerDescShipTo.TabIndex = 204
        Me.TxtCustomerDescShipTo.TabStop = False
        '
        'RboCustomerShipTo
        '
        Me.RboCustomerShipTo.AutoSize = True
        Me.RboCustomerShipTo.Location = New System.Drawing.Point(439, 31)
        Me.RboCustomerShipTo.Name = "RboCustomerShipTo"
        Me.RboCustomerShipTo.Size = New System.Drawing.Size(105, 17)
        Me.RboCustomerShipTo.TabIndex = 212
        Me.RboCustomerShipTo.Text = "Ship to Customer"
        Me.RboCustomerShipTo.UseVisualStyleBackColor = True
        '
        'TxtCustomerCodeShipTo
        '
        Me.TxtCustomerCodeShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerCodeShipTo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCodeShipTo.Location = New System.Drawing.Point(545, 29)
        Me.TxtCustomerCodeShipTo.MaxLength = 50
        Me.TxtCustomerCodeShipTo.Name = "TxtCustomerCodeShipTo"
        Me.TxtCustomerCodeShipTo.Size = New System.Drawing.Size(66, 20)
        Me.TxtCustomerCodeShipTo.TabIndex = 213
        '
        'txtCustomerAddIdentifier
        '
        Me.txtCustomerAddIdentifier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerAddIdentifier.Location = New System.Drawing.Point(112, 51)
        Me.txtCustomerAddIdentifier.MaxLength = 50
        Me.txtCustomerAddIdentifier.Name = "txtCustomerAddIdentifier"
        Me.txtCustomerAddIdentifier.Size = New System.Drawing.Size(25, 20)
        Me.txtCustomerAddIdentifier.TabIndex = 161
        Me.txtCustomerAddIdentifier.Visible = False
        '
        'gboTerm
        '
        Me.gboTerm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboTerm.Controls.Add(Me.txtTermsNCondition)
        Me.gboTerm.Location = New System.Drawing.Point(436, 76)
        Me.gboTerm.Name = "gboTerm"
        Me.gboTerm.Size = New System.Drawing.Size(419, 68)
        Me.gboTerm.TabIndex = 216
        Me.gboTerm.TabStop = False
        Me.gboTerm.Text = "Terms && Condition"
        '
        'txtTermsNCondition
        '
        Me.txtTermsNCondition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtTermsNCondition.Location = New System.Drawing.Point(3, 16)
        Me.txtTermsNCondition.MaxLength = 255
        Me.txtTermsNCondition.Multiline = True
        Me.txtTermsNCondition.Name = "txtTermsNCondition"
        Me.txtTermsNCondition.Size = New System.Drawing.Size(413, 49)
        Me.txtTermsNCondition.TabIndex = 216
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(9, 129)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(57, 13)
        Me.Label10.TabIndex = 159
        Me.Label10.Text = "NPWP No"
        '
        'txtNoNPWP
        '
        Me.txtNoNPWP.Location = New System.Drawing.Point(112, 125)
        Me.txtNoNPWP.MaxLength = 50
        Me.txtNoNPWP.Name = "txtNoNPWP"
        Me.txtNoNPWP.Size = New System.Drawing.Size(305, 20)
        Me.txtNoNPWP.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "Customer Name"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 13)
        Me.Label5.TabIndex = 136
        Me.Label5.Text = "Address Code"
        '
        'txtContactPerson
        '
        Me.txtContactPerson.Location = New System.Drawing.Point(112, 76)
        Me.txtContactPerson.MaxLength = 50
        Me.txtContactPerson.Name = "txtContactPerson"
        Me.txtContactPerson.Size = New System.Drawing.Size(305, 20)
        Me.txtContactPerson.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 138
        Me.Label7.Text = "Contact Person"
        '
        'llbAddress
        '
        Me.llbAddress.AutoSize = True
        Me.llbAddress.Location = New System.Drawing.Point(388, 55)
        Me.llbAddress.Name = "llbAddress"
        Me.llbAddress.Size = New System.Drawing.Size(29, 13)
        Me.llbAddress.TabIndex = 4
        Me.llbAddress.TabStop = True
        Me.llbAddress.Text = "view"
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.Location = New System.Drawing.Point(112, 100)
        Me.txtPhoneNo.MaxLength = 50
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.Size = New System.Drawing.Size(305, 20)
        Me.txtPhoneNo.TabIndex = 6
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(9, 104)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(55, 13)
        Me.Label23.TabIndex = 141
        Me.Label23.Text = "Phone No"
        '
        'cboCustomerAddress
        '
        Me.cboCustomerAddress.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerAddress.FormattingEnabled = True
        Me.cboCustomerAddress.Location = New System.Drawing.Point(112, 51)
        Me.cboCustomerAddress.Name = "cboCustomerAddress"
        Me.cboCustomerAddress.Size = New System.Drawing.Size(267, 21)
        Me.cboCustomerAddress.TabIndex = 3
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.SplitContainer3)
        Me.gboItem.Location = New System.Drawing.Point(4, 200)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(878, 116)
        Me.gboItem.TabIndex = 1
        Me.gboItem.TabStop = False
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.llbUpdateAllDeliveryDate)
        Me.SplitContainer3.Panel1.Controls.Add(Me.CboUnit)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtQtyperBatchMulti)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.BtnFindCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.llbDetailStock)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label27)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label26)
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtPrimaryCategory)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtDesc)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer3.Panel1.Controls.Add(Me.chkUseCustCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer3.Panel1.Controls.Add(Me.dtpdeliverydate)
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnFindCustomerProductCode)
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtCustomerProductCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtCustomerProductDesc)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.txtPriceByMultiUnit)
        Me.SplitContainer3.Panel2.Controls.Add(Me.LblUnitPrice)
        Me.SplitContainer3.Panel2.Controls.Add(Me.cboPriceByMultiUnit)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtDiscount)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtUnitPrice)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtDiscFromPerct)
        Me.SplitContainer3.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label21)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtDiscAmount)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtRemarkItem)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label19)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer3.Panel2.Controls.Add(Me.txtNettPrice)
        Me.SplitContainer3.Panel2.Controls.Add(Me.txtAmount)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label28)
        Me.SplitContainer3.Panel2.Controls.Add(Me.chkUseSecondary)
        Me.SplitContainer3.Size = New System.Drawing.Size(872, 97)
        Me.SplitContainer3.SplitterDistance = 420
        Me.SplitContainer3.TabIndex = 7
        Me.SplitContainer3.TabStop = False
        '
        'llbUpdateAllDeliveryDate
        '
        Me.llbUpdateAllDeliveryDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUpdateAllDeliveryDate.AutoSize = True
        Me.llbUpdateAllDeliveryDate.Location = New System.Drawing.Point(302, 53)
        Me.llbUpdateAllDeliveryDate.Name = "llbUpdateAllDeliveryDate"
        Me.llbUpdateAllDeliveryDate.Size = New System.Drawing.Size(116, 13)
        Me.llbUpdateAllDeliveryDate.TabIndex = 167
        Me.llbUpdateAllDeliveryDate.TabStop = True
        Me.llbUpdateAllDeliveryDate.Text = "update all delivery date"
        Me.llbUpdateAllDeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'CboUnit
        '
        Me.CboUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboUnit.FormattingEnabled = True
        Me.CboUnit.Location = New System.Drawing.Point(348, 72)
        Me.CboUnit.Name = "CboUnit"
        Me.CboUnit.Size = New System.Drawing.Size(70, 21)
        Me.CboUnit.TabIndex = 90
        '
        'TxtQtyperBatchMulti
        '
        Me.TxtQtyperBatchMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchMulti.Location = New System.Drawing.Point(142, 73)
        Me.TxtQtyperBatchMulti.MaxLength = 18
        Me.TxtQtyperBatchMulti.Name = "TxtQtyperBatchMulti"
        Me.TxtQtyperBatchMulti.Size = New System.Drawing.Size(202, 20)
        Me.TxtQtyperBatchMulti.TabIndex = 88
        Me.TxtQtyperBatchMulti.Text = "0"
        Me.TxtQtyperBatchMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(142, 2)
        Me.TxtCodeNo.MaxLength = 50
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(175, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(320, 2)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCodeNo.TabIndex = 1
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "...."
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'llbDetailStock
        '
        Me.llbDetailStock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbDetailStock.AutoSize = True
        Me.llbDetailStock.Location = New System.Drawing.Point(357, 6)
        Me.llbDetailStock.Name = "llbDetailStock"
        Me.llbDetailStock.Size = New System.Drawing.Size(61, 13)
        Me.llbDetailStock.TabIndex = 3
        Me.llbDetailStock.TabStop = True
        Me.llbDetailStock.Text = "show stock"
        Me.ToolTip1.SetToolTip(Me.llbDetailStock, "Click here to see stock condition")
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(11, 77)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(23, 13)
        Me.Label27.TabIndex = 63
        Me.Label27.Text = "Qty"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(11, 53)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(71, 13)
        Me.Label26.TabIndex = 154
        Me.Label26.Text = "Delivery Date"
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(179, 2)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryCategory.TabIndex = 110
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Text = "-1"
        Me.txtPrimaryCategory.Visible = False
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(148, 2)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 108
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Text = "-1"
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(142, 25)
        Me.txtDesc.MaxLength = 150
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(276, 20)
        Me.txtDesc.TabIndex = 3
        Me.txtDesc.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(11, 6)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(49, 13)
        Me.Label13.TabIndex = 107
        Me.Label13.Text = "Code No"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(11, 29)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(60, 13)
        Me.Label12.TabIndex = 109
        Me.Label12.Text = "Description"
        '
        'chkUseCustCodeNo
        '
        Me.chkUseCustCodeNo.AutoSize = True
        Me.chkUseCustCodeNo.Location = New System.Drawing.Point(11, 51)
        Me.chkUseCustCodeNo.Name = "chkUseCustCodeNo"
        Me.chkUseCustCodeNo.Size = New System.Drawing.Size(122, 17)
        Me.chkUseCustCodeNo.TabIndex = 4
        Me.chkUseCustCodeNo.TabStop = False
        Me.chkUseCustCodeNo.Text = "Customer's Code No"
        Me.chkUseCustCodeNo.UseVisualStyleBackColor = True
        Me.chkUseCustCodeNo.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(11, 77)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(114, 13)
        Me.Label9.TabIndex = 161
        Me.Label9.Text = "Customer's Description"
        Me.Label9.Visible = False
        '
        'dtpdeliverydate
        '
        Me.dtpdeliverydate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpdeliverydate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpdeliverydate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpdeliverydate.Location = New System.Drawing.Point(142, 49)
        Me.dtpdeliverydate.Name = "dtpdeliverydate"
        Me.dtpdeliverydate.ShowCheckBox = True
        Me.dtpdeliverydate.Size = New System.Drawing.Size(154, 20)
        Me.dtpdeliverydate.TabIndex = 8
        '
        'btnFindCustomerProductCode
        '
        Me.btnFindCustomerProductCode.Location = New System.Drawing.Point(382, 49)
        Me.btnFindCustomerProductCode.Name = "btnFindCustomerProductCode"
        Me.btnFindCustomerProductCode.Size = New System.Drawing.Size(40, 20)
        Me.btnFindCustomerProductCode.TabIndex = 6
        Me.btnFindCustomerProductCode.TabStop = False
        Me.btnFindCustomerProductCode.Text = "...."
        Me.btnFindCustomerProductCode.UseVisualStyleBackColor = True
        Me.btnFindCustomerProductCode.Visible = False
        '
        'txtCustomerProductCodeNo
        '
        Me.txtCustomerProductCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerProductCodeNo.Location = New System.Drawing.Point(142, 49)
        Me.txtCustomerProductCodeNo.MaxLength = 50
        Me.txtCustomerProductCodeNo.Name = "txtCustomerProductCodeNo"
        Me.txtCustomerProductCodeNo.Size = New System.Drawing.Size(234, 20)
        Me.txtCustomerProductCodeNo.TabIndex = 5
        Me.txtCustomerProductCodeNo.TabStop = False
        Me.txtCustomerProductCodeNo.Visible = False
        '
        'txtCustomerProductDesc
        '
        Me.txtCustomerProductDesc.Location = New System.Drawing.Point(142, 73)
        Me.txtCustomerProductDesc.MaxLength = 150
        Me.txtCustomerProductDesc.Name = "txtCustomerProductDesc"
        Me.txtCustomerProductDesc.ReadOnly = True
        Me.txtCustomerProductDesc.Size = New System.Drawing.Size(280, 20)
        Me.txtCustomerProductDesc.TabIndex = 7
        Me.txtCustomerProductDesc.TabStop = False
        Me.txtCustomerProductDesc.Visible = False
        '
        'txtPriceByMultiUnit
        '
        Me.txtPriceByMultiUnit.Location = New System.Drawing.Point(121, 27)
        Me.txtPriceByMultiUnit.Name = "txtPriceByMultiUnit"
        Me.txtPriceByMultiUnit.Size = New System.Drawing.Size(28, 20)
        Me.txtPriceByMultiUnit.TabIndex = 189
        Me.txtPriceByMultiUnit.TabStop = False
        Me.txtPriceByMultiUnit.Text = "1"
        Me.txtPriceByMultiUnit.Visible = False
        '
        'LblUnitPrice
        '
        Me.LblUnitPrice.AutoSize = True
        Me.LblUnitPrice.Location = New System.Drawing.Point(6, 6)
        Me.LblUnitPrice.Name = "LblUnitPrice"
        Me.LblUnitPrice.Size = New System.Drawing.Size(84, 13)
        Me.LblUnitPrice.TabIndex = 10
        Me.LblUnitPrice.TabStop = True
        Me.LblUnitPrice.Text = "Unit Price (/unit)"
        '
        'cboPriceByMultiUnit
        '
        Me.cboPriceByMultiUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPriceByMultiUnit.FormattingEnabled = True
        Me.cboPriceByMultiUnit.Location = New System.Drawing.Point(64, 2)
        Me.cboPriceByMultiUnit.Name = "cboPriceByMultiUnit"
        Me.cboPriceByMultiUnit.Size = New System.Drawing.Size(85, 21)
        Me.cboPriceByMultiUnit.TabIndex = 188
        '
        'TxtDiscount
        '
        Me.TxtDiscount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDiscount.Location = New System.Drawing.Point(320, 2)
        Me.TxtDiscount.MaxLength = 9
        Me.TxtDiscount.Name = "TxtDiscount"
        Me.TxtDiscount.Size = New System.Drawing.Size(122, 20)
        Me.TxtDiscount.TabIndex = 12
        Me.TxtDiscount.Text = "0"
        Me.TxtDiscount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtUnitPrice
        '
        Me.TxtUnitPrice.Location = New System.Drawing.Point(154, 2)
        Me.TxtUnitPrice.MaxLength = 14
        Me.TxtUnitPrice.Name = "TxtUnitPrice"
        Me.TxtUnitPrice.Size = New System.Drawing.Size(93, 20)
        Me.TxtUnitPrice.TabIndex = 11
        Me.TxtUnitPrice.Text = "0"
        Me.TxtUnitPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtDiscFromPerct
        '
        Me.TxtDiscFromPerct.Location = New System.Drawing.Point(285, 2)
        Me.TxtDiscFromPerct.Name = "TxtDiscFromPerct"
        Me.TxtDiscFromPerct.Size = New System.Drawing.Size(28, 20)
        Me.TxtDiscFromPerct.TabIndex = 187
        Me.TxtDiscFromPerct.TabStop = False
        Me.TxtDiscFromPerct.Text = "1"
        Me.TxtDiscFromPerct.Visible = False
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(394, 72)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(48, 22)
        Me.BtnInput.TabIndex = 17
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 29)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(70, 13)
        Me.Label14.TabIndex = 186
        Me.Label14.Text = "Disc. Amount"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(6, 77)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(49, 13)
        Me.Label21.TabIndex = 63
        Me.Label21.Text = "Remarks"
        '
        'TxtDiscAmount
        '
        Me.TxtDiscAmount.Location = New System.Drawing.Point(154, 25)
        Me.TxtDiscAmount.MaxLength = 18
        Me.TxtDiscAmount.Name = "TxtDiscAmount"
        Me.TxtDiscAmount.Size = New System.Drawing.Size(93, 20)
        Me.TxtDiscAmount.TabIndex = 13
        Me.TxtDiscAmount.Text = "0"
        Me.TxtDiscAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(154, 73)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(236, 20)
        Me.TxtRemarkItem.TabIndex = 16
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(252, 6)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(66, 13)
        Me.Label19.TabIndex = 11
        Me.Label19.Text = "Discount (%)"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(6, 53)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(43, 13)
        Me.Label16.TabIndex = 101
        Me.Label16.Text = "Amount"
        '
        'txtNettPrice
        '
        Me.txtNettPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNettPrice.Location = New System.Drawing.Point(320, 25)
        Me.txtNettPrice.MaxLength = 12
        Me.txtNettPrice.Name = "txtNettPrice"
        Me.txtNettPrice.ReadOnly = True
        Me.txtNettPrice.Size = New System.Drawing.Size(122, 20)
        Me.txtNettPrice.TabIndex = 14
        Me.txtNettPrice.TabStop = False
        Me.txtNettPrice.Text = "0"
        Me.txtNettPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAmount
        '
        Me.txtAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAmount.Location = New System.Drawing.Point(154, 49)
        Me.txtAmount.MaxLength = 12
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.ReadOnly = True
        Me.txtAmount.Size = New System.Drawing.Size(288, 20)
        Me.txtAmount.TabIndex = 15
        Me.txtAmount.TabStop = False
        Me.txtAmount.Text = "0"
        Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(252, 28)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(54, 13)
        Me.Label28.TabIndex = 164
        Me.Label28.Text = "Nett Price"
        '
        'chkUseSecondary
        '
        Me.chkUseSecondary.AutoSize = True
        Me.chkUseSecondary.Location = New System.Drawing.Point(112, 51)
        Me.chkUseSecondary.Name = "chkUseSecondary"
        Me.chkUseSecondary.Size = New System.Drawing.Size(99, 17)
        Me.chkUseSecondary.TabIndex = 165
        Me.chkUseSecondary.Text = "Use Secondary"
        Me.chkUseSecondary.UseVisualStyleBackColor = True
        Me.chkUseSecondary.Visible = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Sales_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 474)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 6
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Sales_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 474)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 5
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'lblUnitsDefault
        '
        Me.lblUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsDefault.AutoSize = True
        Me.lblUnitsDefault.Location = New System.Drawing.Point(47, 412)
        Me.lblUnitsDefault.Name = "lblUnitsDefault"
        Me.lblUnitsDefault.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsDefault.TabIndex = 10
        Me.lblUnitsDefault.Text = "Units"
        '
        'txtPrimaryUnitsDefault
        '
        Me.txtPrimaryUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryUnitsDefault.Location = New System.Drawing.Point(13, 408)
        Me.txtPrimaryUnitsDefault.Name = "txtPrimaryUnitsDefault"
        Me.txtPrimaryUnitsDefault.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryUnitsDefault.TabIndex = 87
        Me.txtPrimaryUnitsDefault.TabStop = False
        Me.txtPrimaryUnitsDefault.Text = "-1"
        Me.txtPrimaryUnitsDefault.Visible = False
        '
        'TxtQtyperBatchDefault
        '
        Me.TxtQtyperBatchDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchDefault.Location = New System.Drawing.Point(84, 408)
        Me.TxtQtyperBatchDefault.MaxLength = 18
        Me.TxtQtyperBatchDefault.Name = "TxtQtyperBatchDefault"
        Me.TxtQtyperBatchDefault.Size = New System.Drawing.Size(67, 20)
        Me.TxtQtyperBatchDefault.TabIndex = 9
        Me.TxtQtyperBatchDefault.Text = "0"
        Me.TxtQtyperBatchDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(278, 512)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(536, 512)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(622, 512)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(192, 512)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(450, 512)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(364, 512)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'FrmTransaksiQuotation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(894, 563)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(900, 591)
        Me.Name = "FrmTransaksiQuotation"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quotation"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.gboTerm.ResumeLayout(False)
        Me.gboTerm.PerformLayout()
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryQuotation As System.Windows.Forms.TextBox
    Friend WithEvents TxtQuoNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents llbAddress As System.Windows.Forms.LinkLabel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtContactPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboCustomerAddress As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents txtCustomerProductDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomerProductCode As System.Windows.Forms.Button
    Friend WithEvents txtCustomerProductCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents dtpdeliverydate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtDiscount As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TxtUnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitsDefault As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsDefault As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchDefault As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents txtNettPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents chkUseCustCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtNoNPWP As System.Windows.Forms.TextBox
    Friend WithEvents llbAuthorizationDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents txtAuthorizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkTax As System.Windows.Forms.CheckBox
    Friend WithEvents txtTax As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtInquiryNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents cboPaymentDueDateReference As System.Windows.Forms.ComboBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtPaymentDueDate As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPaymentTerm As System.Windows.Forms.TextBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents gboTerm As System.Windows.Forms.GroupBox
    Friend WithEvents txtTermsNCondition As System.Windows.Forms.TextBox
    Friend WithEvents llbDetailStock As System.Windows.Forms.LinkLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents txtCustomerAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents chkUseSecondary As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtNamaSigned As System.Windows.Forms.TextBox
    Friend WithEvents TxtJabatanSigned As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents LblUnitPrice As System.Windows.Forms.LinkLabel
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtDiscAmount As System.Windows.Forms.TextBox
    Friend WithEvents TxtDiscFromPerct As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtQtyperBatchMulti As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsMulti As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryUnitsMulti As System.Windows.Forms.TextBox
    Friend WithEvents CboUnit As System.Windows.Forms.ComboBox
    Friend WithEvents llbUpdateAllDeliveryDate As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtDPFromPercent As System.Windows.Forms.TextBox
    Friend WithEvents TxtDPAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents chkUseDP As System.Windows.Forms.CheckBox
    Friend WithEvents TxtDPPercent As System.Windows.Forms.TextBox
    Friend WithEvents cboPriceByMultiUnit As System.Windows.Forms.ComboBox
    Friend WithEvents txtPriceByMultiUnit As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupplierCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceByMultiUnit_Str As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryCont As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDown As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDownPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUpPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUseSecondary As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDownFromPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUpFromPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceByMultiUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LlbViewAddressCustomerShipTo As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtIdentifierAddressCustomerShipTo As System.Windows.Forms.TextBox
    Friend WithEvents CboAddressCustomerShipTo As System.Windows.Forms.ComboBox
    Friend WithEvents BtnFindCustomerShipTo As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCustomerShipTo As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerDescShipTo As System.Windows.Forms.TextBox
    Friend WithEvents RboCustomerShipTo As System.Windows.Forms.RadioButton
    Friend WithEvents TxtCustomerCodeShipTo As System.Windows.Forms.TextBox
    Friend WithEvents RboAddressShipTo As System.Windows.Forms.RadioButton
    Friend WithEvents LlbViewAddressShipTo As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtIdentifierAddressShipTo As System.Windows.Forms.TextBox
    Friend WithEvents CboAddressShipTo As System.Windows.Forms.ComboBox
End Class
