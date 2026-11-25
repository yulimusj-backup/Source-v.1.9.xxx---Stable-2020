<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO3
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
        Me.txtCustomerAddIdentifier = New System.Windows.Forms.TextBox
        Me.chkFakturPajak = New System.Windows.Forms.CheckBox
        Me.chkKwitansi = New System.Windows.Forms.CheckBox
        Me.llbAddress = New System.Windows.Forms.LinkLabel
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboCustomerAddress = New System.Windows.Forms.ComboBox
        Me.DtpExpired = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.chkExpiryDateControlled = New System.Windows.Forms.CheckBox
        Me.chkBatchNoControlled = New System.Windows.Forms.CheckBox
        Me.chkUseSecondaryUnit = New System.Windows.Forms.CheckBox
        Me.chkUseCustCodeNo = New System.Windows.Forms.CheckBox
        Me.cboCustCodeNo = New System.Windows.Forms.ComboBox
        Me.txtCustProductDesc = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.BtnInput = New System.Windows.Forms.Button
        Me.cboSalesOrderNo = New System.Windows.Forms.ComboBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.lblCodeNo = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cboCodeNo = New System.Windows.Forms.ComboBox
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.llbAddBooking = New System.Windows.Forms.LinkLabel
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkPdoControlled = New System.Windows.Forms.CheckBox
        Me.cboSecondaryQty = New System.Windows.Forms.ComboBox
        Me.lblSecQty = New System.Windows.Forms.Label
        Me.txtSecondaryQty = New System.Windows.Forms.TextBox
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtToleranceUp = New System.Windows.Forms.TextBox
        Me.ChkIgnoreTolerance = New System.Windows.Forms.CheckBox
        Me.cboCustPONO = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtQtyOrderBalance = New System.Windows.Forms.TextBox
        Me.llbDetailStock = New System.Windows.Forms.LinkLabel
        Me.LblUnitOB = New System.Windows.Forms.Label
        Me.txtPrimarySO = New System.Windows.Forms.TextBox
        Me.lblCustPONo = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.lblUnitsPendingDO = New System.Windows.Forms.Label
        Me.lblSalesOrder = New System.Windows.Forms.Label
        Me.txtPendingDO = New System.Windows.Forms.TextBox
        Me.LlbQty = New System.Windows.Forms.LinkLabel
        Me.TxtRatioDefault = New System.Windows.Forms.TextBox
        Me.TxtRatioReference = New System.Windows.Forms.TextBox
        Me.TxtQtyperBatchDefault = New System.Windows.Forms.TextBox
        Me.lblAvailableUnits = New System.Windows.Forms.Label
        Me.lblAvailableUnitsMulti = New System.Windows.Forms.Label
        Me.TxtAvailableMulti = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.lblUnitsMulti = New System.Windows.Forms.Label
        Me.lblUnitsDefault = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.TxtQtyperBatchMulti = New System.Windows.Forms.TextBox
        Me.txtAvailable = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitsDefault = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitsMulti = New System.Windows.Forms.TextBox
        Me.txtQTyFree = New System.Windows.Forms.TextBox
        Me.lblUnitFree = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtDONo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnStopProcess = New System.Windows.Forms.Button
        Me.lblProgress = New System.Windows.Forms.Label
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryDO = New System.Windows.Forms.TextBox
        Me.lblSecondaryUnits = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitsCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboItem.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCustomerAddIdentifier
        '
        Me.txtCustomerAddIdentifier.Location = New System.Drawing.Point(117, 69)
        Me.txtCustomerAddIdentifier.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtCustomerAddIdentifier.Name = "txtCustomerAddIdentifier"
        Me.txtCustomerAddIdentifier.Size = New System.Drawing.Size(32, 22)
        Me.txtCustomerAddIdentifier.TabIndex = 100
        Me.txtCustomerAddIdentifier.TabStop = False
        Me.txtCustomerAddIdentifier.Visible = False
        '
        'chkFakturPajak
        '
        Me.chkFakturPajak.AutoSize = True
        Me.chkFakturPajak.Location = New System.Drawing.Point(319, 119)
        Me.chkFakturPajak.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkFakturPajak.Name = "chkFakturPajak"
        Me.chkFakturPajak.Size = New System.Drawing.Size(170, 21)
        Me.chkFakturPajak.TabIndex = 13
        Me.chkFakturPajak.TabStop = False
        Me.chkFakturPajak.Text = "Faktur Pajak && Invoice"
        Me.chkFakturPajak.UseVisualStyleBackColor = True
        Me.chkFakturPajak.Visible = False
        '
        'chkKwitansi
        '
        Me.chkKwitansi.AutoSize = True
        Me.chkKwitansi.Location = New System.Drawing.Point(148, 119)
        Me.chkKwitansi.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkKwitansi.Name = "chkKwitansi"
        Me.chkKwitansi.Size = New System.Drawing.Size(81, 21)
        Me.chkKwitansi.TabIndex = 12
        Me.chkKwitansi.TabStop = False
        Me.chkKwitansi.Text = "Kwitansi"
        Me.chkKwitansi.UseVisualStyleBackColor = True
        Me.chkKwitansi.Visible = False
        '
        'llbAddress
        '
        Me.llbAddress.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAddress.AutoSize = True
        Me.llbAddress.Location = New System.Drawing.Point(507, 69)
        Me.llbAddress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbAddress.Name = "llbAddress"
        Me.llbAddress.Size = New System.Drawing.Size(77, 17)
        Me.llbAddress.TabIndex = 8
        Me.llbAddress.TabStop = True
        Me.llbAddress.Text = "View Detail"
        Me.ToolTip1.SetToolTip(Me.llbAddress, "click here to view customer's address detail")
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(17, 70)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(97, 17)
        Me.Label7.TabIndex = 157
        Me.Label7.Text = "Address Code"
        '
        'cboCustomerAddress
        '
        Me.cboCustomerAddress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCustomerAddress.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerAddress.FormattingEnabled = True
        Me.cboCustomerAddress.Location = New System.Drawing.Point(153, 66)
        Me.cboCustomerAddress.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboCustomerAddress.Name = "cboCustomerAddress"
        Me.cboCustomerAddress.Size = New System.Drawing.Size(344, 24)
        Me.cboCustomerAddress.TabIndex = 7
        '
        'DtpExpired
        '
        Me.DtpExpired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpExpired.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpired.Enabled = False
        Me.DtpExpired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpired.Location = New System.Drawing.Point(143, 70)
        Me.DtpExpired.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpExpired.Name = "DtpExpired"
        Me.DtpExpired.Size = New System.Drawing.Size(436, 22)
        Me.DtpExpired.TabIndex = 25
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(608, 87)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 17)
        Me.Label12.TabIndex = 282
        Me.Label12.Text = "Expiry Date"
        '
        'chkExpiryDateControlled
        '
        Me.chkExpiryDateControlled.AutoSize = True
        Me.chkExpiryDateControlled.Location = New System.Drawing.Point(772, 306)
        Me.chkExpiryDateControlled.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkExpiryDateControlled.Name = "chkExpiryDateControlled"
        Me.chkExpiryDateControlled.Size = New System.Drawing.Size(102, 21)
        Me.chkExpiryDateControlled.TabIndex = 279
        Me.chkExpiryDateControlled.TabStop = False
        Me.chkExpiryDateControlled.Text = "Expiry Date"
        Me.chkExpiryDateControlled.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlled.Visible = False
        '
        'chkBatchNoControlled
        '
        Me.chkBatchNoControlled.AutoSize = True
        Me.chkBatchNoControlled.Location = New System.Drawing.Point(772, 278)
        Me.chkBatchNoControlled.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkBatchNoControlled.Name = "chkBatchNoControlled"
        Me.chkBatchNoControlled.Size = New System.Drawing.Size(88, 21)
        Me.chkBatchNoControlled.TabIndex = 278
        Me.chkBatchNoControlled.TabStop = False
        Me.chkBatchNoControlled.Text = "Batch No"
        Me.chkBatchNoControlled.UseVisualStyleBackColor = True
        Me.chkBatchNoControlled.Visible = False
        '
        'chkUseSecondaryUnit
        '
        Me.chkUseSecondaryUnit.AutoSize = True
        Me.chkUseSecondaryUnit.Location = New System.Drawing.Point(772, 330)
        Me.chkUseSecondaryUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkUseSecondaryUnit.Name = "chkUseSecondaryUnit"
        Me.chkUseSecondaryUnit.Size = New System.Drawing.Size(156, 21)
        Me.chkUseSecondaryUnit.TabIndex = 277
        Me.chkUseSecondaryUnit.TabStop = False
        Me.chkUseSecondaryUnit.Text = "Use Secondary Unit"
        Me.chkUseSecondaryUnit.UseVisualStyleBackColor = True
        Me.chkUseSecondaryUnit.Visible = False
        '
        'chkUseCustCodeNo
        '
        Me.chkUseCustCodeNo.AutoSize = True
        Me.chkUseCustCodeNo.Checked = True
        Me.chkUseCustCodeNo.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseCustCodeNo.Location = New System.Drawing.Point(21, 304)
        Me.chkUseCustCodeNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkUseCustCodeNo.Name = "chkUseCustCodeNo"
        Me.chkUseCustCodeNo.Size = New System.Drawing.Size(159, 21)
        Me.chkUseCustCodeNo.TabIndex = 276
        Me.chkUseCustCodeNo.TabStop = False
        Me.chkUseCustCodeNo.Text = "Customer's Code No"
        Me.chkUseCustCodeNo.UseVisualStyleBackColor = True
        Me.chkUseCustCodeNo.Visible = False
        '
        'cboCustCodeNo
        '
        Me.cboCustCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustCodeNo.FormattingEnabled = True
        Me.cboCustCodeNo.Location = New System.Drawing.Point(192, 302)
        Me.cboCustCodeNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboCustCodeNo.Name = "cboCustCodeNo"
        Me.cboCustCodeNo.Size = New System.Drawing.Size(372, 24)
        Me.cboCustCodeNo.TabIndex = 8
        Me.cboCustCodeNo.TabStop = False
        Me.cboCustCodeNo.Visible = False
        '
        'txtCustProductDesc
        '
        Me.txtCustProductDesc.Location = New System.Drawing.Point(192, 337)
        Me.txtCustProductDesc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtCustProductDesc.MaxLength = 150
        Me.txtCustProductDesc.Name = "txtCustProductDesc"
        Me.txtCustProductDesc.ReadOnly = True
        Me.txtCustProductDesc.Size = New System.Drawing.Size(372, 22)
        Me.txtCustProductDesc.TabIndex = 7
        Me.txtCustProductDesc.TabStop = False
        Me.txtCustProductDesc.Visible = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(17, 341)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(153, 17)
        Me.Label19.TabIndex = 266
        Me.Label19.Text = "Customer's Description"
        Me.Label19.Visible = False
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(91, 68)
        Me.txtPrimaryCodeNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryCodeNo.MaxLength = 30
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(48, 22)
        Me.txtPrimaryCodeNo.TabIndex = 14
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Text = "-1"
        Me.txtPrimaryCodeNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCodeNo.Visible = False
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(143, 134)
        Me.TxtRemarkItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(350, 22)
        Me.TxtRemarkItem.TabIndex = 29
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(608, 156)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 17)
        Me.Label16.TabIndex = 226
        Me.Label16.Text = "Remarks"
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(502, 133)
        Me.BtnInput.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(69, 27)
        Me.BtnInput.TabIndex = 30
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'cboSalesOrderNo
        '
        Me.cboSalesOrderNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSalesOrderNo.FormattingEnabled = True
        Me.cboSalesOrderNo.Location = New System.Drawing.Point(148, 5)
        Me.cboSalesOrderNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboSalesOrderNo.Name = "cboSalesOrderNo"
        Me.cboSalesOrderNo.Size = New System.Drawing.Size(429, 24)
        Me.cboSalesOrderNo.TabIndex = 11
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(148, 98)
        Me.txtDesc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDesc.MaxLength = 150
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(429, 22)
        Me.txtDesc.TabIndex = 17
        Me.txtDesc.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(7, 107)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(79, 17)
        Me.Label11.TabIndex = 197
        Me.Label11.Text = "Description"
        '
        'lblCodeNo
        '
        Me.lblCodeNo.AutoSize = True
        Me.lblCodeNo.Location = New System.Drawing.Point(7, 74)
        Me.lblCodeNo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCodeNo.Name = "lblCodeNo"
        Me.lblCodeNo.Size = New System.Drawing.Size(63, 17)
        Me.lblCodeNo.TabIndex = 196
        Me.lblCodeNo.Text = "Code No"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(7, 74)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(50, 17)
        Me.Label9.TabIndex = 195
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'cboCodeNo
        '
        Me.cboCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeNo.FormattingEnabled = True
        Me.cboCodeNo.Location = New System.Drawing.Point(148, 66)
        Me.cboCodeNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboCodeNo.Name = "cboCodeNo"
        Me.cboCodeNo.Size = New System.Drawing.Size(328, 24)
        Me.cboCodeNo.TabIndex = 15
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCustomer.Location = New System.Drawing.Point(252, 37)
        Me.btnFindCustomer.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(43, 25)
        Me.btnFindCustomer.TabIndex = 5
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "...."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'llbAddBooking
        '
        Me.llbAddBooking.AutoSize = True
        Me.llbAddBooking.Location = New System.Drawing.Point(485, 14)
        Me.llbAddBooking.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbAddBooking.Name = "llbAddBooking"
        Me.llbAddBooking.Size = New System.Drawing.Size(56, 17)
        Me.llbAddBooking.TabIndex = 318
        Me.llbAddBooking.TabStop = True
        Me.llbAddBooking.Text = "Booked"
        Me.ToolTip1.SetToolTip(Me.llbAddBooking, "Click here to quickly add data from SO")
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.TxtReadOnly)
        Me.gboItem.Controls.Add(Me.chkLocationControlled)
        Me.gboItem.Controls.Add(Me.chkPdoControlled)
        Me.gboItem.Controls.Add(Me.cboSecondaryQty)
        Me.gboItem.Controls.Add(Me.lblSecQty)
        Me.gboItem.Controls.Add(Me.txtSecondaryQty)
        Me.gboItem.Controls.Add(Me.lblExpiryDate)
        Me.gboItem.Controls.Add(Me.Label5)
        Me.gboItem.Controls.Add(Me.Label12)
        Me.gboItem.Controls.Add(Me.chkExpiryDateControlled)
        Me.gboItem.Controls.Add(Me.chkBatchNoControlled)
        Me.gboItem.Controls.Add(Me.chkUseSecondaryUnit)
        Me.gboItem.Controls.Add(Me.chkUseCustCodeNo)
        Me.gboItem.Controls.Add(Me.cboCustCodeNo)
        Me.gboItem.Controls.Add(Me.txtCustProductDesc)
        Me.gboItem.Controls.Add(Me.Label19)
        Me.gboItem.Controls.Add(Me.Label16)
        Me.gboItem.Controls.Add(Me.SplitContainer2)
        Me.gboItem.Controls.Add(Me.txtQTyFree)
        Me.gboItem.Controls.Add(Me.lblUnitFree)
        Me.gboItem.Location = New System.Drawing.Point(8, 116)
        Me.gboItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboItem.Size = New System.Drawing.Size(1189, 230)
        Me.gboItem.TabIndex = 0
        Me.gboItem.TabStop = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReadOnly.Location = New System.Drawing.Point(961, 271)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtReadOnly.MaxLength = 50
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(131, 22)
        Me.TxtReadOnly.TabIndex = 158
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'chkLocationControlled
        '
        Me.chkLocationControlled.AutoSize = True
        Me.chkLocationControlled.Location = New System.Drawing.Point(772, 356)
        Me.chkLocationControlled.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkLocationControlled.Name = "chkLocationControlled"
        Me.chkLocationControlled.Size = New System.Drawing.Size(84, 21)
        Me.chkLocationControlled.TabIndex = 302
        Me.chkLocationControlled.TabStop = False
        Me.chkLocationControlled.Text = "Location"
        Me.chkLocationControlled.UseVisualStyleBackColor = True
        Me.chkLocationControlled.Visible = False
        '
        'chkPdoControlled
        '
        Me.chkPdoControlled.AutoSize = True
        Me.chkPdoControlled.Location = New System.Drawing.Point(772, 384)
        Me.chkPdoControlled.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkPdoControlled.Name = "chkPdoControlled"
        Me.chkPdoControlled.Size = New System.Drawing.Size(58, 21)
        Me.chkPdoControlled.TabIndex = 301
        Me.chkPdoControlled.TabStop = False
        Me.chkPdoControlled.Text = "PdO"
        Me.chkPdoControlled.UseVisualStyleBackColor = True
        Me.chkPdoControlled.Visible = False
        '
        'cboSecondaryQty
        '
        Me.cboSecondaryQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSecondaryQty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSecondaryQty.FormattingEnabled = True
        Me.cboSecondaryQty.Location = New System.Drawing.Point(961, 238)
        Me.cboSecondaryQty.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboSecondaryQty.Name = "cboSecondaryQty"
        Me.cboSecondaryQty.Size = New System.Drawing.Size(124, 24)
        Me.cboSecondaryQty.TabIndex = 12
        '
        'lblSecQty
        '
        Me.lblSecQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQty.AutoSize = True
        Me.lblSecQty.Location = New System.Drawing.Point(931, 242)
        Me.lblSecQty.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSecQty.Name = "lblSecQty"
        Me.lblSecQty.Size = New System.Drawing.Size(22, 17)
        Me.lblSecQty.TabIndex = 292
        Me.lblSecQty.Text = "@"
        '
        'txtSecondaryQty
        '
        Me.txtSecondaryQty.Location = New System.Drawing.Point(688, 117)
        Me.txtSecondaryQty.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtSecondaryQty.MaxLength = 30
        Me.txtSecondaryQty.Name = "txtSecondaryQty"
        Me.txtSecondaryQty.Size = New System.Drawing.Size(45, 22)
        Me.txtSecondaryQty.TabIndex = 26
        Me.txtSecondaryQty.TabStop = False
        Me.txtSecondaryQty.Text = "-1"
        Me.txtSecondaryQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSecondaryQty.Visible = False
        '
        'lblExpiryDate
        '
        Me.lblExpiryDate.AutoSize = True
        Me.lblExpiryDate.Location = New System.Drawing.Point(743, 89)
        Me.lblExpiryDate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblExpiryDate.Name = "lblExpiryDate"
        Me.lblExpiryDate.Size = New System.Drawing.Size(28, 17)
        Me.lblExpiryDate.TabIndex = 284
        Me.lblExpiryDate.Text = "n/a"
        Me.lblExpiryDate.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(608, 55)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 17)
        Me.Label5.TabIndex = 283
        Me.Label5.Text = "Batch No"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Location = New System.Drawing.Point(5, 14)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtToleranceUp)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ChkIgnoreTolerance)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboCustPONO)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboSalesOrderNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtQtyOrderBalance)
        Me.SplitContainer2.Panel1.Controls.Add(Me.llbDetailStock)
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblUnitOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.llbAddBooking)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPrimaryCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPrimarySO)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblCustPONo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label18)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPrimaryCategory)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDesc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblUnitsPendingDO)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblSalesOrder)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPendingDO)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.LlbQty)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRatioDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRatioReference)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyperBatchDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblAvailableUnits)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblAvailableUnitsMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtAvailableMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblBatchNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboBatchNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblUnitsMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblUnitsDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label20)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyperBatchMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtAvailable)
        Me.SplitContainer2.Panel2.Controls.Add(Me.DtpExpired)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtPrimaryUnitsDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtPrimaryUnitsMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemarkItem)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer2.Size = New System.Drawing.Size(1179, 209)
        Me.SplitContainer2.SplitterDistance = 589
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 164
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(385, 133)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 17)
        Me.Label2.TabIndex = 322
        Me.Label2.Text = "+"
        '
        'TxtToleranceUp
        '
        Me.TxtToleranceUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtToleranceUp.Location = New System.Drawing.Point(411, 129)
        Me.TxtToleranceUp.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtToleranceUp.MaxLength = 150
        Me.TxtToleranceUp.Name = "TxtToleranceUp"
        Me.TxtToleranceUp.ReadOnly = True
        Me.TxtToleranceUp.Size = New System.Drawing.Size(95, 22)
        Me.TxtToleranceUp.TabIndex = 19
        Me.TxtToleranceUp.TabStop = False
        Me.TxtToleranceUp.Text = "0"
        Me.TxtToleranceUp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkIgnoreTolerance
        '
        Me.ChkIgnoreTolerance.AutoSize = True
        Me.ChkIgnoreTolerance.Location = New System.Drawing.Point(148, 158)
        Me.ChkIgnoreTolerance.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkIgnoreTolerance.Name = "ChkIgnoreTolerance"
        Me.ChkIgnoreTolerance.Size = New System.Drawing.Size(180, 21)
        Me.ChkIgnoreTolerance.TabIndex = 20
        Me.ChkIgnoreTolerance.Text = "Ignore Lower Tolerance"
        Me.ChkIgnoreTolerance.UseVisualStyleBackColor = True
        '
        'cboCustPONO
        '
        Me.cboCustPONO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustPONO.FormattingEnabled = True
        Me.cboCustPONO.Location = New System.Drawing.Point(148, 36)
        Me.cboCustPONO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboCustPONO.Name = "cboCustPONO"
        Me.cboCustPONO.Size = New System.Drawing.Size(429, 24)
        Me.cboCustPONO.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 138)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 17)
        Me.Label8.TabIndex = 304
        Me.Label8.Text = "Order Balance"
        '
        'txtQtyOrderBalance
        '
        Me.txtQtyOrderBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQtyOrderBalance.Location = New System.Drawing.Point(148, 129)
        Me.txtQtyOrderBalance.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtQtyOrderBalance.MaxLength = 150
        Me.txtQtyOrderBalance.Name = "txtQtyOrderBalance"
        Me.txtQtyOrderBalance.ReadOnly = True
        Me.txtQtyOrderBalance.Size = New System.Drawing.Size(228, 22)
        Me.txtQtyOrderBalance.TabIndex = 18
        Me.txtQtyOrderBalance.TabStop = False
        Me.txtQtyOrderBalance.Text = "0"
        Me.txtQtyOrderBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'llbDetailStock
        '
        Me.llbDetailStock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbDetailStock.AutoSize = True
        Me.llbDetailStock.Location = New System.Drawing.Point(487, 70)
        Me.llbDetailStock.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbDetailStock.Name = "llbDetailStock"
        Me.llbDetailStock.Size = New System.Drawing.Size(81, 17)
        Me.llbDetailStock.TabIndex = 16
        Me.llbDetailStock.TabStop = True
        Me.llbDetailStock.Text = "Show Stock"
        '
        'LblUnitOB
        '
        Me.LblUnitOB.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitOB.AutoSize = True
        Me.LblUnitOB.Location = New System.Drawing.Point(519, 133)
        Me.LblUnitOB.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblUnitOB.Name = "LblUnitOB"
        Me.LblUnitOB.Size = New System.Drawing.Size(40, 17)
        Me.LblUnitOB.TabIndex = 305
        Me.LblUnitOB.Text = "Units"
        '
        'txtPrimarySO
        '
        Me.txtPrimarySO.Location = New System.Drawing.Point(116, 5)
        Me.txtPrimarySO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimarySO.MaxLength = 30
        Me.txtPrimarySO.Name = "txtPrimarySO"
        Me.txtPrimarySO.Size = New System.Drawing.Size(48, 22)
        Me.txtPrimarySO.TabIndex = 10
        Me.txtPrimarySO.TabStop = False
        Me.txtPrimarySO.Text = "-1"
        Me.txtPrimarySO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimarySO.Visible = False
        '
        'lblCustPONo
        '
        Me.lblCustPONo.AutoSize = True
        Me.lblCustPONo.Location = New System.Drawing.Point(7, 42)
        Me.lblCustPONo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCustPONo.Name = "lblCustPONo"
        Me.lblCustPONo.Size = New System.Drawing.Size(124, 17)
        Me.lblCustPONo.TabIndex = 310
        Me.lblCustPONo.Text = "Customer's PO No"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(7, 10)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(50, 17)
        Me.Label18.TabIndex = 213
        Me.Label18.Text = "PO No"
        Me.Label18.Visible = False
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(60, 68)
        Me.txtPrimaryCategory.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryCategory.MaxLength = 30
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(45, 22)
        Me.txtPrimaryCategory.TabIndex = 13
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Text = "-1"
        Me.txtPrimaryCategory.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCategory.Visible = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(7, 183)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(85, 17)
        Me.Label14.TabIndex = 312
        Me.Label14.Text = "Pending DO"
        '
        'lblUnitsPendingDO
        '
        Me.lblUnitsPendingDO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsPendingDO.AutoSize = True
        Me.lblUnitsPendingDO.Location = New System.Drawing.Point(519, 183)
        Me.lblUnitsPendingDO.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUnitsPendingDO.Name = "lblUnitsPendingDO"
        Me.lblUnitsPendingDO.Size = New System.Drawing.Size(40, 17)
        Me.lblUnitsPendingDO.TabIndex = 313
        Me.lblUnitsPendingDO.Text = "Units"
        '
        'lblSalesOrder
        '
        Me.lblSalesOrder.AutoSize = True
        Me.lblSalesOrder.Location = New System.Drawing.Point(7, 10)
        Me.lblSalesOrder.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSalesOrder.Name = "lblSalesOrder"
        Me.lblSalesOrder.Size = New System.Drawing.Size(106, 17)
        Me.lblSalesOrder.TabIndex = 214
        Me.lblSalesOrder.Text = "Sales Order No"
        '
        'txtPendingDO
        '
        Me.txtPendingDO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPendingDO.Location = New System.Drawing.Point(148, 180)
        Me.txtPendingDO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPendingDO.MaxLength = 150
        Me.txtPendingDO.Name = "txtPendingDO"
        Me.txtPendingDO.ReadOnly = True
        Me.txtPendingDO.Size = New System.Drawing.Size(357, 22)
        Me.txtPendingDO.TabIndex = 21
        Me.txtPendingDO.TabStop = False
        Me.txtPendingDO.Text = "0"
        Me.txtPendingDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LlbQty
        '
        Me.LlbQty.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbQty.AutoSize = True
        Me.LlbQty.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.LlbQty.Location = New System.Drawing.Point(11, 107)
        Me.LlbQty.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LlbQty.Name = "LlbQty"
        Me.LlbQty.Size = New System.Drawing.Size(30, 17)
        Me.LlbQty.TabIndex = 323
        Me.LlbQty.TabStop = True
        Me.LlbQty.Text = "Qty"
        Me.LlbQty.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtRatioDefault
        '
        Me.TxtRatioDefault.Location = New System.Drawing.Point(184, 166)
        Me.TxtRatioDefault.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRatioDefault.Name = "TxtRatioDefault"
        Me.TxtRatioDefault.Size = New System.Drawing.Size(33, 22)
        Me.TxtRatioDefault.TabIndex = 32
        Me.TxtRatioDefault.TabStop = False
        Me.TxtRatioDefault.Visible = False
        '
        'TxtRatioReference
        '
        Me.TxtRatioReference.Location = New System.Drawing.Point(141, 166)
        Me.TxtRatioReference.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRatioReference.Name = "TxtRatioReference"
        Me.TxtRatioReference.Size = New System.Drawing.Size(33, 22)
        Me.TxtRatioReference.TabIndex = 31
        Me.TxtRatioReference.TabStop = False
        Me.TxtRatioReference.Visible = False
        '
        'TxtQtyperBatchDefault
        '
        Me.TxtQtyperBatchDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchDefault.Location = New System.Drawing.Point(404, 103)
        Me.TxtQtyperBatchDefault.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtQtyperBatchDefault.MaxLength = 18
        Me.TxtQtyperBatchDefault.Name = "TxtQtyperBatchDefault"
        Me.TxtQtyperBatchDefault.Size = New System.Drawing.Size(89, 22)
        Me.TxtQtyperBatchDefault.TabIndex = 28
        Me.TxtQtyperBatchDefault.Text = "0"
        Me.TxtQtyperBatchDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblAvailableUnits
        '
        Me.lblAvailableUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAvailableUnits.AutoSize = True
        Me.lblAvailableUnits.Location = New System.Drawing.Point(502, 10)
        Me.lblAvailableUnits.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAvailableUnits.Name = "lblAvailableUnits"
        Me.lblAvailableUnits.Size = New System.Drawing.Size(40, 17)
        Me.lblAvailableUnits.TabIndex = 316
        Me.lblAvailableUnits.Text = "Units"
        '
        'lblAvailableUnitsMulti
        '
        Me.lblAvailableUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAvailableUnitsMulti.AutoSize = True
        Me.lblAvailableUnitsMulti.Location = New System.Drawing.Point(342, 9)
        Me.lblAvailableUnitsMulti.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAvailableUnitsMulti.Name = "lblAvailableUnitsMulti"
        Me.lblAvailableUnitsMulti.Size = New System.Drawing.Size(40, 17)
        Me.lblAvailableUnitsMulti.TabIndex = 318
        Me.lblAvailableUnitsMulti.Text = "Units"
        '
        'TxtAvailableMulti
        '
        Me.TxtAvailableMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAvailableMulti.Location = New System.Drawing.Point(143, 5)
        Me.TxtAvailableMulti.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAvailableMulti.MaxLength = 150
        Me.TxtAvailableMulti.Name = "TxtAvailableMulti"
        Me.TxtAvailableMulti.ReadOnly = True
        Me.TxtAvailableMulti.Size = New System.Drawing.Size(190, 22)
        Me.TxtAvailableMulti.TabIndex = 22
        Me.TxtAvailableMulti.TabStop = False
        Me.TxtAvailableMulti.Text = "0"
        Me.TxtAvailableMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(11, 9)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(65, 17)
        Me.Label15.TabIndex = 315
        Me.Label15.Text = "Available"
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Location = New System.Drawing.Point(144, 42)
        Me.lblBatchNo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(28, 17)
        Me.lblBatchNo.TabIndex = 285
        Me.lblBatchNo.Text = "n/a"
        Me.lblBatchNo.Visible = False
        '
        'cboBatchNo
        '
        Me.cboBatchNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(143, 37)
        Me.cboBatchNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(436, 24)
        Me.cboBatchNo.TabIndex = 24
        '
        'lblUnitsMulti
        '
        Me.lblUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsMulti.AutoSize = True
        Me.lblUnitsMulti.Location = New System.Drawing.Point(342, 107)
        Me.lblUnitsMulti.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUnitsMulti.Name = "lblUnitsMulti"
        Me.lblUnitsMulti.Size = New System.Drawing.Size(40, 17)
        Me.lblUnitsMulti.TabIndex = 290
        Me.lblUnitsMulti.Text = "Units"
        '
        'lblUnitsDefault
        '
        Me.lblUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsDefault.AutoSize = True
        Me.lblUnitsDefault.Location = New System.Drawing.Point(502, 106)
        Me.lblUnitsDefault.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUnitsDefault.Name = "lblUnitsDefault"
        Me.lblUnitsDefault.Size = New System.Drawing.Size(40, 17)
        Me.lblUnitsDefault.TabIndex = 291
        Me.lblUnitsDefault.Text = "Units"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(8, 10)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(63, 17)
        Me.Label20.TabIndex = 307
        Me.Label20.Text = "Free Qty"
        Me.Label20.Visible = False
        '
        'TxtQtyperBatchMulti
        '
        Me.TxtQtyperBatchMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchMulti.Location = New System.Drawing.Point(143, 103)
        Me.TxtQtyperBatchMulti.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtQtyperBatchMulti.MaxLength = 18
        Me.TxtQtyperBatchMulti.Name = "TxtQtyperBatchMulti"
        Me.TxtQtyperBatchMulti.Size = New System.Drawing.Size(190, 22)
        Me.TxtQtyperBatchMulti.TabIndex = 27
        Me.TxtQtyperBatchMulti.Text = "0"
        Me.TxtQtyperBatchMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvailable
        '
        Me.txtAvailable.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAvailable.Location = New System.Drawing.Point(404, 5)
        Me.txtAvailable.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtAvailable.MaxLength = 150
        Me.txtAvailable.Name = "txtAvailable"
        Me.txtAvailable.ReadOnly = True
        Me.txtAvailable.Size = New System.Drawing.Size(89, 22)
        Me.txtAvailable.TabIndex = 23
        Me.txtAvailable.TabStop = False
        Me.txtAvailable.Text = "0"
        Me.txtAvailable.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPrimaryUnitsDefault
        '
        Me.txtPrimaryUnitsDefault.Location = New System.Drawing.Point(308, 166)
        Me.txtPrimaryUnitsDefault.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryUnitsDefault.MaxLength = 30
        Me.txtPrimaryUnitsDefault.Name = "txtPrimaryUnitsDefault"
        Me.txtPrimaryUnitsDefault.Size = New System.Drawing.Size(45, 22)
        Me.txtPrimaryUnitsDefault.TabIndex = 34
        Me.txtPrimaryUnitsDefault.TabStop = False
        Me.txtPrimaryUnitsDefault.Text = "-1"
        Me.txtPrimaryUnitsDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnitsDefault.Visible = False
        '
        'txtPrimaryUnitsMulti
        '
        Me.txtPrimaryUnitsMulti.Location = New System.Drawing.Point(252, 166)
        Me.txtPrimaryUnitsMulti.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryUnitsMulti.MaxLength = 30
        Me.txtPrimaryUnitsMulti.Name = "txtPrimaryUnitsMulti"
        Me.txtPrimaryUnitsMulti.Size = New System.Drawing.Size(45, 22)
        Me.txtPrimaryUnitsMulti.TabIndex = 33
        Me.txtPrimaryUnitsMulti.TabStop = False
        Me.txtPrimaryUnitsMulti.Text = "-1"
        Me.txtPrimaryUnitsMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnitsMulti.Visible = False
        '
        'txtQTyFree
        '
        Me.txtQTyFree.Location = New System.Drawing.Point(743, 23)
        Me.txtQTyFree.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtQTyFree.MaxLength = 150
        Me.txtQTyFree.Name = "txtQTyFree"
        Me.txtQTyFree.ReadOnly = True
        Me.txtQTyFree.Size = New System.Drawing.Size(105, 22)
        Me.txtQTyFree.TabIndex = 306
        Me.txtQTyFree.TabStop = False
        Me.txtQTyFree.Text = "0"
        Me.txtQTyFree.Visible = False
        '
        'lblUnitFree
        '
        Me.lblUnitFree.AutoSize = True
        Me.lblUnitFree.Location = New System.Drawing.Point(857, 27)
        Me.lblUnitFree.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUnitFree.Name = "lblUnitFree"
        Me.lblUnitFree.Size = New System.Drawing.Size(40, 17)
        Me.lblUnitFree.TabIndex = 308
        Me.lblUnitFree.Text = "Units"
        Me.lblUnitFree.Visible = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(655, 399)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(30, 17)
        Me.Label13.TabIndex = 289
        Me.Label13.Text = "Qty"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(148, 5)
        Me.TxtRemark.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRemark.MaxLength = 250
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(430, 86)
        Me.TxtRemark.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 9)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 17)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(298, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(298, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(298, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(298, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 41)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Customer"
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(117, 41)
        Me.txtPrimaryCustomer.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(32, 22)
        Me.txtPrimaryCustomer.TabIndex = 3
        Me.txtPrimaryCustomer.TabStop = False
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(153, 37)
        Me.txtCustomerCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtCustomerCode.MaxLength = 50
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(90, 22)
        Me.txtCustomerCode.TabIndex = 4
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerName.Location = New System.Drawing.Point(303, 37)
        Me.TxtCustomerName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(279, 22)
        Me.TxtCustomerName.TabIndex = 6
        Me.TxtCustomerName.TabStop = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(8, 693)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 4
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(53, 693)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 5
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(420, 6)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(161, 22)
        Me.dtpDate.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(611, 693)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(139, 32)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(456, 693)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(139, 32)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtDONo
        '
        Me.TxtDONo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDONo.Location = New System.Drawing.Point(153, 6)
        Me.TxtDONo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDONo.MaxLength = 50
        Me.TxtDONo.Name = "TxtDONo"
        Me.TxtDONo.ReadOnly = True
        Me.TxtDONo.Size = New System.Drawing.Size(210, 22)
        Me.TxtDONo.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(372, 10)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 17)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 11)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 17)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "DO No"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 736)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1215, 25)
        Me.StatusStrip1.TabIndex = 55
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(611, 740)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 51
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(496, 740)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 50
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.btnStopProcess)
        Me.gboDetail.Controls.Add(Me.lblProgress)
        Me.gboDetail.Controls.Add(Me.prgAll)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.chkFakturPajak)
        Me.gboDetail.Controls.Add(Me.Label13)
        Me.gboDetail.Controls.Add(Me.chkKwitansi)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.crv1)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.lblSecondaryUnits)
        Me.gboDetail.Location = New System.Drawing.Point(5, 0)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(1205, 732)
        Me.gboDetail.TabIndex = 54
        Me.gboDetail.TabStop = False
        '
        'btnStopProcess
        '
        Me.btnStopProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStopProcess.Location = New System.Drawing.Point(1092, 686)
        Me.btnStopProcess.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStopProcess.Name = "btnStopProcess"
        Me.btnStopProcess.Size = New System.Drawing.Size(100, 32)
        Me.btnStopProcess.TabIndex = 161
        Me.btnStopProcess.Text = "Stop Process"
        Me.btnStopProcess.UseVisualStyleBackColor = True
        Me.btnStopProcess.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(823, 711)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(49, 17)
        Me.lblProgress.TabIndex = 160
        Me.lblProgress.Text = "Ready"
        Me.lblProgress.Visible = False
        '
        'prgAll
        '
        Me.prgAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgAll.Location = New System.Drawing.Point(827, 686)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(260, 23)
        Me.prgAll.TabIndex = 159
        Me.prgAll.Visible = False
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.Location = New System.Drawing.Point(812, 421)
        Me.crv1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.Size = New System.Drawing.Size(150, 150)
        Me.crv1.TabIndex = 162
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 15)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDONo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryDO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerAddIdentifier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboCustomerAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Size = New System.Drawing.Size(1189, 98)
        Me.SplitContainer1.SplitterDistance = 594
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 163
        '
        'TxtPrimaryDO
        '
        Me.TxtPrimaryDO.Location = New System.Drawing.Point(117, 9)
        Me.TxtPrimaryDO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryDO.Name = "TxtPrimaryDO"
        Me.TxtPrimaryDO.Size = New System.Drawing.Size(32, 22)
        Me.TxtPrimaryDO.TabIndex = 0
        Me.TxtPrimaryDO.TabStop = False
        Me.TxtPrimaryDO.Text = "-1"
        Me.TxtPrimaryDO.Visible = False
        '
        'lblSecondaryUnits
        '
        Me.lblSecondaryUnits.AutoSize = True
        Me.lblSecondaryUnits.Location = New System.Drawing.Point(532, 399)
        Me.lblSecondaryUnits.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSecondaryUnits.Name = "lblSecondaryUnits"
        Me.lblSecondaryUnits.Size = New System.Drawing.Size(80, 17)
        Me.lblSecondaryUnits.TabIndex = 309
        Me.lblSecondaryUnits.Text = "Expiry Date"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(725, 740)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(107, 32)
        Me.btnPrint.TabIndex = 52
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(381, 740)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 49
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(267, 740)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 48
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(840, 740)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 53
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'Timer1
        '
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColCustCodeNo, Me.ColCustDesc, Me.ColSONO, Me.ColBatchNo, Me.ColExpiryDate, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColMasterUnits, Me.ColQty, Me.ColTotalQtyMulti, Me.ColUnitsCodeMulti, Me.ColRatioDesc, Me.ColTotalQty, Me.ColUnit, Me.ColSecQty, Me.ColSecUnit, Me.ColLocation, Me.ColWarehouse, Me.ColDeliveryDate, Me.ColRemarks, Me.ColPrimaryCodeNo, Me.ColPrimarySO, Me.ColPrimaryMasterUnits, Me.ColPrimaryUnit, Me.ColPrimarySecUnit, Me.ColRatioReference, Me.ColRatioDefault, Me.ColPrimaryUnitsMulti, Me.ColPrimaryLocation})
        Me.dgvItem.Location = New System.Drawing.Point(8, 358)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(1189, 322)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 310
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 88
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 104
        '
        'ColCustCodeNo
        '
        Me.ColCustCodeNo.HeaderText = "Customer's Code No"
        Me.ColCustCodeNo.Name = "ColCustCodeNo"
        Me.ColCustCodeNo.Visible = False
        Me.ColCustCodeNo.Width = 162
        '
        'ColCustDesc
        '
        Me.ColCustDesc.HeaderText = "Customer's Item Description"
        Me.ColCustDesc.Name = "ColCustDesc"
        Me.ColCustDesc.Visible = False
        Me.ColCustDesc.Width = 208
        '
        'ColSONO
        '
        Me.ColSONO.HeaderText = "SO. No"
        Me.ColSONO.Name = "ColSONO"
        Me.ColSONO.Width = 79
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.Width = 91
        '
        'ColExpiryDate
        '
        Me.ColExpiryDate.HeaderText = "Expiry Date"
        Me.ColExpiryDate.Name = "ColExpiryDate"
        Me.ColExpiryDate.Width = 105
        '
        'ColQtyMaster
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMaster.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 102
        '
        'ColQtyPerMaster
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyPerMaster.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 128
        '
        'ColMasterUnits
        '
        Me.ColMasterUnits.HeaderText = "Master Unit"
        Me.ColMasterUnits.Name = "ColMasterUnits"
        Me.ColMasterUnits.Visible = False
        Me.ColMasterUnits.Width = 105
        '
        'ColQty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Visible = False
        Me.ColQty.Width = 55
        '
        'ColTotalQtyMulti
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQtyMulti.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColTotalQtyMulti.HeaderText = "Total Qty"
        Me.ColTotalQtyMulti.Name = "ColTotalQtyMulti"
        Me.ColTotalQtyMulti.Width = 91
        '
        'ColUnitsCodeMulti
        '
        Me.ColUnitsCodeMulti.HeaderText = "Unit"
        Me.ColUnitsCodeMulti.Name = "ColUnitsCodeMulti"
        Me.ColUnitsCodeMulti.Width = 58
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.Width = 47
        '
        'ColTotalQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColTotalQty.HeaderText = "Qty Default"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.Width = 104
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 107
        '
        'ColSecQty
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecQty.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColSecQty.HeaderText = "@"
        Me.ColSecQty.Name = "ColSecQty"
        Me.ColSecQty.Width = 47
        '
        'ColSecUnit
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecUnit.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColSecUnit.HeaderText = "@ Unit"
        Me.ColSecUnit.Name = "ColSecUnit"
        Me.ColSecUnit.Width = 76
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 87
        '
        'ColWarehouse
        '
        Me.ColWarehouse.HeaderText = "Warehouse"
        Me.ColWarehouse.Name = "ColWarehouse"
        Me.ColWarehouse.Width = 106
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.Width = 118
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 89
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 132
        '
        'ColPrimarySO
        '
        Me.ColPrimarySO.HeaderText = "PrimarySO"
        Me.ColPrimarySO.Name = "ColPrimarySO"
        Me.ColPrimarySO.Visible = False
        Me.ColPrimarySO.Width = 101
        '
        'ColPrimaryMasterUnits
        '
        Me.ColPrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColPrimaryMasterUnits.Name = "ColPrimaryMasterUnits"
        Me.ColPrimaryMasterUnits.Visible = False
        Me.ColPrimaryMasterUnits.Width = 157
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Units"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 117
        '
        'ColPrimarySecUnit
        '
        Me.ColPrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColPrimarySecUnit.Name = "ColPrimarySecUnit"
        Me.ColPrimarySecUnit.Visible = False
        Me.ColPrimarySecUnit.Width = 138
        '
        'ColRatioReference
        '
        Me.ColRatioReference.HeaderText = "ColRatioReference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 152
        '
        'ColRatioDefault
        '
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 115
        '
        'ColPrimaryUnitsMulti
        '
        Me.ColPrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColPrimaryUnitsMulti.Name = "ColPrimaryUnitsMulti"
        Me.ColPrimaryUnitsMulti.Visible = False
        Me.ColPrimaryUnitsMulti.Width = 150
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 139
        '
        'FrmTransaksiDO3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1215, 761)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1221, 724)
        Me.Name = "FrmTransaksiDO3"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Delivery Order"
        Me.gboItem.ResumeLayout(False)
        Me.gboItem.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtCustomerAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents chkFakturPajak As System.Windows.Forms.CheckBox
    Friend WithEvents chkKwitansi As System.Windows.Forms.CheckBox
    Friend WithEvents llbAddress As System.Windows.Forms.LinkLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboCustomerAddress As System.Windows.Forms.ComboBox
    Friend WithEvents DtpExpired As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents chkExpiryDateControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatchNoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseSecondaryUnit As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseCustCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents cboCustCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtCustProductDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents cboSalesOrderNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblCodeNo As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents lblAvailableUnits As System.Windows.Forms.Label
    Friend WithEvents txtAvailable As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblUnitsPendingDO As System.Windows.Forms.Label
    Friend WithEvents txtPendingDO As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblCustPONo As System.Windows.Forms.Label
    Friend WithEvents lblUnitFree As System.Windows.Forms.Label
    Friend WithEvents txtQTyFree As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents LblUnitOB As System.Windows.Forms.Label
    Friend WithEvents txtQtyOrderBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents llbDetailStock As System.Windows.Forms.LinkLabel
    Friend WithEvents chkLocationControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkPdoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents cboSecondaryQty As System.Windows.Forms.ComboBox
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimarySO As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitsDefault As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitsMulti As System.Windows.Forms.TextBox
    Friend WithEvents lblSecQty As System.Windows.Forms.Label
    Friend WithEvents txtSecondaryQty As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsMulti As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchMulti As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lblUnitsDefault As System.Windows.Forms.Label
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDate As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtDONo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryDO As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents llbAddBooking As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents cboCustPONO As System.Windows.Forms.ComboBox
    Friend WithEvents lblSalesOrder As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents btnStopProcess As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtToleranceUp As System.Windows.Forms.TextBox
    Friend WithEvents ChkIgnoreTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents TxtQtyperBatchDefault As System.Windows.Forms.TextBox
    Friend WithEvents lblAvailableUnitsMulti As System.Windows.Forms.Label
    Friend WithEvents TxtAvailableMulti As System.Windows.Forms.TextBox
    Friend WithEvents TxtRatioDefault As System.Windows.Forms.TextBox
    Friend WithEvents TxtRatioReference As System.Windows.Forms.TextBox
    Friend WithEvents LlbQty As System.Windows.Forms.LinkLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents lblSecondaryUnits As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitsCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
