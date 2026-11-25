<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiBookingInv
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.txtPrimaryBooking = New System.Windows.Forms.TextBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.lblUnitBooked = New System.Windows.Forms.Label
        Me.TxtBooked = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.lblWarehouse = New System.Windows.Forms.Label
        Me.lblLocation = New System.Windows.Forms.Label
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.txtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtWarehouse = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.dtpValid1 = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.lblAvailableUnits = New System.Windows.Forms.Label
        Me.txtAvailable = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.lblUnitsPendingDO = New System.Windows.Forms.Label
        Me.txtPendingDO = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.lblUnitFree = New System.Windows.Forms.Label
        Me.txtQtyOnHand = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.LblUnitOB = New System.Windows.Forms.Label
        Me.txtQtyOrderBalance = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.llbDetailStock = New System.Windows.Forms.LinkLabel
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkPdoControlled = New System.Windows.Forms.CheckBox
        Me.cboSecondaryQty = New System.Windows.Forms.ComboBox
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.txtPrimarySecondaryUnits = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnits = New System.Windows.Forms.TextBox
        Me.lblSecQty = New System.Windows.Forms.Label
        Me.txtSecondaryQty = New System.Windows.Forms.TextBox
        Me.lblUnits = New System.Windows.Forms.Label
        Me.TxtQtyperBatch = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblSecondaryUnits = New System.Windows.Forms.Label
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
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
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.lblCodeNo = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cboCodeNo = New System.Windows.Forms.ComboBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColValidUntil = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.BtnFindSONumber = New System.Windows.Forms.Button
        Me.TxtPrimarySONumber = New System.Windows.Forms.TextBox
        Me.TxtSONumber = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label37 = New System.Windows.Forms.Label
        Me.cboValidUntil = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DtpCreatedDate = New System.Windows.Forms.DateTimePicker
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtBookingNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslStatus = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.gboDetail.SuspendLayout()
        Me.gboItem.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.txtPrimaryBooking)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.TxtCustomerCode)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryCustomer)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.BtnFindSONumber)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySONumber)
        Me.gboDetail.Controls.Add(Me.TxtSONumber)
        Me.gboDetail.Controls.Add(Me.Label13)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label37)
        Me.gboDetail.Controls.Add(Me.cboValidUntil)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.DtpCreatedDate)
        Me.gboDetail.Controls.Add(Me.Label21)
        Me.gboDetail.Controls.Add(Me.TxtBookingNo)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Location = New System.Drawing.Point(8, 12)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(967, 486)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'txtPrimaryBooking
        '
        Me.txtPrimaryBooking.Location = New System.Drawing.Point(91, 15)
        Me.txtPrimaryBooking.MaxLength = 50
        Me.txtPrimaryBooking.Name = "txtPrimaryBooking"
        Me.txtPrimaryBooking.Size = New System.Drawing.Size(24, 20)
        Me.txtPrimaryBooking.TabIndex = 232
        Me.txtPrimaryBooking.TabStop = False
        Me.txtPrimaryBooking.Visible = False
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.lblUnitBooked)
        Me.gboItem.Controls.Add(Me.TxtBooked)
        Me.gboItem.Controls.Add(Me.Label18)
        Me.gboItem.Controls.Add(Me.lblWarehouse)
        Me.gboItem.Controls.Add(Me.lblLocation)
        Me.gboItem.Controls.Add(Me.cboLocation)
        Me.gboItem.Controls.Add(Me.txtPrimaryLocation)
        Me.gboItem.Controls.Add(Me.Label7)
        Me.gboItem.Controls.Add(Me.TxtWarehouse)
        Me.gboItem.Controls.Add(Me.Label10)
        Me.gboItem.Controls.Add(Me.dtpValid1)
        Me.gboItem.Controls.Add(Me.Label6)
        Me.gboItem.Controls.Add(Me.lblAvailableUnits)
        Me.gboItem.Controls.Add(Me.txtAvailable)
        Me.gboItem.Controls.Add(Me.Label15)
        Me.gboItem.Controls.Add(Me.lblUnitsPendingDO)
        Me.gboItem.Controls.Add(Me.txtPendingDO)
        Me.gboItem.Controls.Add(Me.Label14)
        Me.gboItem.Controls.Add(Me.lblUnitFree)
        Me.gboItem.Controls.Add(Me.txtQtyOnHand)
        Me.gboItem.Controls.Add(Me.Label20)
        Me.gboItem.Controls.Add(Me.LblUnitOB)
        Me.gboItem.Controls.Add(Me.txtQtyOrderBalance)
        Me.gboItem.Controls.Add(Me.Label8)
        Me.gboItem.Controls.Add(Me.llbDetailStock)
        Me.gboItem.Controls.Add(Me.chkLocationControlled)
        Me.gboItem.Controls.Add(Me.chkPdoControlled)
        Me.gboItem.Controls.Add(Me.cboSecondaryQty)
        Me.gboItem.Controls.Add(Me.cboBatchNo)
        Me.gboItem.Controls.Add(Me.txtPrimaryCategory)
        Me.gboItem.Controls.Add(Me.txtPrimarySecondaryUnits)
        Me.gboItem.Controls.Add(Me.txtPrimaryUnits)
        Me.gboItem.Controls.Add(Me.lblSecQty)
        Me.gboItem.Controls.Add(Me.txtSecondaryQty)
        Me.gboItem.Controls.Add(Me.lblUnits)
        Me.gboItem.Controls.Add(Me.TxtQtyperBatch)
        Me.gboItem.Controls.Add(Me.Label4)
        Me.gboItem.Controls.Add(Me.lblSecondaryUnits)
        Me.gboItem.Controls.Add(Me.lblBatchNo)
        Me.gboItem.Controls.Add(Me.lblExpiryDate)
        Me.gboItem.Controls.Add(Me.Label5)
        Me.gboItem.Controls.Add(Me.DtpExpired)
        Me.gboItem.Controls.Add(Me.Label12)
        Me.gboItem.Controls.Add(Me.chkExpiryDateControlled)
        Me.gboItem.Controls.Add(Me.chkBatchNoControlled)
        Me.gboItem.Controls.Add(Me.chkUseSecondaryUnit)
        Me.gboItem.Controls.Add(Me.chkUseCustCodeNo)
        Me.gboItem.Controls.Add(Me.cboCustCodeNo)
        Me.gboItem.Controls.Add(Me.txtCustProductDesc)
        Me.gboItem.Controls.Add(Me.Label19)
        Me.gboItem.Controls.Add(Me.txtPrimaryCodeNo)
        Me.gboItem.Controls.Add(Me.TxtRemarkItem)
        Me.gboItem.Controls.Add(Me.Label16)
        Me.gboItem.Controls.Add(Me.BtnInput)
        Me.gboItem.Controls.Add(Me.txtDesc)
        Me.gboItem.Controls.Add(Me.Label11)
        Me.gboItem.Controls.Add(Me.lblCodeNo)
        Me.gboItem.Controls.Add(Me.Label9)
        Me.gboItem.Controls.Add(Me.cboCodeNo)
        Me.gboItem.Location = New System.Drawing.Point(6, 123)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(952, 182)
        Me.gboItem.TabIndex = 6
        Me.gboItem.TabStop = False
        '
        'lblUnitBooked
        '
        Me.lblUnitBooked.AutoSize = True
        Me.lblUnitBooked.Location = New System.Drawing.Point(840, 103)
        Me.lblUnitBooked.Name = "lblUnitBooked"
        Me.lblUnitBooked.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitBooked.TabIndex = 328
        Me.lblUnitBooked.Text = "Units"
        Me.lblUnitBooked.Visible = False
        '
        'TxtBooked
        '
        Me.TxtBooked.Location = New System.Drawing.Point(760, 99)
        Me.TxtBooked.MaxLength = 150
        Me.TxtBooked.Name = "TxtBooked"
        Me.TxtBooked.ReadOnly = True
        Me.TxtBooked.Size = New System.Drawing.Size(74, 20)
        Me.TxtBooked.TabIndex = 326
        Me.TxtBooked.TabStop = False
        Me.TxtBooked.Text = "0"
        Me.TxtBooked.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtBooked.Visible = False
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(689, 102)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(44, 13)
        Me.Label18.TabIndex = 327
        Me.Label18.Text = "Booked"
        Me.Label18.Visible = False
        '
        'lblWarehouse
        '
        Me.lblWarehouse.AutoSize = True
        Me.lblWarehouse.Location = New System.Drawing.Point(560, 48)
        Me.lblWarehouse.Name = "lblWarehouse"
        Me.lblWarehouse.Size = New System.Drawing.Size(24, 13)
        Me.lblWarehouse.TabIndex = 325
        Me.lblWarehouse.Text = "n/a"
        Me.lblWarehouse.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = True
        Me.lblLocation.Location = New System.Drawing.Point(560, 22)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(24, 13)
        Me.lblLocation.TabIndex = 324
        Me.lblLocation.Text = "n/a"
        Me.lblLocation.Visible = False
        '
        'cboLocation
        '
        Me.cboLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.FormattingEnabled = True
        Me.cboLocation.Location = New System.Drawing.Point(560, 17)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Size = New System.Drawing.Size(386, 21)
        Me.cboLocation.TabIndex = 319
        '
        'txtPrimaryLocation
        '
        Me.txtPrimaryLocation.Location = New System.Drawing.Point(529, 19)
        Me.txtPrimaryLocation.Name = "txtPrimaryLocation"
        Me.txtPrimaryLocation.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryLocation.TabIndex = 323
        Me.txtPrimaryLocation.TabStop = False
        Me.txtPrimaryLocation.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(456, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 13)
        Me.Label7.TabIndex = 322
        Me.Label7.Text = "Location"
        '
        'TxtWarehouse
        '
        Me.TxtWarehouse.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouse.Location = New System.Drawing.Point(560, 44)
        Me.TxtWarehouse.MaxLength = 50
        Me.TxtWarehouse.Name = "TxtWarehouse"
        Me.TxtWarehouse.ReadOnly = True
        Me.TxtWarehouse.Size = New System.Drawing.Size(386, 20)
        Me.TxtWarehouse.TabIndex = 320
        Me.TxtWarehouse.TabStop = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(456, 44)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 321
        Me.Label10.Text = "Warehouse"
        '
        'dtpValid1
        '
        Me.dtpValid1.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpValid1.Enabled = False
        Me.dtpValid1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpValid1.Location = New System.Drawing.Point(108, 74)
        Me.dtpValid1.Name = "dtpValid1"
        Me.dtpValid1.Size = New System.Drawing.Size(312, 20)
        Me.dtpValid1.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 76)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 13)
        Me.Label6.TabIndex = 318
        Me.Label6.Text = "Valid Until"
        '
        'lblAvailableUnits
        '
        Me.lblAvailableUnits.AutoSize = True
        Me.lblAvailableUnits.Location = New System.Drawing.Point(646, 103)
        Me.lblAvailableUnits.Name = "lblAvailableUnits"
        Me.lblAvailableUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblAvailableUnits.TabIndex = 316
        Me.lblAvailableUnits.Text = "Units"
        '
        'txtAvailable
        '
        Me.txtAvailable.Location = New System.Drawing.Point(560, 100)
        Me.txtAvailable.MaxLength = 150
        Me.txtAvailable.Name = "txtAvailable"
        Me.txtAvailable.ReadOnly = True
        Me.txtAvailable.Size = New System.Drawing.Size(80, 20)
        Me.txtAvailable.TabIndex = 314
        Me.txtAvailable.TabStop = False
        Me.txtAvailable.Text = "0"
        Me.txtAvailable.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(459, 103)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(89, 13)
        Me.Label15.TabIndex = 315
        Me.Label15.Text = "Available to book"
        '
        'lblUnitsPendingDO
        '
        Me.lblUnitsPendingDO.AutoSize = True
        Me.lblUnitsPendingDO.Location = New System.Drawing.Point(840, 77)
        Me.lblUnitsPendingDO.Name = "lblUnitsPendingDO"
        Me.lblUnitsPendingDO.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsPendingDO.TabIndex = 313
        Me.lblUnitsPendingDO.Text = "Units"
        Me.lblUnitsPendingDO.Visible = False
        '
        'txtPendingDO
        '
        Me.txtPendingDO.Location = New System.Drawing.Point(760, 73)
        Me.txtPendingDO.MaxLength = 150
        Me.txtPendingDO.Name = "txtPendingDO"
        Me.txtPendingDO.ReadOnly = True
        Me.txtPendingDO.Size = New System.Drawing.Size(74, 20)
        Me.txtPendingDO.TabIndex = 311
        Me.txtPendingDO.TabStop = False
        Me.txtPendingDO.Text = "0"
        Me.txtPendingDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPendingDO.Visible = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(689, 76)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(65, 13)
        Me.Label14.TabIndex = 312
        Me.Label14.Text = "Pending DO"
        Me.Label14.Visible = False
        '
        'lblUnitFree
        '
        Me.lblUnitFree.AutoSize = True
        Me.lblUnitFree.Location = New System.Drawing.Point(646, 78)
        Me.lblUnitFree.Name = "lblUnitFree"
        Me.lblUnitFree.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitFree.TabIndex = 308
        Me.lblUnitFree.Text = "Units"
        Me.lblUnitFree.Visible = False
        '
        'txtQtyOnHand
        '
        Me.txtQtyOnHand.Location = New System.Drawing.Point(560, 74)
        Me.txtQtyOnHand.MaxLength = 150
        Me.txtQtyOnHand.Name = "txtQtyOnHand"
        Me.txtQtyOnHand.ReadOnly = True
        Me.txtQtyOnHand.Size = New System.Drawing.Size(80, 20)
        Me.txtQtyOnHand.TabIndex = 306
        Me.txtQtyOnHand.TabStop = False
        Me.txtQtyOnHand.Text = "0"
        Me.txtQtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtQtyOnHand.Visible = False
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(459, 77)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(69, 13)
        Me.Label20.TabIndex = 307
        Me.Label20.Text = "Qty On Hand"
        Me.Label20.Visible = False
        '
        'LblUnitOB
        '
        Me.LblUnitOB.AutoSize = True
        Me.LblUnitOB.Location = New System.Drawing.Point(370, 103)
        Me.LblUnitOB.Name = "LblUnitOB"
        Me.LblUnitOB.Size = New System.Drawing.Size(31, 13)
        Me.LblUnitOB.TabIndex = 305
        Me.LblUnitOB.Text = "Units"
        '
        'txtQtyOrderBalance
        '
        Me.txtQtyOrderBalance.Location = New System.Drawing.Point(108, 100)
        Me.txtQtyOrderBalance.MaxLength = 150
        Me.txtQtyOrderBalance.Name = "txtQtyOrderBalance"
        Me.txtQtyOrderBalance.ReadOnly = True
        Me.txtQtyOrderBalance.Size = New System.Drawing.Size(259, 20)
        Me.txtQtyOrderBalance.TabIndex = 303
        Me.txtQtyOrderBalance.TabStop = False
        Me.txtQtyOrderBalance.Text = "0"
        Me.txtQtyOrderBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 103)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 13)
        Me.Label8.TabIndex = 304
        Me.Label8.Text = "Qty SO to book"
        '
        'llbDetailStock
        '
        Me.llbDetailStock.AutoSize = True
        Me.llbDetailStock.Location = New System.Drawing.Point(363, 25)
        Me.llbDetailStock.Name = "llbDetailStock"
        Me.llbDetailStock.Size = New System.Drawing.Size(65, 13)
        Me.llbDetailStock.TabIndex = 1
        Me.llbDetailStock.TabStop = True
        Me.llbDetailStock.Text = "Show Stock"
        '
        'chkLocationControlled
        '
        Me.chkLocationControlled.AutoSize = True
        Me.chkLocationControlled.Location = New System.Drawing.Point(253, 241)
        Me.chkLocationControlled.Name = "chkLocationControlled"
        Me.chkLocationControlled.Size = New System.Drawing.Size(67, 17)
        Me.chkLocationControlled.TabIndex = 302
        Me.chkLocationControlled.TabStop = False
        Me.chkLocationControlled.Text = "Location"
        Me.chkLocationControlled.UseVisualStyleBackColor = True
        Me.chkLocationControlled.Visible = False
        '
        'chkPdoControlled
        '
        Me.chkPdoControlled.AutoSize = True
        Me.chkPdoControlled.Location = New System.Drawing.Point(253, 264)
        Me.chkPdoControlled.Name = "chkPdoControlled"
        Me.chkPdoControlled.Size = New System.Drawing.Size(47, 17)
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
        Me.cboSecondaryQty.Location = New System.Drawing.Point(798, 126)
        Me.cboSecondaryQty.Name = "cboSecondaryQty"
        Me.cboSecondaryQty.Size = New System.Drawing.Size(94, 21)
        Me.cboSecondaryQty.TabIndex = 6
        '
        'cboBatchNo
        '
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(108, 126)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(312, 21)
        Me.cboBatchNo.TabIndex = 3
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(44, 17)
        Me.txtPrimaryCategory.MaxLength = 30
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryCategory.TabIndex = 4
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Text = "-1"
        Me.txtPrimaryCategory.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCategory.Visible = False
        '
        'txtPrimarySecondaryUnits
        '
        Me.txtPrimarySecondaryUnits.Location = New System.Drawing.Point(623, 212)
        Me.txtPrimarySecondaryUnits.MaxLength = 30
        Me.txtPrimarySecondaryUnits.Name = "txtPrimarySecondaryUnits"
        Me.txtPrimarySecondaryUnits.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimarySecondaryUnits.TabIndex = 296
        Me.txtPrimarySecondaryUnits.TabStop = False
        Me.txtPrimarySecondaryUnits.Text = "-1"
        Me.txtPrimarySecondaryUnits.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimarySecondaryUnits.Visible = False
        '
        'txtPrimaryUnits
        '
        Me.txtPrimaryUnits.Location = New System.Drawing.Point(571, 214)
        Me.txtPrimaryUnits.MaxLength = 30
        Me.txtPrimaryUnits.Name = "txtPrimaryUnits"
        Me.txtPrimaryUnits.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryUnits.TabIndex = 295
        Me.txtPrimaryUnits.TabStop = False
        Me.txtPrimaryUnits.Text = "-1"
        Me.txtPrimaryUnits.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnits.Visible = False
        '
        'lblSecQty
        '
        Me.lblSecQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQty.AutoSize = True
        Me.lblSecQty.Location = New System.Drawing.Point(774, 129)
        Me.lblSecQty.Name = "lblSecQty"
        Me.lblSecQty.Size = New System.Drawing.Size(18, 13)
        Me.lblSecQty.TabIndex = 292
        Me.lblSecQty.Text = "@"
        '
        'txtSecondaryQty
        '
        Me.txtSecondaryQty.Location = New System.Drawing.Point(519, 126)
        Me.txtSecondaryQty.MaxLength = 30
        Me.txtSecondaryQty.Name = "txtSecondaryQty"
        Me.txtSecondaryQty.Size = New System.Drawing.Size(35, 20)
        Me.txtSecondaryQty.TabIndex = 288
        Me.txtSecondaryQty.TabStop = False
        Me.txtSecondaryQty.Text = "0"
        Me.txtSecondaryQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSecondaryQty.Visible = False
        '
        'lblUnits
        '
        Me.lblUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnits.AutoSize = True
        Me.lblUnits.Location = New System.Drawing.Point(736, 129)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblUnits.TabIndex = 290
        Me.lblUnits.Text = "Units"
        '
        'TxtQtyperBatch
        '
        Me.TxtQtyperBatch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatch.Location = New System.Drawing.Point(560, 126)
        Me.TxtQtyperBatch.MaxLength = 18
        Me.TxtQtyperBatch.Name = "TxtQtyperBatch"
        Me.TxtQtyperBatch.Size = New System.Drawing.Size(170, 20)
        Me.TxtQtyperBatch.TabIndex = 5
        Me.TxtQtyperBatch.Text = "0"
        Me.TxtQtyperBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(459, 130)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(23, 13)
        Me.Label4.TabIndex = 289
        Me.Label4.Text = "Qty"
        '
        'lblSecondaryUnits
        '
        Me.lblSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnits.AutoSize = True
        Me.lblSecondaryUnits.Location = New System.Drawing.Point(894, 129)
        Me.lblSecondaryUnits.Name = "lblSecondaryUnits"
        Me.lblSecondaryUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnits.TabIndex = 291
        Me.lblSecondaryUnits.Text = "Units"
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Location = New System.Drawing.Point(108, 129)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNo.TabIndex = 285
        Me.lblBatchNo.Text = "n/a"
        Me.lblBatchNo.Visible = False
        '
        'lblExpiryDate
        '
        Me.lblExpiryDate.AutoSize = True
        Me.lblExpiryDate.Location = New System.Drawing.Point(108, 155)
        Me.lblExpiryDate.Name = "lblExpiryDate"
        Me.lblExpiryDate.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDate.TabIndex = 284
        Me.lblExpiryDate.Text = "n/a"
        Me.lblExpiryDate.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 283
        Me.Label5.Text = "Batch No"
        '
        'DtpExpired
        '
        Me.DtpExpired.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpired.Enabled = False
        Me.DtpExpired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpired.Location = New System.Drawing.Point(108, 152)
        Me.DtpExpired.Name = "DtpExpired"
        Me.DtpExpired.Size = New System.Drawing.Size(312, 20)
        Me.DtpExpired.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(9, 154)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 282
        Me.Label12.Text = "Expiry Date"
        '
        'chkExpiryDateControlled
        '
        Me.chkExpiryDateControlled.AutoSize = True
        Me.chkExpiryDateControlled.Location = New System.Drawing.Point(253, 201)
        Me.chkExpiryDateControlled.Name = "chkExpiryDateControlled"
        Me.chkExpiryDateControlled.Size = New System.Drawing.Size(80, 17)
        Me.chkExpiryDateControlled.TabIndex = 279
        Me.chkExpiryDateControlled.TabStop = False
        Me.chkExpiryDateControlled.Text = "Expiry Date"
        Me.chkExpiryDateControlled.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlled.Visible = False
        '
        'chkBatchNoControlled
        '
        Me.chkBatchNoControlled.AutoSize = True
        Me.chkBatchNoControlled.Location = New System.Drawing.Point(253, 178)
        Me.chkBatchNoControlled.Name = "chkBatchNoControlled"
        Me.chkBatchNoControlled.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlled.TabIndex = 278
        Me.chkBatchNoControlled.TabStop = False
        Me.chkBatchNoControlled.Text = "Batch No"
        Me.chkBatchNoControlled.UseVisualStyleBackColor = True
        Me.chkBatchNoControlled.Visible = False
        '
        'chkUseSecondaryUnit
        '
        Me.chkUseSecondaryUnit.AutoSize = True
        Me.chkUseSecondaryUnit.Location = New System.Drawing.Point(253, 220)
        Me.chkUseSecondaryUnit.Name = "chkUseSecondaryUnit"
        Me.chkUseSecondaryUnit.Size = New System.Drawing.Size(121, 17)
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
        Me.chkUseCustCodeNo.Location = New System.Drawing.Point(20, 215)
        Me.chkUseCustCodeNo.Name = "chkUseCustCodeNo"
        Me.chkUseCustCodeNo.Size = New System.Drawing.Size(122, 17)
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
        Me.cboCustCodeNo.Location = New System.Drawing.Point(148, 213)
        Me.cboCustCodeNo.Name = "cboCustCodeNo"
        Me.cboCustCodeNo.Size = New System.Drawing.Size(280, 21)
        Me.cboCustCodeNo.TabIndex = 8
        Me.cboCustCodeNo.TabStop = False
        Me.cboCustCodeNo.Visible = False
        '
        'txtCustProductDesc
        '
        Me.txtCustProductDesc.Location = New System.Drawing.Point(148, 242)
        Me.txtCustProductDesc.MaxLength = 150
        Me.txtCustProductDesc.Name = "txtCustProductDesc"
        Me.txtCustProductDesc.ReadOnly = True
        Me.txtCustProductDesc.Size = New System.Drawing.Size(280, 20)
        Me.txtCustProductDesc.TabIndex = 7
        Me.txtCustProductDesc.TabStop = False
        Me.txtCustProductDesc.Visible = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(17, 245)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(114, 13)
        Me.Label19.TabIndex = 266
        Me.Label19.Text = "Customer's Description"
        Me.Label19.Visible = False
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(86, 17)
        Me.txtPrimaryCodeNo.MaxLength = 30
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(37, 20)
        Me.txtPrimaryCodeNo.TabIndex = 3
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Text = "-1"
        Me.txtPrimaryCodeNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCodeNo.Visible = False
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(560, 152)
        Me.TxtRemarkItem.MaxLength = 150
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(332, 20)
        Me.TxtRemarkItem.TabIndex = 7
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(459, 154)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 13)
        Me.Label16.TabIndex = 226
        Me.Label16.Text = "Remarks"
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(898, 149)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(52, 25)
        Me.BtnInput.TabIndex = 8
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(108, 48)
        Me.txtDesc.MaxLength = 150
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(312, 20)
        Me.txtDesc.TabIndex = 7
        Me.txtDesc.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(9, 51)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 197
        Me.Label11.Text = "Description"
        '
        'lblCodeNo
        '
        Me.lblCodeNo.AutoSize = True
        Me.lblCodeNo.Location = New System.Drawing.Point(9, 25)
        Me.lblCodeNo.Name = "lblCodeNo"
        Me.lblCodeNo.Size = New System.Drawing.Size(49, 13)
        Me.lblCodeNo.TabIndex = 196
        Me.lblCodeNo.Text = "Code No"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 25)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 195
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'cboCodeNo
        '
        Me.cboCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeNo.FormattingEnabled = True
        Me.cboCodeNo.Location = New System.Drawing.Point(108, 19)
        Me.cboCodeNo.Name = "cboCodeNo"
        Me.cboCodeNo.Size = New System.Drawing.Size(249, 21)
        Me.cboCodeNo.TabIndex = 0
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWarehouse, Me.ColValidUntil, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits, Me.ColPrimaryLocation})
        Me.dgvItem.Location = New System.Drawing.Point(6, 311)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(952, 136)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 7
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
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 56
        '
        'ColSecondaryQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecondaryQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSecondaryQty.HeaderText = "@"
        Me.ColSecondaryQty.Name = "ColSecondaryQty"
        Me.ColSecondaryQty.Width = 43
        '
        'ColSecondaryUnit
        '
        Me.ColSecondaryUnit.HeaderText = "Units"
        Me.ColSecondaryUnit.Name = "ColSecondaryUnit"
        Me.ColSecondaryUnit.Width = 56
        '
        'ColBatch
        '
        Me.ColBatch.HeaderText = "Batch No"
        Me.ColBatch.Name = "ColBatch"
        Me.ColBatch.ReadOnly = True
        Me.ColBatch.Width = 77
        '
        'ColExpiredDate
        '
        Me.ColExpiredDate.HeaderText = "Expiry Date"
        Me.ColExpiredDate.Name = "ColExpiredDate"
        Me.ColExpiredDate.Width = 86
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 73
        '
        'ColWarehouse
        '
        Me.ColWarehouse.HeaderText = "Warehouse"
        Me.ColWarehouse.Name = "ColWarehouse"
        Me.ColWarehouse.Width = 87
        '
        'ColValidUntil
        '
        DataGridViewCellStyle4.Format = "D"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.ColValidUntil.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColValidUntil.HeaderText = "Valid Until"
        Me.ColValidUntil.Name = "ColValidUntil"
        Me.ColValidUntil.Width = 79
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'ColPrimarySecondaryUnits
        '
        Me.ColPrimarySecondaryUnits.HeaderText = "PrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Name = "ColPrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Visible = False
        Me.ColPrimarySecondaryUnits.Width = 141
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "PrimaryLocation"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 107
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.Location = New System.Drawing.Point(125, 97)
        Me.TxtCustomerCode.MaxLength = 50
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.ReadOnly = True
        Me.TxtCustomerCode.Size = New System.Drawing.Size(273, 20)
        Me.TxtCustomerCode.TabIndex = 225
        Me.TxtCustomerCode.TabStop = False
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(95, 97)
        Me.TxtPrimaryCustomer.MaxLength = 50
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(24, 20)
        Me.TxtPrimaryCustomer.TabIndex = 229
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 228
        Me.Label3.Text = "Customer Code"
        '
        'BtnFindSONumber
        '
        Me.BtnFindSONumber.Location = New System.Drawing.Point(363, 70)
        Me.BtnFindSONumber.Name = "BtnFindSONumber"
        Me.BtnFindSONumber.Size = New System.Drawing.Size(35, 20)
        Me.BtnFindSONumber.TabIndex = 3
        Me.BtnFindSONumber.TabStop = False
        Me.BtnFindSONumber.Text = "...."
        Me.BtnFindSONumber.UseVisualStyleBackColor = True
        '
        'TxtPrimarySONumber
        '
        Me.TxtPrimarySONumber.Location = New System.Drawing.Point(95, 72)
        Me.TxtPrimarySONumber.MaxLength = 50
        Me.TxtPrimarySONumber.Name = "TxtPrimarySONumber"
        Me.TxtPrimarySONumber.Size = New System.Drawing.Size(24, 20)
        Me.TxtPrimarySONumber.TabIndex = 227
        Me.TxtPrimarySONumber.TabStop = False
        Me.TxtPrimarySONumber.Visible = False
        '
        'TxtSONumber
        '
        Me.TxtSONumber.Location = New System.Drawing.Point(125, 71)
        Me.TxtSONumber.Name = "TxtSONumber"
        Me.TxtSONumber.Size = New System.Drawing.Size(232, 20)
        Me.TxtSONumber.TabIndex = 2
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(13, 72)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(62, 13)
        Me.Label13.TabIndex = 226
        Me.Label13.Text = "SO Number"
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 453)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 221
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 453)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 222
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(501, 453)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(385, 453)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(482, 51)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(472, 71)
        Me.TxtRemark.TabIndex = 5
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(422, 51)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(44, 13)
        Me.Label37.TabIndex = 219
        Me.Label37.Text = "Remark"
        '
        'cboValidUntil
        '
        Me.cboValidUntil.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboValidUntil.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboValidUntil.FormattingEnabled = True
        Me.cboValidUntil.Items.AddRange(New Object() {"1 day after booking date", "2 days after booking date", "3 days after booking date"})
        Me.cboValidUntil.Location = New System.Drawing.Point(482, 19)
        Me.cboValidUntil.Name = "cboValidUntil"
        Me.cboValidUntil.Size = New System.Drawing.Size(472, 21)
        Me.cboValidUntil.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(422, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 216
        Me.Label1.Text = "Valid Until"
        '
        'DtpCreatedDate
        '
        Me.DtpCreatedDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpCreatedDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpCreatedDate.Location = New System.Drawing.Point(125, 45)
        Me.DtpCreatedDate.Name = "DtpCreatedDate"
        Me.DtpCreatedDate.Size = New System.Drawing.Size(273, 20)
        Me.DtpCreatedDate.TabIndex = 1
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(13, 45)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(30, 13)
        Me.Label21.TabIndex = 215
        Me.Label21.Text = "Date"
        '
        'TxtBookingNo
        '
        Me.TxtBookingNo.Location = New System.Drawing.Point(125, 19)
        Me.TxtBookingNo.Name = "TxtBookingNo"
        Me.TxtBookingNo.ReadOnly = True
        Me.TxtBookingNo.Size = New System.Drawing.Size(273, 20)
        Me.TxtBookingNo.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Booking No"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedDate, Me.tslModifiedBy, Me.tslStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 540)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(984, 22)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(193, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(193, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(193, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(193, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslStatus
        '
        Me.tslStatus.Name = "tslStatus"
        Me.tslStatus.Size = New System.Drawing.Size(193, 17)
        Me.tslStatus.Spring = True
        Me.tslStatus.Text = "Status: "
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(634, 503)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(549, 503)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(457, 503)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(275, 503)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(366, 503)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'FrmTransaksiBookingInv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(984, 562)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "FrmTransaksiBookingInv"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Booking"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.gboItem.ResumeLayout(False)
        Me.gboItem.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtBookingNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpCreatedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboValidUntil As System.Windows.Forms.ComboBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnFindSONumber As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySONumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtSONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents dtpValid1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblAvailableUnits As System.Windows.Forms.Label
    Friend WithEvents txtAvailable As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblUnitsPendingDO As System.Windows.Forms.Label
    Friend WithEvents txtPendingDO As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblUnitFree As System.Windows.Forms.Label
    Friend WithEvents txtQtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents LblUnitOB As System.Windows.Forms.Label
    Friend WithEvents txtQtyOrderBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents llbDetailStock As System.Windows.Forms.LinkLabel
    Friend WithEvents chkLocationControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkPdoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents cboSecondaryQty As System.Windows.Forms.ComboBox
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySecondaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents lblSecQty As System.Windows.Forms.Label
    Friend WithEvents txtSecondaryQty As System.Windows.Forms.TextBox
    Friend WithEvents lblUnits As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatch As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblSecondaryUnits As System.Windows.Forms.Label
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDate As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
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
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblCodeNo As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryBooking As System.Windows.Forms.TextBox
    Friend WithEvents cboLocation As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblWarehouse As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColValidUntil As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblUnitBooked As System.Windows.Forms.Label
    Friend WithEvents TxtBooked As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
End Class
