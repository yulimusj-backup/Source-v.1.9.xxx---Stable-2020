<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPOCOtorisasi
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
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.TsbExit = New System.Windows.Forms.ToolStripButton
        Me.TsbAuthorize = New System.Windows.Forms.ToolStripButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtCancellationReason = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtPrimaryPO = New System.Windows.Forms.TextBox
        Me.txtPONO = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtPhoneNo = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtContactPerson = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.llbAuthorization = New System.Windows.Forms.LinkLabel
        Me.txtAutohirizationStatus = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryPOC = New System.Windows.Forms.TextBox
        Me.TxtPOCNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabBlmOtorisasi = New System.Windows.Forms.TabPage
        Me.llbUncheckAllAuthorize = New System.Windows.Forms.LinkLabel
        Me.llbCheckAllAuthorize = New System.Windows.Forms.LinkLabel
        Me.chkBelumOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabCancel = New System.Windows.Forms.TabPage
        Me.llbUncheckAllCancel = New System.Windows.Forms.LinkLabel
        Me.llbCheckAllCancel = New System.Windows.Forms.LinkLabel
        Me.chkSudahOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.ToolStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabBlmOtorisasi.SuspendLayout()
        Me.TabCancel.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TsbExit, Me.TsbAuthorize})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 611)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip1.Size = New System.Drawing.Size(884, 30)
        Me.ToolStrip1.TabIndex = 2
        '
        'TsbExit
        '
        Me.TsbExit.Image = Global.Purchase_Module.My.Resources.Resources.cancel1
        Me.TsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsbExit.Name = "TsbExit"
        Me.TsbExit.Size = New System.Drawing.Size(59, 27)
        Me.TsbExit.Text = "E&xit"
        Me.TsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'TsbAuthorize
        '
        Me.TsbAuthorize.Image = Global.Purchase_Module.My.Resources.Resources.lock_closed1
        Me.TsbAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsbAuthorize.Name = "TsbAuthorize"
        Me.TsbAuthorize.Size = New System.Drawing.Size(92, 27)
        Me.TsbAuthorize.Text = "Authorize"
        Me.TsbAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 641)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 50
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.StatusStrip2)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.txtPrimarySupplier)
        Me.gboDetail.Controls.Add(Me.txtCancellationReason)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.txtPrimaryPO)
        Me.gboDetail.Controls.Add(Me.txtPONO)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.txtSupplierCode)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label23)
        Me.gboDetail.Controls.Add(Me.txtPhoneNo)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.txtContactPerson)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtSupplierName)
        Me.gboDetail.Controls.Add(Me.llbAuthorization)
        Me.gboDetail.Controls.Add(Me.txtAutohirizationStatus)
        Me.gboDetail.Controls.Add(Me.Label21)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryPOC)
        Me.gboDetail.Controls.Add(Me.TxtPOCNo)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Location = New System.Drawing.Point(12, 149)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(860, 459)
        Me.gboDetail.TabIndex = 1
        Me.gboDetail.TabStop = False
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(3, 434)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(854, 22)
        Me.StatusStrip2.TabIndex = 188
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(209, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(209, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(209, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(209, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColMasterUnit, Me.ColQty, Me.ColTotalQty, Me.ColUnit, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryMasterUnit, Me.ColPrimaryUnits})
        Me.dgvItem.Location = New System.Drawing.Point(6, 173)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(845, 258)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 187
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
        'ColQtyMaster
        '
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 83
        '
        'ColQtyPerMaster
        '
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 102
        '
        'ColMasterUnit
        '
        Me.ColMasterUnit.HeaderText = "Master Unit"
        Me.ColMasterUnit.Name = "ColMasterUnit"
        Me.ColMasterUnit.Visible = False
        Me.ColMasterUnit.Width = 86
        '
        'ColQty
        '
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Visible = False
        Me.ColQty.Width = 48
        '
        'ColTotalQty
        '
        Me.ColTotalQty.HeaderText = "Total Qty"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.Width = 75
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 56
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
        'ColPrimaryMasterUnit
        '
        Me.ColPrimaryMasterUnit.HeaderText = "PrimaryMasterUnit"
        Me.ColPrimaryMasterUnit.Name = "ColPrimaryMasterUnit"
        Me.ColPrimaryMasterUnit.Visible = False
        Me.ColPrimaryMasterUnit.Width = 117
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(110, 95)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySupplier.TabIndex = 186
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Visible = False
        '
        'txtCancellationReason
        '
        Me.txtCancellationReason.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCancellationReason.Location = New System.Drawing.Point(543, 72)
        Me.txtCancellationReason.MaxLength = 150
        Me.txtCancellationReason.Multiline = True
        Me.txtCancellationReason.Name = "txtCancellationReason"
        Me.txtCancellationReason.ReadOnly = True
        Me.txtCancellationReason.Size = New System.Drawing.Size(308, 43)
        Me.txtCancellationReason.TabIndex = 172
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(439, 75)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(105, 13)
        Me.Label8.TabIndex = 185
        Me.Label8.Text = "Cancellation Reason"
        '
        'txtPrimaryPO
        '
        Me.txtPrimaryPO.Location = New System.Drawing.Point(182, 72)
        Me.txtPrimaryPO.Name = "txtPrimaryPO"
        Me.txtPrimaryPO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryPO.TabIndex = 184
        Me.txtPrimaryPO.TabStop = False
        Me.txtPrimaryPO.Visible = False
        '
        'txtPONO
        '
        Me.txtPONO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPONO.Location = New System.Drawing.Point(110, 72)
        Me.txtPONO.MaxLength = 50
        Me.txtPONO.Name = "txtPONO"
        Me.txtPONO.ReadOnly = True
        Me.txtPONO.Size = New System.Drawing.Size(319, 20)
        Me.txtPONO.TabIndex = 163
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 75)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 13)
        Me.Label5.TabIndex = 183
        Me.Label5.Text = "Purchase Order No"
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.Location = New System.Drawing.Point(110, 94)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.ReadOnly = True
        Me.txtSupplierCode.Size = New System.Drawing.Size(319, 20)
        Me.txtSupplierCode.TabIndex = 166
        Me.txtSupplierCode.TabStop = False
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(543, 121)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.ReadOnly = True
        Me.TxtRemark.Size = New System.Drawing.Size(308, 43)
        Me.TxtRemark.TabIndex = 173
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(439, 121)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 182
        Me.Label6.Text = "Remarks"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(439, 26)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(55, 13)
        Me.Label23.TabIndex = 181
        Me.Label23.Text = "Phone No"
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPhoneNo.Location = New System.Drawing.Point(543, 23)
        Me.txtPhoneNo.MaxLength = 50
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.ReadOnly = True
        Me.txtPhoneNo.Size = New System.Drawing.Size(308, 20)
        Me.txtPhoneNo.TabIndex = 169
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 150)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 180
        Me.Label7.Text = "Contact Person"
        '
        'txtContactPerson
        '
        Me.txtContactPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtContactPerson.Location = New System.Drawing.Point(110, 147)
        Me.txtContactPerson.MaxLength = 150
        Me.txtContactPerson.Name = "txtContactPerson"
        Me.txtContactPerson.ReadOnly = True
        Me.txtContactPerson.Size = New System.Drawing.Size(319, 20)
        Me.txtContactPerson.TabIndex = 168
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 123)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 179
        Me.Label2.Text = "Supplier Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 98)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 178
        Me.Label1.Text = "Supplier Code"
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Location = New System.Drawing.Point(110, 121)
        Me.TxtSupplierName.MaxLength = 150
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(319, 20)
        Me.TxtSupplierName.TabIndex = 167
        Me.TxtSupplierName.TabStop = False
        '
        'llbAuthorization
        '
        Me.llbAuthorization.AutoSize = True
        Me.llbAuthorization.Location = New System.Drawing.Point(794, 52)
        Me.llbAuthorization.Name = "llbAuthorization"
        Me.llbAuthorization.Size = New System.Drawing.Size(57, 13)
        Me.llbAuthorization.TabIndex = 171
        Me.llbAuthorization.TabStop = True
        Me.llbAuthorization.Text = "view detail"
        '
        'txtAutohirizationStatus
        '
        Me.txtAutohirizationStatus.Location = New System.Drawing.Point(543, 49)
        Me.txtAutohirizationStatus.MaxLength = 50
        Me.txtAutohirizationStatus.Name = "txtAutohirizationStatus"
        Me.txtAutohirizationStatus.ReadOnly = True
        Me.txtAutohirizationStatus.Size = New System.Drawing.Size(245, 20)
        Me.txtAutohirizationStatus.TabIndex = 170
        Me.txtAutohirizationStatus.TabStop = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(439, 52)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(101, 13)
        Me.Label21.TabIndex = 177
        Me.Label21.Text = "Authorization Status"
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(110, 46)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(319, 20)
        Me.dtpDate.TabIndex = 162
        '
        'TxtPrimaryPOC
        '
        Me.TxtPrimaryPOC.Location = New System.Drawing.Point(168, 20)
        Me.TxtPrimaryPOC.Name = "TxtPrimaryPOC"
        Me.TxtPrimaryPOC.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryPOC.TabIndex = 176
        Me.TxtPrimaryPOC.TabStop = False
        Me.TxtPrimaryPOC.Visible = False
        '
        'TxtPOCNo
        '
        Me.TxtPOCNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPOCNo.Location = New System.Drawing.Point(110, 23)
        Me.TxtPOCNo.MaxLength = 50
        Me.TxtPOCNo.Name = "TxtPOCNo"
        Me.TxtPOCNo.ReadOnly = True
        Me.TxtPOCNo.Size = New System.Drawing.Size(319, 20)
        Me.TxtPOCNo.TabIndex = 161
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 175
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 174
        Me.Label4.Text = "POC No"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabBlmOtorisasi)
        Me.TabControl1.Controls.Add(Me.TabCancel)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(860, 131)
        Me.TabControl1.TabIndex = 51
        '
        'TabBlmOtorisasi
        '
        Me.TabBlmOtorisasi.Controls.Add(Me.llbUncheckAllAuthorize)
        Me.TabBlmOtorisasi.Controls.Add(Me.llbCheckAllAuthorize)
        Me.TabBlmOtorisasi.Controls.Add(Me.chkBelumOtorisasi)
        Me.TabBlmOtorisasi.Controls.Add(Me.btnExit)
        Me.TabBlmOtorisasi.Controls.Add(Me.Button1)
        Me.TabBlmOtorisasi.Location = New System.Drawing.Point(4, 22)
        Me.TabBlmOtorisasi.Name = "TabBlmOtorisasi"
        Me.TabBlmOtorisasi.Padding = New System.Windows.Forms.Padding(3)
        Me.TabBlmOtorisasi.Size = New System.Drawing.Size(852, 105)
        Me.TabBlmOtorisasi.TabIndex = 0
        Me.TabBlmOtorisasi.Text = "Authorize"
        Me.TabBlmOtorisasi.UseVisualStyleBackColor = True
        '
        'llbUncheckAllAuthorize
        '
        Me.llbUncheckAllAuthorize.AutoSize = True
        Me.llbUncheckAllAuthorize.Location = New System.Drawing.Point(86, 92)
        Me.llbUncheckAllAuthorize.Name = "llbUncheckAllAuthorize"
        Me.llbUncheckAllAuthorize.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAllAuthorize.TabIndex = 45
        Me.llbUncheckAllAuthorize.TabStop = True
        Me.llbUncheckAllAuthorize.Text = "Uncheck All"
        '
        'llbCheckAllAuthorize
        '
        Me.llbCheckAllAuthorize.AutoSize = True
        Me.llbCheckAllAuthorize.Location = New System.Drawing.Point(26, 92)
        Me.llbCheckAllAuthorize.Name = "llbCheckAllAuthorize"
        Me.llbCheckAllAuthorize.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAllAuthorize.TabIndex = 44
        Me.llbCheckAllAuthorize.TabStop = True
        Me.llbCheckAllAuthorize.Text = "Check All"
        '
        'chkBelumOtorisasi
        '
        Me.chkBelumOtorisasi.CheckOnClick = True
        Me.chkBelumOtorisasi.ColumnWidth = 200
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
        Me.btnExit.Location = New System.Drawing.Point(652, 37)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(62, 29)
        Me.btnExit.TabIndex = 43
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(356, 33)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 32)
        Me.Button1.TabIndex = 46
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TabCancel
        '
        Me.TabCancel.Controls.Add(Me.llbUncheckAllCancel)
        Me.TabCancel.Controls.Add(Me.llbCheckAllCancel)
        Me.TabCancel.Controls.Add(Me.chkSudahOtorisasi)
        Me.TabCancel.Location = New System.Drawing.Point(4, 22)
        Me.TabCancel.Name = "TabCancel"
        Me.TabCancel.Padding = New System.Windows.Forms.Padding(3)
        Me.TabCancel.Size = New System.Drawing.Size(852, 105)
        Me.TabCancel.TabIndex = 1
        Me.TabCancel.Text = "Cancel Authorize"
        Me.TabCancel.UseVisualStyleBackColor = True
        '
        'llbUncheckAllCancel
        '
        Me.llbUncheckAllCancel.AutoSize = True
        Me.llbUncheckAllCancel.Location = New System.Drawing.Point(86, 92)
        Me.llbUncheckAllCancel.Name = "llbUncheckAllCancel"
        Me.llbUncheckAllCancel.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAllCancel.TabIndex = 3
        Me.llbUncheckAllCancel.TabStop = True
        Me.llbUncheckAllCancel.Text = "Uncheck All"
        '
        'llbCheckAllCancel
        '
        Me.llbCheckAllCancel.AutoSize = True
        Me.llbCheckAllCancel.Location = New System.Drawing.Point(26, 92)
        Me.llbCheckAllCancel.Name = "llbCheckAllCancel"
        Me.llbCheckAllCancel.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAllCancel.TabIndex = 2
        Me.llbCheckAllCancel.TabStop = True
        Me.llbCheckAllCancel.Text = "Check All"
        '
        'chkSudahOtorisasi
        '
        Me.chkSudahOtorisasi.CheckOnClick = True
        Me.chkSudahOtorisasi.ColumnWidth = 200
        Me.chkSudahOtorisasi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkSudahOtorisasi.FormattingEnabled = True
        Me.chkSudahOtorisasi.Location = New System.Drawing.Point(3, 3)
        Me.chkSudahOtorisasi.MultiColumn = True
        Me.chkSudahOtorisasi.Name = "chkSudahOtorisasi"
        Me.chkSudahOtorisasi.Size = New System.Drawing.Size(846, 94)
        Me.chkSudahOtorisasi.TabIndex = 1
        '
        'FrmTransaksiPOCOtorisasi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(884, 663)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 699)
        Me.Name = "FrmTransaksiPOCOtorisasi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Authorize Purchase Order Cancellation"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabBlmOtorisasi.ResumeLayout(False)
        Me.TabBlmOtorisasi.PerformLayout()
        Me.TabCancel.ResumeLayout(False)
        Me.TabCancel.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents TsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents TsbAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtCancellationReason As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents txtPONO As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtContactPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents llbAuthorization As System.Windows.Forms.LinkLabel
    Friend WithEvents txtAutohirizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryPOC As System.Windows.Forms.TextBox
    Friend WithEvents TxtPOCNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabBlmOtorisasi As System.Windows.Forms.TabPage
    Friend WithEvents llbUncheckAllAuthorize As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAllAuthorize As System.Windows.Forms.LinkLabel
    Friend WithEvents chkBelumOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TabCancel As System.Windows.Forms.TabPage
    Friend WithEvents llbUncheckAllCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAllCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents chkSudahOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
