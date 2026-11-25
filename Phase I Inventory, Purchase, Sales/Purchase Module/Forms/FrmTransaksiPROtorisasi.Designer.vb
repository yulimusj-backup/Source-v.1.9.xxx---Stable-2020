<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPROtorisasi
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPRitem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryCont = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrackUsing = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbExit = New System.Windows.Forms.ToolStripButton
        Me.tsbAuthorize = New System.Windows.Forms.ToolStripButton
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnCLose = New System.Windows.Forms.Button
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryReqBy = New System.Windows.Forms.TextBox
        Me.txtReqByCode = New System.Windows.Forms.TextBox
        Me.TxtReqByDesc = New System.Windows.Forms.TextBox
        Me.txtAutohirizationStatus = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.Button2 = New System.Windows.Forms.Button
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvOtorisasi = New System.Windows.Forms.DataGridView
        Me.ColDgvAuth_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvAuth_Authorize = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDgvAuth_PRNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvAuth_ReqBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbUncheckAll_Athorize = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_Authorize = New System.Windows.Forms.LinkLabel
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvCancelOtorisasi = New System.Windows.Forms.DataGridView
        Me.ColDgvUnAuth_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvUnAuth_Cancel = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDgvUnAuth_PRNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvUnAuth_ReqBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbUncheckAll_Cancel = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_Cancel = New System.Windows.Forms.LinkLabel
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.DgvAlreadyAuthorize = New System.Windows.Forms.DataGridView
        Me.ColAA_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAA_Print = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColAA_PRNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAA_ReqBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAA_Status = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtNoPR = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnFindPR = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.ChkPRDate = New System.Windows.Forms.CheckBox
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CboStatusPR = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryPR = New System.Windows.Forms.TextBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip
        Me.tsbRemove = New System.Windows.Forms.ToolStripButton
        Me.tsbShowRemoved = New System.Windows.Forms.ToolStripButton
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvOtorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvCancelOtorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DgvAlreadyAuthorize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.SuspendLayout()
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
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPrimaryPRitem, Me.ColPrimaryProduct, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnitCode, Me.ColDeliveryCont, Me.ColDeliveryDate, Me.ColRemark, Me.ColTrackUsing, Me.ColDiscFromPerct, Me.ColStatus, Me.ColIdentifier})
        Me.dgvItem.Location = New System.Drawing.Point(6, 104)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(697, 330)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 167
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
        'ColPrimaryPRitem
        '
        Me.ColPrimaryPRitem.HeaderText = "Primary PR Item"
        Me.ColPrimaryPRitem.Name = "ColPrimaryPRitem"
        Me.ColPrimaryPRitem.ReadOnly = True
        Me.ColPrimaryPRitem.Visible = False
        Me.ColPrimaryPRitem.Width = 107
        '
        'ColPrimaryProduct
        '
        Me.ColPrimaryProduct.HeaderText = "Primary Product"
        Me.ColPrimaryProduct.Name = "ColPrimaryProduct"
        Me.ColPrimaryProduct.ReadOnly = True
        Me.ColPrimaryProduct.Visible = False
        Me.ColPrimaryProduct.Width = 106
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
        'ColUnitCode
        '
        Me.ColUnitCode.HeaderText = "Units"
        Me.ColUnitCode.Name = "ColUnitCode"
        Me.ColUnitCode.ReadOnly = True
        Me.ColUnitCode.Width = 56
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
        'ColTrackUsing
        '
        Me.ColTrackUsing.HeaderText = "Track Using"
        Me.ColTrackUsing.Name = "ColTrackUsing"
        Me.ColTrackUsing.Width = 71
        '
        'ColDiscFromPerct
        '
        Me.ColDiscFromPerct.HeaderText = "OB"
        Me.ColDiscFromPerct.Name = "ColDiscFromPerct"
        Me.ColDiscFromPerct.Visible = False
        Me.ColDiscFromPerct.Width = 47
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColStatus.Width = 62
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 72
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(253, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.tsbAuthorize})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip1.Size = New System.Drawing.Size(514, 33)
        Me.ToolStrip1.TabIndex = 59
        '
        'tsbExit
        '
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(29, 30)
        Me.tsbExit.Text = "E&xit"
        Me.tsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbAuthorize
        '
        Me.tsbAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAuthorize.Name = "tsbAuthorize"
        Me.tsbAuthorize.Size = New System.Drawing.Size(62, 30)
        Me.tsbAuthorize.Text = "Authorize"
        Me.tsbAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(253, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(253, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 490)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(1028, 22)
        Me.StatusStrip2.TabIndex = 58
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(253, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(84, 4)
        Me.TxtRemark.MaxLength = 50
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.ReadOnly = True
        Me.TxtRemark.Size = New System.Drawing.Size(260, 72)
        Me.TxtRemark.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 7)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'btnCLose
        '
        Me.btnCLose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCLose.Location = New System.Drawing.Point(83, 275)
        Me.btnCLose.Name = "btnCLose"
        Me.btnCLose.Size = New System.Drawing.Size(122, 42)
        Me.btnCLose.TabIndex = 155
        Me.btnCLose.TabStop = False
        Me.btnCLose.Text = "Button1"
        Me.btnCLose.UseVisualStyleBackColor = True
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Request By"
        '
        'txtPrimaryReqBy
        '
        Me.txtPrimaryReqBy.Location = New System.Drawing.Point(76, 28)
        Me.txtPrimaryReqBy.Name = "txtPrimaryReqBy"
        Me.txtPrimaryReqBy.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryReqBy.TabIndex = 133
        Me.txtPrimaryReqBy.TabStop = False
        Me.txtPrimaryReqBy.Visible = False
        '
        'txtReqByCode
        '
        Me.txtReqByCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtReqByCode.Location = New System.Drawing.Point(105, 30)
        Me.txtReqByCode.MaxLength = 50
        Me.txtReqByCode.Name = "txtReqByCode"
        Me.txtReqByCode.ReadOnly = True
        Me.txtReqByCode.Size = New System.Drawing.Size(65, 20)
        Me.txtReqByCode.TabIndex = 2
        '
        'TxtReqByDesc
        '
        Me.TxtReqByDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReqByDesc.Location = New System.Drawing.Point(176, 30)
        Me.TxtReqByDesc.MaxLength = 150
        Me.TxtReqByDesc.Name = "TxtReqByDesc"
        Me.TxtReqByDesc.ReadOnly = True
        Me.TxtReqByDesc.Size = New System.Drawing.Size(167, 20)
        Me.TxtReqByDesc.TabIndex = 3
        Me.TxtReqByDesc.TabStop = False
        '
        'txtAutohirizationStatus
        '
        Me.txtAutohirizationStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAutohirizationStatus.Location = New System.Drawing.Point(107, 56)
        Me.txtAutohirizationStatus.MaxLength = 50
        Me.txtAutohirizationStatus.Name = "txtAutohirizationStatus"
        Me.txtAutohirizationStatus.ReadOnly = True
        Me.txtAutohirizationStatus.Size = New System.Drawing.Size(236, 20)
        Me.txtAutohirizationStatus.TabIndex = 4
        Me.txtAutohirizationStatus.TabStop = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(6, 59)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(101, 13)
        Me.Label21.TabIndex = 120
        Me.Label21.Text = "Authorization Status"
        '
        'gboDetail
        '
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCLose)
        Me.gboDetail.Controls.Add(Me.Button2)
        Me.gboDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboDetail.Location = New System.Drawing.Point(0, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(709, 441)
        Me.gboDetail.TabIndex = 57
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(7, 18)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtStockInNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryStockIn)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label21)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtAutohirizationStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtReqByDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryReqBy)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReqByCode)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer1.Size = New System.Drawing.Size(697, 81)
        Me.SplitContainer1.SplitterDistance = 346
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 168
        '
        'TxtStockInNo
        '
        Me.TxtStockInNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStockInNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStockInNo.Location = New System.Drawing.Point(107, 4)
        Me.TxtStockInNo.MaxLength = 50
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(120, 20)
        Me.TxtStockInNo.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(92, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Purchase Req No"
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(147, 4)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(233, 4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(111, 20)
        Me.dtpDate.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(431, 275)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 166
        Me.Button2.TabStop = False
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(9, 10)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.gboDetail)
        Me.SplitContainer2.Size = New System.Drawing.Size(1017, 441)
        Me.SplitContainer2.SplitterDistance = 305
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 169
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TabControl2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(305, 441)
        Me.SplitContainer3.SplitterDistance = 220
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 169
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(305, 220)
        Me.TabControl2.TabIndex = 169
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvOtorisasi)
        Me.TabPage2.Controls.Add(Me.LlbUncheckAll_Athorize)
        Me.TabPage2.Controls.Add(Me.LlbCheckAll_Authorize)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(297, 194)
        Me.TabPage2.TabIndex = 0
        Me.TabPage2.Text = "Authorize"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvOtorisasi
        '
        Me.DgvOtorisasi.AllowUserToAddRows = False
        Me.DgvOtorisasi.AllowUserToDeleteRows = False
        Me.DgvOtorisasi.AllowUserToResizeColumns = False
        Me.DgvOtorisasi.AllowUserToResizeRows = False
        Me.DgvOtorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvOtorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvOtorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvOtorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDgvAuth_Primary, Me.ColDgvAuth_Authorize, Me.ColDgvAuth_PRNo, Me.ColDgvAuth_ReqBy})
        Me.DgvOtorisasi.Location = New System.Drawing.Point(2, 2)
        Me.DgvOtorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvOtorisasi.MultiSelect = False
        Me.DgvOtorisasi.Name = "DgvOtorisasi"
        Me.DgvOtorisasi.RowHeadersVisible = False
        Me.DgvOtorisasi.RowTemplate.Height = 24
        Me.DgvOtorisasi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvOtorisasi.Size = New System.Drawing.Size(294, 175)
        Me.DgvOtorisasi.TabIndex = 170
        '
        'ColDgvAuth_Primary
        '
        Me.ColDgvAuth_Primary.HeaderText = "Primary"
        Me.ColDgvAuth_Primary.Name = "ColDgvAuth_Primary"
        Me.ColDgvAuth_Primary.ReadOnly = True
        Me.ColDgvAuth_Primary.Visible = False
        Me.ColDgvAuth_Primary.Width = 66
        '
        'ColDgvAuth_Authorize
        '
        Me.ColDgvAuth_Authorize.HeaderText = "Authorize"
        Me.ColDgvAuth_Authorize.Name = "ColDgvAuth_Authorize"
        Me.ColDgvAuth_Authorize.Width = 57
        '
        'ColDgvAuth_PRNo
        '
        Me.ColDgvAuth_PRNo.HeaderText = "PR No"
        Me.ColDgvAuth_PRNo.MinimumWidth = 64
        Me.ColDgvAuth_PRNo.Name = "ColDgvAuth_PRNo"
        Me.ColDgvAuth_PRNo.ReadOnly = True
        Me.ColDgvAuth_PRNo.Width = 64
        '
        'ColDgvAuth_ReqBy
        '
        Me.ColDgvAuth_ReqBy.HeaderText = "Request By"
        Me.ColDgvAuth_ReqBy.Name = "ColDgvAuth_ReqBy"
        Me.ColDgvAuth_ReqBy.ReadOnly = True
        Me.ColDgvAuth_ReqBy.Width = 87
        '
        'LlbUncheckAll_Athorize
        '
        Me.LlbUncheckAll_Athorize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_Athorize.AutoSize = True
        Me.LlbUncheckAll_Athorize.Location = New System.Drawing.Point(231, 180)
        Me.LlbUncheckAll_Athorize.Name = "LlbUncheckAll_Athorize"
        Me.LlbUncheckAll_Athorize.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_Athorize.TabIndex = 172
        Me.LlbUncheckAll_Athorize.TabStop = True
        Me.LlbUncheckAll_Athorize.Text = "Uncheck All"
        '
        'LlbCheckAll_Authorize
        '
        Me.LlbCheckAll_Authorize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_Authorize.AutoSize = True
        Me.LlbCheckAll_Authorize.Location = New System.Drawing.Point(176, 180)
        Me.LlbCheckAll_Authorize.Name = "LlbCheckAll_Authorize"
        Me.LlbCheckAll_Authorize.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_Authorize.TabIndex = 171
        Me.LlbCheckAll_Authorize.TabStop = True
        Me.LlbCheckAll_Authorize.Text = "Check All"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvCancelOtorisasi)
        Me.TabPage3.Controls.Add(Me.LlbUncheckAll_Cancel)
        Me.TabPage3.Controls.Add(Me.LlbCheckAll_Cancel)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(297, 194)
        Me.TabPage3.TabIndex = 1
        Me.TabPage3.Text = "Cancel Authorize"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DgvCancelOtorisasi
        '
        Me.DgvCancelOtorisasi.AllowUserToAddRows = False
        Me.DgvCancelOtorisasi.AllowUserToDeleteRows = False
        Me.DgvCancelOtorisasi.AllowUserToResizeColumns = False
        Me.DgvCancelOtorisasi.AllowUserToResizeRows = False
        Me.DgvCancelOtorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvCancelOtorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCancelOtorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCancelOtorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDgvUnAuth_Primary, Me.ColDgvUnAuth_Cancel, Me.ColDgvUnAuth_PRNo, Me.ColDgvUnAuth_ReqBy})
        Me.DgvCancelOtorisasi.Location = New System.Drawing.Point(2, 2)
        Me.DgvCancelOtorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvCancelOtorisasi.MultiSelect = False
        Me.DgvCancelOtorisasi.Name = "DgvCancelOtorisasi"
        Me.DgvCancelOtorisasi.RowHeadersVisible = False
        Me.DgvCancelOtorisasi.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCancelOtorisasi.RowTemplate.Height = 24
        Me.DgvCancelOtorisasi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvCancelOtorisasi.Size = New System.Drawing.Size(294, 175)
        Me.DgvCancelOtorisasi.TabIndex = 173
        '
        'ColDgvUnAuth_Primary
        '
        Me.ColDgvUnAuth_Primary.HeaderText = "Primary"
        Me.ColDgvUnAuth_Primary.Name = "ColDgvUnAuth_Primary"
        Me.ColDgvUnAuth_Primary.ReadOnly = True
        Me.ColDgvUnAuth_Primary.Visible = False
        Me.ColDgvUnAuth_Primary.Width = 66
        '
        'ColDgvUnAuth_Cancel
        '
        Me.ColDgvUnAuth_Cancel.HeaderText = "Cancel Authorize "
        Me.ColDgvUnAuth_Cancel.Name = "ColDgvUnAuth_Cancel"
        Me.ColDgvUnAuth_Cancel.Width = 96
        '
        'ColDgvUnAuth_PRNo
        '
        Me.ColDgvUnAuth_PRNo.HeaderText = "PR No"
        Me.ColDgvUnAuth_PRNo.Name = "ColDgvUnAuth_PRNo"
        Me.ColDgvUnAuth_PRNo.ReadOnly = True
        Me.ColDgvUnAuth_PRNo.Width = 64
        '
        'ColDgvUnAuth_ReqBy
        '
        Me.ColDgvUnAuth_ReqBy.HeaderText = "Request By"
        Me.ColDgvUnAuth_ReqBy.Name = "ColDgvUnAuth_ReqBy"
        Me.ColDgvUnAuth_ReqBy.ReadOnly = True
        Me.ColDgvUnAuth_ReqBy.Width = 87
        '
        'LlbUncheckAll_Cancel
        '
        Me.LlbUncheckAll_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_Cancel.AutoSize = True
        Me.LlbUncheckAll_Cancel.Location = New System.Drawing.Point(231, 180)
        Me.LlbUncheckAll_Cancel.Name = "LlbUncheckAll_Cancel"
        Me.LlbUncheckAll_Cancel.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_Cancel.TabIndex = 172
        Me.LlbUncheckAll_Cancel.TabStop = True
        Me.LlbUncheckAll_Cancel.Text = "Uncheck All"
        '
        'LlbCheckAll_Cancel
        '
        Me.LlbCheckAll_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_Cancel.AutoSize = True
        Me.LlbCheckAll_Cancel.Location = New System.Drawing.Point(176, 180)
        Me.LlbCheckAll_Cancel.Name = "LlbCheckAll_Cancel"
        Me.LlbCheckAll_Cancel.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_Cancel.TabIndex = 171
        Me.LlbCheckAll_Cancel.TabStop = True
        Me.LlbCheckAll_Cancel.Text = "Check All"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DgvAlreadyAuthorize)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(297, 194)
        Me.TabPage4.TabIndex = 2
        Me.TabPage4.Text = "Already Authorize"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DgvAlreadyAuthorize
        '
        Me.DgvAlreadyAuthorize.AllowUserToAddRows = False
        Me.DgvAlreadyAuthorize.AllowUserToDeleteRows = False
        Me.DgvAlreadyAuthorize.AllowUserToResizeColumns = False
        Me.DgvAlreadyAuthorize.AllowUserToResizeRows = False
        Me.DgvAlreadyAuthorize.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvAlreadyAuthorize.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvAlreadyAuthorize.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAA_Primary, Me.ColAA_Print, Me.ColAA_PRNo, Me.ColAA_ReqBy, Me.ColAA_Status})
        Me.DgvAlreadyAuthorize.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvAlreadyAuthorize.Location = New System.Drawing.Point(2, 2)
        Me.DgvAlreadyAuthorize.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvAlreadyAuthorize.MultiSelect = False
        Me.DgvAlreadyAuthorize.Name = "DgvAlreadyAuthorize"
        Me.DgvAlreadyAuthorize.RowHeadersVisible = False
        Me.DgvAlreadyAuthorize.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvAlreadyAuthorize.RowTemplate.Height = 24
        Me.DgvAlreadyAuthorize.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvAlreadyAuthorize.Size = New System.Drawing.Size(293, 190)
        Me.DgvAlreadyAuthorize.TabIndex = 172
        '
        'ColAA_Primary
        '
        Me.ColAA_Primary.HeaderText = "Primary"
        Me.ColAA_Primary.Name = "ColAA_Primary"
        Me.ColAA_Primary.ReadOnly = True
        Me.ColAA_Primary.Visible = False
        Me.ColAA_Primary.Width = 66
        '
        'ColAA_Print
        '
        Me.ColAA_Print.HeaderText = "Print"
        Me.ColAA_Print.Name = "ColAA_Print"
        Me.ColAA_Print.Width = 34
        '
        'ColAA_PRNo
        '
        Me.ColAA_PRNo.HeaderText = "PR No"
        Me.ColAA_PRNo.Name = "ColAA_PRNo"
        Me.ColAA_PRNo.ReadOnly = True
        Me.ColAA_PRNo.Width = 64
        '
        'ColAA_ReqBy
        '
        Me.ColAA_ReqBy.HeaderText = "Request By"
        Me.ColAA_ReqBy.Name = "ColAA_ReqBy"
        Me.ColAA_ReqBy.ReadOnly = True
        Me.ColAA_ReqBy.Width = 87
        '
        'ColAA_Status
        '
        Me.ColAA_Status.HeaderText = "Status"
        Me.ColAA_Status.Name = "ColAA_Status"
        Me.ColAA_Status.ReadOnly = True
        Me.ColAA_Status.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAA_Status.Width = 62
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnFilter)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.TxtNoPR)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.BtnFindPR)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.DgvUser)
        Me.GroupBox1.Controls.Add(Me.CboStatusPR)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryPR)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(305, 218)
        Me.GroupBox1.TabIndex = 169
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(255, 12)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(45, 23)
        Me.BtnFilter.TabIndex = 171
        Me.BtnFilter.TabStop = False
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(39, 13)
        Me.Label10.TabIndex = 174
        Me.Label10.Text = "PR No"
        '
        'TxtNoPR
        '
        Me.TxtNoPR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoPR.Location = New System.Drawing.Point(60, 38)
        Me.TxtNoPR.Name = "TxtNoPR"
        Me.TxtNoPR.Size = New System.Drawing.Size(190, 20)
        Me.TxtNoPR.TabIndex = 174
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(37, 13)
        Me.Label9.TabIndex = 172
        Me.Label9.Text = "Status"
        '
        'BtnFindPR
        '
        Me.BtnFindPR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindPR.Location = New System.Drawing.Point(255, 37)
        Me.BtnFindPR.Name = "BtnFindPR"
        Me.BtnFindPR.Size = New System.Drawing.Size(45, 22)
        Me.BtnFindPR.TabIndex = 171
        Me.BtnFindPR.TabStop = False
        Me.BtnFindPR.Text = "...."
        Me.BtnFindPR.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.DtpTo)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.DtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.ChkPRDate)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 59)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(296, 64)
        Me.GroupBox2.TabIndex = 175
        Me.GroupBox2.TabStop = False
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(56, 40)
        Me.DtpTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(235, 20)
        Me.DtpTo.TabIndex = 177
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 43)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 170
        Me.Label3.Text = "To"
        '
        'DtpFrom
        '
        Me.DtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(56, 17)
        Me.DtpFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(235, 20)
        Me.DtpFrom.TabIndex = 176
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 169
        Me.Label2.Text = "From"
        '
        'ChkPRDate
        '
        Me.ChkPRDate.AutoSize = True
        Me.ChkPRDate.Location = New System.Drawing.Point(5, 0)
        Me.ChkPRDate.Name = "ChkPRDate"
        Me.ChkPRDate.Size = New System.Drawing.Size(67, 17)
        Me.ChkPRDate.TabIndex = 175
        Me.ChkPRDate.Text = "PR Date"
        Me.ChkPRDate.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToAddRows = False
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_Primary, Me.ColCreatedBy, Me.ColCreatedBy_Name})
        Me.DgvUser.Location = New System.Drawing.Point(4, 127)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(296, 85)
        Me.DgvUser.TabIndex = 178
        '
        'ColCreatedBy_Primary
        '
        Me.ColCreatedBy_Primary.HeaderText = "Primary"
        Me.ColCreatedBy_Primary.Name = "ColCreatedBy_Primary"
        Me.ColCreatedBy_Primary.ReadOnly = True
        Me.ColCreatedBy_Primary.Visible = False
        Me.ColCreatedBy_Primary.Width = 66
        '
        'ColCreatedBy
        '
        Me.ColCreatedBy.HeaderText = "User Code"
        Me.ColCreatedBy.Name = "ColCreatedBy"
        Me.ColCreatedBy.Width = 82
        '
        'ColCreatedBy_Name
        '
        Me.ColCreatedBy_Name.HeaderText = "Name"
        Me.ColCreatedBy_Name.Name = "ColCreatedBy_Name"
        Me.ColCreatedBy_Name.ReadOnly = True
        Me.ColCreatedBy_Name.Width = 60
        '
        'CboStatusPR
        '
        Me.CboStatusPR.AllowDrop = True
        Me.CboStatusPR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboStatusPR.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatusPR.FormattingEnabled = True
        Me.CboStatusPR.Items.AddRange(New Object() {"All Purc. Request", "Authorized Only", "Unauthorized Only"})
        Me.CboStatusPR.Location = New System.Drawing.Point(60, 13)
        Me.CboStatusPR.Name = "CboStatusPR"
        Me.CboStatusPR.Size = New System.Drawing.Size(190, 21)
        Me.CboStatusPR.TabIndex = 173
        '
        'TxtPrimaryPR
        '
        Me.TxtPrimaryPR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryPR.Location = New System.Drawing.Point(26, 37)
        Me.TxtPrimaryPR.MaxLength = 50
        Me.TxtPrimaryPR.Name = "TxtPrimaryPR"
        Me.TxtPrimaryPR.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryPR.TabIndex = 167
        Me.TxtPrimaryPR.TabStop = False
        Me.TxtPrimaryPR.Visible = False
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 457)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.ToolStrip2)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.ToolStrip1)
        Me.SplitContainer4.Size = New System.Drawing.Size(1028, 33)
        Me.SplitContainer4.SplitterDistance = 511
        Me.SplitContainer4.SplitterWidth = 3
        Me.SplitContainer4.TabIndex = 170
        '
        'ToolStrip2
        '
        Me.ToolStrip2.AutoSize = False
        Me.ToolStrip2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip2.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip2.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbRemove, Me.tsbShowRemoved})
        Me.ToolStrip2.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(511, 33)
        Me.ToolStrip2.TabIndex = 60
        '
        'tsbRemove
        '
        Me.tsbRemove.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbRemove.Name = "tsbRemove"
        Me.tsbRemove.Size = New System.Drawing.Size(54, 30)
        Me.tsbRemove.Text = "Remove"
        Me.tsbRemove.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.tsbRemove.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.tsbRemove.Visible = False
        '
        'tsbShowRemoved
        '
        Me.tsbShowRemoved.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbShowRemoved.Name = "tsbShowRemoved"
        Me.tsbShowRemoved.Size = New System.Drawing.Size(93, 30)
        Me.tsbShowRemoved.Text = "Show Removed"
        Me.tsbShowRemoved.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.tsbShowRemoved.Visible = False
        '
        'BgwUser
        '
        '
        'FrmTransaksiPROtorisasi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCLose
        Me.ClientSize = New System.Drawing.Size(1028, 512)
        Me.Controls.Add(Me.SplitContainer4)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(1044, 550)
        Me.Name = "FrmTransaksiPROtorisasi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Authorize Purchase Request"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DgvOtorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.DgvCancelOtorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DgvAlreadyAuthorize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnCLose As System.Windows.Forms.Button
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryReqBy As System.Windows.Forms.TextBox
    Friend WithEvents txtReqByCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtReqByDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtAutohirizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvOtorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents LlbUncheckAll_Athorize As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll_Authorize As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvCancelOtorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents LlbUncheckAll_Cancel As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll_Cancel As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents DgvAlreadyAuthorize As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbRemove As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbShowRemoved As System.Windows.Forms.ToolStripButton
    Friend WithEvents CboStatusPR As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtNoPR As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPR As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryPR As System.Windows.Forms.TextBox
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkPRDate As System.Windows.Forms.CheckBox
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_Authorize As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDgvAuth_PRNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_ReqBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_Print As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColAA_PRNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_ReqBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPRitem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryCont As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrackUsing As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_Cancel As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDgvUnAuth_PRNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_ReqBy As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
