<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPOC
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.llbAuthorization = New System.Windows.Forms.LinkLabel
        Me.txtAutohirizationStatus = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryPOC = New System.Windows.Forms.TextBox
        Me.TxtPOCNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.cboCodeNo = New System.Windows.Forms.ComboBox
        Me.TxtRatioDefault = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtRatioReference = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.LlbQty = New System.Windows.Forms.LinkLabel
        Me.TxtQtyperBatchMulti = New System.Windows.Forms.TextBox
        Me.BtnInput = New System.Windows.Forms.Button
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.lblUnitsMulti = New System.Windows.Forms.Label
        Me.TxtQtyperBatchDefault = New System.Windows.Forms.TextBox
        Me.lblUnitsDefault = New System.Windows.Forms.Label
        Me.txtPrimaryUnitsDefault = New System.Windows.Forms.TextBox
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
        Me.ColMasterUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.btnFindPO = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtContactPerson = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtPrimaryPO = New System.Windows.Forms.TextBox
        Me.llbPODetail = New System.Windows.Forms.LinkLabel
        Me.txtPONO = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitsMulti = New System.Windows.Forms.TextBox
        Me.txtPhoneNo = New System.Windows.Forms.TextBox
        Me.txtCancellationReason = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnStopProcess = New System.Windows.Forms.Button
        Me.prgProgress = New System.Windows.Forms.ProgressBar
        Me.lblProgress = New System.Windows.Forms.Label
        Me.gboItem.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Supplier Code"
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSupplierName.Location = New System.Drawing.Point(105, 101)
        Me.TxtSupplierName.MaxLength = 150
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(305, 20)
        Me.TxtSupplierName.TabIndex = 6
        Me.TxtSupplierName.TabStop = False
        '
        'llbAuthorization
        '
        Me.llbAuthorization.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAuthorization.AutoSize = True
        Me.llbAuthorization.Location = New System.Drawing.Point(795, 48)
        Me.llbAuthorization.Name = "llbAuthorization"
        Me.llbAuthorization.Size = New System.Drawing.Size(57, 13)
        Me.llbAuthorization.TabIndex = 10
        Me.llbAuthorization.TabStop = True
        Me.llbAuthorization.Text = "view detail"
        Me.llbAuthorization.Visible = False
        '
        'txtAutohirizationStatus
        '
        Me.txtAutohirizationStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAutohirizationStatus.Location = New System.Drawing.Point(112, 35)
        Me.txtAutohirizationStatus.MaxLength = 50
        Me.txtAutohirizationStatus.Name = "txtAutohirizationStatus"
        Me.txtAutohirizationStatus.ReadOnly = True
        Me.txtAutohirizationStatus.Size = New System.Drawing.Size(243, 20)
        Me.txtAutohirizationStatus.TabIndex = 9
        Me.txtAutohirizationStatus.TabStop = False
        Me.txtAutohirizationStatus.Visible = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(8, 36)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(101, 13)
        Me.Label21.TabIndex = 120
        Me.Label21.Text = "Authorization Status"
        Me.Label21.Visible = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 503)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 22
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 503)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 23
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(105, 29)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(305, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtPrimaryPOC
        '
        Me.TxtPrimaryPOC.Location = New System.Drawing.Point(188, 7)
        Me.TxtPrimaryPOC.Name = "TxtPrimaryPOC"
        Me.TxtPrimaryPOC.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryPOC.TabIndex = 47
        Me.TxtPrimaryPOC.TabStop = False
        Me.TxtPrimaryPOC.Visible = False
        '
        'TxtPOCNo
        '
        Me.TxtPOCNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPOCNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPOCNo.Location = New System.Drawing.Point(105, 7)
        Me.TxtPOCNo.MaxLength = 50
        Me.TxtPOCNo.Name = "TxtPOCNo"
        Me.TxtPOCNo.ReadOnly = True
        Me.TxtPOCNo.Size = New System.Drawing.Size(305, 20)
        Me.TxtPOCNo.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 11)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "POC No"
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.SplitContainer2)
        Me.gboItem.Location = New System.Drawing.Point(6, 169)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(854, 70)
        Me.gboItem.TabIndex = 20
        Me.gboItem.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(1, 1)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtRatioDefault)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtRatioReference)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPrimaryCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDesc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.LlbQty)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyperBatchMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemarkItem)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblUnitsMulti)
        Me.SplitContainer2.Size = New System.Drawing.Size(852, 60)
        Me.SplitContainer2.SplitterDistance = 423
        Me.SplitContainer2.TabIndex = 162
        Me.SplitContainer2.TabStop = False
        '
        'cboCodeNo
        '
        Me.cboCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeNo.FormattingEnabled = True
        Me.cboCodeNo.Location = New System.Drawing.Point(105, 10)
        Me.cboCodeNo.Name = "cboCodeNo"
        Me.cboCodeNo.Size = New System.Drawing.Size(304, 21)
        Me.cboCodeNo.TabIndex = 0
        '
        'TxtRatioDefault
        '
        Me.TxtRatioDefault.Location = New System.Drawing.Point(352, 35)
        Me.TxtRatioDefault.Name = "TxtRatioDefault"
        Me.TxtRatioDefault.Size = New System.Drawing.Size(68, 20)
        Me.TxtRatioDefault.TabIndex = 192
        Me.TxtRatioDefault.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(9, 14)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 107
        Me.Label10.Text = "Code No"
        '
        'TxtRatioReference
        '
        Me.TxtRatioReference.Location = New System.Drawing.Point(271, 35)
        Me.TxtRatioReference.Name = "TxtRatioReference"
        Me.TxtRatioReference.Size = New System.Drawing.Size(68, 20)
        Me.TxtRatioReference.TabIndex = 191
        Me.TxtRatioReference.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(8, 39)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 109
        Me.Label11.Text = "Description"
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(105, 10)
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(38, 20)
        Me.txtPrimaryCodeNo.TabIndex = 112
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Visible = False
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(105, 35)
        Me.txtDesc.MaxLength = 150
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(304, 20)
        Me.txtDesc.TabIndex = 1
        Me.txtDesc.TabStop = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(18, 14)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 106
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'LlbQty
        '
        Me.LlbQty.ActiveLinkColor = System.Drawing.Color.Black
        Me.LlbQty.AutoSize = True
        Me.LlbQty.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.LlbQty.LinkColor = System.Drawing.Color.Black
        Me.LlbQty.Location = New System.Drawing.Point(9, 14)
        Me.LlbQty.Name = "LlbQty"
        Me.LlbQty.Size = New System.Drawing.Size(23, 13)
        Me.LlbQty.TabIndex = 269
        Me.LlbQty.TabStop = True
        Me.LlbQty.Text = "Qty"
        Me.LlbQty.VisitedLinkColor = System.Drawing.Color.Black
        '
        'TxtQtyperBatchMulti
        '
        Me.TxtQtyperBatchMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchMulti.Location = New System.Drawing.Point(112, 10)
        Me.TxtQtyperBatchMulti.MaxLength = 50
        Me.TxtQtyperBatchMulti.Name = "TxtQtyperBatchMulti"
        Me.TxtQtyperBatchMulti.Size = New System.Drawing.Size(257, 20)
        Me.TxtQtyperBatchMulti.TabIndex = 6
        Me.TxtQtyperBatchMulti.Text = "0"
        Me.TxtQtyperBatchMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(372, 35)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(48, 21)
        Me.BtnInput.TabIndex = 10
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(9, 39)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 13)
        Me.Label16.TabIndex = 63
        Me.Label16.Text = "Remarks"
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(112, 35)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(257, 20)
        Me.TxtRemarkItem.TabIndex = 9
        '
        'lblUnitsMulti
        '
        Me.lblUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsMulti.AutoSize = True
        Me.lblUnitsMulti.Location = New System.Drawing.Point(372, 14)
        Me.lblUnitsMulti.Name = "lblUnitsMulti"
        Me.lblUnitsMulti.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsMulti.TabIndex = 7
        Me.lblUnitsMulti.Text = "Units"
        '
        'TxtQtyperBatchDefault
        '
        Me.TxtQtyperBatchDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchDefault.Location = New System.Drawing.Point(711, 258)
        Me.TxtQtyperBatchDefault.Name = "TxtQtyperBatchDefault"
        Me.TxtQtyperBatchDefault.Size = New System.Drawing.Size(92, 20)
        Me.TxtQtyperBatchDefault.TabIndex = 272
        Me.TxtQtyperBatchDefault.Visible = False
        '
        'lblUnitsDefault
        '
        Me.lblUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsDefault.AutoSize = True
        Me.lblUnitsDefault.Location = New System.Drawing.Point(809, 262)
        Me.lblUnitsDefault.Name = "lblUnitsDefault"
        Me.lblUnitsDefault.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsDefault.TabIndex = 271
        Me.lblUnitsDefault.Text = "Units"
        Me.lblUnitsDefault.Visible = False
        '
        'txtPrimaryUnitsDefault
        '
        Me.txtPrimaryUnitsDefault.Location = New System.Drawing.Point(9, 101)
        Me.txtPrimaryUnitsDefault.Name = "txtPrimaryUnitsDefault"
        Me.txtPrimaryUnitsDefault.Size = New System.Drawing.Size(24, 20)
        Me.txtPrimaryUnitsDefault.TabIndex = 122
        Me.txtPrimaryUnitsDefault.TabStop = False
        Me.txtPrimaryUnitsDefault.Visible = False
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(218, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(439, 503)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 25
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(323, 503)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 24
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 570)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(890, 22)
        Me.StatusStrip1.TabIndex = 39
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(218, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(218, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(218, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.TxtQtyperBatchDefault)
        Me.gboDetail.Controls.Add(Me.lblUnitsDefault)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.llbAuthorization)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(8, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(866, 535)
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColMasterUnit, Me.ColQtyMulti, Me.ColUnitCodeMulti, Me.ColRatioDesc, Me.ColTotalQty, Me.ColUnit, Me.ColRemark, Me.ColRatioReference, Me.ColRatioDefault, Me.ColPrimaryCodeNo, Me.ColPrimaryMasterUnit, Me.ColPrimaryUnits, Me.ColPrimaryUnitMulti, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColQty})
        Me.dgvItem.Location = New System.Drawing.Point(6, 245)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(848, 252)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 21
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
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColMasterUnit
        '
        Me.ColMasterUnit.HeaderText = "Master Unit"
        Me.ColMasterUnit.Name = "ColMasterUnit"
        Me.ColMasterUnit.Visible = False
        Me.ColMasterUnit.Width = 86
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColUnitCodeMulti
        '
        Me.ColUnitCodeMulti.HeaderText = "Units Code"
        Me.ColUnitCodeMulti.Name = "ColUnitCodeMulti"
        Me.ColUnitCodeMulti.Width = 84
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRatioDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatioDesc.Width = 38
        '
        'ColTotalQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTotalQty.HeaderText = "Qty Default"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 93
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColRatioReference
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioReference.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 110
        '
        'ColRatioDefault
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioDefault.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 94
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
        'ColPrimaryUnitMulti
        '
        Me.ColPrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColPrimaryUnitMulti.Name = "ColPrimaryUnitMulti"
        Me.ColPrimaryUnitMulti.Visible = False
        Me.ColPrimaryUnitMulti.Width = 113
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
        'ColQty
        '
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Visible = False
        Me.ColQty.Width = 48
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 10)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtSupplierCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPOCNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimarySupplier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPOC)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindPO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSupplierName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtContactPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryPO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbPODetail)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPONO)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryUnitsDefault)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryUnitsMulti)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPhoneNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtCancellationReason)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtAutohirizationStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label23)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label21)
        Me.SplitContainer1.Size = New System.Drawing.Size(854, 148)
        Me.SplitContainer1.SplitterDistance = 424
        Me.SplitContainer1.TabIndex = 161
        Me.SplitContainer1.TabStop = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSupplierCode.Location = New System.Drawing.Point(105, 77)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.ReadOnly = True
        Me.txtSupplierCode.Size = New System.Drawing.Size(305, 20)
        Me.txtSupplierCode.TabIndex = 5
        Me.txtSupplierCode.TabStop = False
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(105, 77)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(50, 20)
        Me.txtPrimarySupplier.TabIndex = 160
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Visible = False
        '
        'btnFindPO
        '
        Me.btnFindPO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindPO.Location = New System.Drawing.Point(378, 53)
        Me.btnFindPO.Name = "btnFindPO"
        Me.btnFindPO.Size = New System.Drawing.Size(32, 20)
        Me.btnFindPO.TabIndex = 3
        Me.btnFindPO.TabStop = False
        Me.btnFindPO.Text = "...."
        Me.btnFindPO.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 57)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 13)
        Me.Label5.TabIndex = 156
        Me.Label5.Text = "Purchase Order No"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 105)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "Supplier Name"
        '
        'txtContactPerson
        '
        Me.txtContactPerson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtContactPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtContactPerson.Location = New System.Drawing.Point(105, 126)
        Me.txtContactPerson.MaxLength = 150
        Me.txtContactPerson.Name = "txtContactPerson"
        Me.txtContactPerson.ReadOnly = True
        Me.txtContactPerson.Size = New System.Drawing.Size(305, 20)
        Me.txtContactPerson.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 130)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 138
        Me.Label7.Text = "Contact Person"
        '
        'txtPrimaryPO
        '
        Me.txtPrimaryPO.Location = New System.Drawing.Point(202, 53)
        Me.txtPrimaryPO.Name = "txtPrimaryPO"
        Me.txtPrimaryPO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryPO.TabIndex = 157
        Me.txtPrimaryPO.TabStop = False
        Me.txtPrimaryPO.Visible = False
        '
        'llbPODetail
        '
        Me.llbPODetail.AutoSize = True
        Me.llbPODetail.Location = New System.Drawing.Point(276, 60)
        Me.llbPODetail.Name = "llbPODetail"
        Me.llbPODetail.Size = New System.Drawing.Size(57, 13)
        Me.llbPODetail.TabIndex = 4
        Me.llbPODetail.TabStop = True
        Me.llbPODetail.Text = "view detail"
        Me.llbPODetail.Visible = False
        '
        'txtPONO
        '
        Me.txtPONO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPONO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPONO.Location = New System.Drawing.Point(105, 53)
        Me.txtPONO.MaxLength = 50
        Me.txtPONO.Name = "txtPONO"
        Me.txtPONO.Size = New System.Drawing.Size(267, 20)
        Me.txtPONO.TabIndex = 2
        '
        'txtPrimaryUnitsMulti
        '
        Me.txtPrimaryUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryUnitsMulti.Location = New System.Drawing.Point(9, 121)
        Me.txtPrimaryUnitsMulti.Name = "txtPrimaryUnitsMulti"
        Me.txtPrimaryUnitsMulti.Size = New System.Drawing.Size(24, 20)
        Me.txtPrimaryUnitsMulti.TabIndex = 192
        Me.txtPrimaryUnitsMulti.TabStop = False
        Me.txtPrimaryUnitsMulti.Visible = False
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPhoneNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPhoneNo.Location = New System.Drawing.Point(112, 8)
        Me.txtPhoneNo.MaxLength = 50
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.ReadOnly = True
        Me.txtPhoneNo.Size = New System.Drawing.Size(305, 20)
        Me.txtPhoneNo.TabIndex = 8
        '
        'txtCancellationReason
        '
        Me.txtCancellationReason.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCancellationReason.Location = New System.Drawing.Point(112, 30)
        Me.txtCancellationReason.MaxLength = 150
        Me.txtCancellationReason.Multiline = True
        Me.txtCancellationReason.Name = "txtCancellationReason"
        Me.txtCancellationReason.Size = New System.Drawing.Size(305, 43)
        Me.txtCancellationReason.TabIndex = 18
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 36)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(105, 13)
        Me.Label8.TabIndex = 159
        Me.Label8.Text = "Cancellation Reason"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(8, 12)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(55, 13)
        Me.Label23.TabIndex = 141
        Me.Label23.Text = "Phone No"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(112, 78)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(305, 69)
        Me.TxtRemark.TabIndex = 19
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 82)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(297, 541)
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
        Me.btnExit.Location = New System.Drawing.Point(585, 541)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(201, 541)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(489, 541)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(393, 541)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnStopProcess
        '
        Me.btnStopProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnStopProcess.Location = New System.Drawing.Point(9, 541)
        Me.btnStopProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnStopProcess.Name = "btnStopProcess"
        Me.btnStopProcess.Size = New System.Drawing.Size(80, 26)
        Me.btnStopProcess.TabIndex = 40
        Me.btnStopProcess.TabStop = False
        Me.btnStopProcess.Text = "Stop Process"
        Me.btnStopProcess.UseVisualStyleBackColor = True
        Me.btnStopProcess.Visible = False
        '
        'prgProgress
        '
        Me.prgProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.prgProgress.Location = New System.Drawing.Point(94, 545)
        Me.prgProgress.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.prgProgress.Name = "prgProgress"
        Me.prgProgress.Size = New System.Drawing.Size(190, 19)
        Me.prgProgress.TabIndex = 41
        Me.prgProgress.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(290, 547)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 13)
        Me.lblProgress.TabIndex = 160
        Me.lblProgress.Visible = False
        '
        'FrmTransaksiPOC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(890, 592)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.prgProgress)
        Me.Controls.Add(Me.btnStopProcess)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.MinimumSize = New System.Drawing.Size(906, 630)
        Me.Name = "FrmTransaksiPOC"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Order Cancellation"
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents llbAuthorization As System.Windows.Forms.LinkLabel
    Friend WithEvents txtAutohirizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryPOC As System.Windows.Forms.TextBox
    Friend WithEvents TxtPOCNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
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
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFindPO As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents txtCancellationReason As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents cboCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryUnitsDefault As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsMulti As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchMulti As System.Windows.Forms.TextBox
    Friend WithEvents llbPODetail As System.Windows.Forms.LinkLabel
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtRatioDefault As System.Windows.Forms.TextBox
    Friend WithEvents TxtRatioReference As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitsMulti As System.Windows.Forms.TextBox
    Friend WithEvents LlbQty As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtQtyperBatchDefault As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsDefault As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnStopProcess As System.Windows.Forms.Button
    Friend WithEvents prgProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
End Class
