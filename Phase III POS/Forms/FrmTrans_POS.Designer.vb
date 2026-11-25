<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Button8 = New System.Windows.Forms.Button
        Me.spc2 = New System.Windows.Forms.SplitContainer
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlled = New System.Windows.Forms.CheckBox
        Me.chkBatchNoControlled = New System.Windows.Forms.CheckBox
        Me.lblLocation = New System.Windows.Forms.Label
        Me.txtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.DtpExpired = New System.Windows.Forms.DateTimePicker
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.LblCodeNo = New System.Windows.Forms.Label
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtDiscString = New System.Windows.Forms.TextBox
        Me.TxtNettPrice = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtDiscAmount = New System.Windows.Forms.TextBox
        Me.LblDiscount = New System.Windows.Forms.Label
        Me.TxtUnitPrice = New System.Windows.Forms.TextBox
        Me.LblUnitPrice = New System.Windows.Forms.Label
        Me.spc1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimarySession = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtIdNumber = New System.Windows.Forms.TextBox
        Me.dtpTanggalTransaksi = New System.Windows.Forms.DateTimePicker
        Me.btnFindMember = New System.Windows.Forms.Button
        Me.TxtPrimaryMember = New System.Windows.Forms.TextBox
        Me.TxtMember = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDisc_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSubTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAutoNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PbxViod = New System.Windows.Forms.PictureBox
        Me.PbxPaid = New System.Windows.Forms.PictureBox
        Me.gboButton = New System.Windows.Forms.GroupBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnReprint = New System.Windows.Forms.Button
        Me.btnVoid = New System.Windows.Forms.Button
        Me.btnPayOutDropCash = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnPayment = New System.Windows.Forms.Button
        Me.TxtDisplay = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.gboDetail.SuspendLayout()
        Me.spc2.Panel1.SuspendLayout()
        Me.spc2.Panel2.SuspendLayout()
        Me.spc2.SuspendLayout()
        Me.spc1.Panel1.SuspendLayout()
        Me.spc1.Panel2.SuspendLayout()
        Me.spc1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PbxViod, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PbxPaid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboButton.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Button8)
        Me.gboDetail.Controls.Add(Me.spc2)
        Me.gboDetail.Controls.Add(Me.spc1)
        Me.gboDetail.Controls.Add(Me.Button1)
        Me.gboDetail.Controls.Add(Me.Button3)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Location = New System.Drawing.Point(12, -1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(851, 408)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'Button8
        '
        Me.Button8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button8.Location = New System.Drawing.Point(206, 346)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(91, 56)
        Me.Button8.TabIndex = 6
        Me.Button8.TabStop = False
        Me.Button8.Text = "Next"
        Me.Button8.UseVisualStyleBackColor = True
        Me.Button8.Visible = False
        '
        'spc2
        '
        Me.spc2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.spc2.Location = New System.Drawing.Point(6, 40)
        Me.spc2.Name = "spc2"
        '
        'spc2.Panel1
        '
        Me.spc2.Panel1.Controls.Add(Me.chkLocationControlled)
        Me.spc2.Panel1.Controls.Add(Me.chkExpiryDateControlled)
        Me.spc2.Panel1.Controls.Add(Me.chkBatchNoControlled)
        Me.spc2.Panel1.Controls.Add(Me.lblLocation)
        Me.spc2.Panel1.Controls.Add(Me.txtPrimaryLocation)
        Me.spc2.Panel1.Controls.Add(Me.Label32)
        Me.spc2.Panel1.Controls.Add(Me.lblBatchNo)
        Me.spc2.Panel1.Controls.Add(Me.lblExpiryDate)
        Me.spc2.Panel1.Controls.Add(Me.Label14)
        Me.spc2.Panel1.Controls.Add(Me.DtpExpired)
        Me.spc2.Panel1.Controls.Add(Me.Label16)
        Me.spc2.Panel1.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.spc2.Panel1.Controls.Add(Me.TxtQty)
        Me.spc2.Panel1.Controls.Add(Me.TxtCodeNo)
        Me.spc2.Panel1.Controls.Add(Me.btnFindCodeNo)
        Me.spc2.Panel1.Controls.Add(Me.Label8)
        Me.spc2.Panel1.Controls.Add(Me.LblCodeNo)
        Me.spc2.Panel1.Controls.Add(Me.cboLocation)
        Me.spc2.Panel1.Controls.Add(Me.cboBatchNo)
        '
        'spc2.Panel2
        '
        Me.spc2.Panel2.Controls.Add(Me.TxtReadOnly)
        Me.spc2.Panel2.Controls.Add(Me.TxtDiscString)
        Me.spc2.Panel2.Controls.Add(Me.TxtNettPrice)
        Me.spc2.Panel2.Controls.Add(Me.Label9)
        Me.spc2.Panel2.Controls.Add(Me.TxtDiscAmount)
        Me.spc2.Panel2.Controls.Add(Me.LblDiscount)
        Me.spc2.Panel2.Controls.Add(Me.TxtUnitPrice)
        Me.spc2.Panel2.Controls.Add(Me.LblUnitPrice)
        Me.spc2.Size = New System.Drawing.Size(839, 139)
        Me.spc2.SplitterDistance = 417
        Me.spc2.TabIndex = 1
        Me.spc2.TabStop = False
        '
        'chkLocationControlled
        '
        Me.chkLocationControlled.AutoSize = True
        Me.chkLocationControlled.Location = New System.Drawing.Point(251, 114)
        Me.chkLocationControlled.Name = "chkLocationControlled"
        Me.chkLocationControlled.Size = New System.Drawing.Size(67, 17)
        Me.chkLocationControlled.TabIndex = 138
        Me.chkLocationControlled.TabStop = False
        Me.chkLocationControlled.Text = "Location"
        Me.chkLocationControlled.UseVisualStyleBackColor = True
        Me.chkLocationControlled.Visible = False
        '
        'chkExpiryDateControlled
        '
        Me.chkExpiryDateControlled.AutoSize = True
        Me.chkExpiryDateControlled.Location = New System.Drawing.Point(151, 87)
        Me.chkExpiryDateControlled.Name = "chkExpiryDateControlled"
        Me.chkExpiryDateControlled.Size = New System.Drawing.Size(80, 17)
        Me.chkExpiryDateControlled.TabIndex = 137
        Me.chkExpiryDateControlled.TabStop = False
        Me.chkExpiryDateControlled.Text = "Expiry Date"
        Me.chkExpiryDateControlled.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlled.Visible = False
        '
        'chkBatchNoControlled
        '
        Me.chkBatchNoControlled.AutoSize = True
        Me.chkBatchNoControlled.Location = New System.Drawing.Point(251, 57)
        Me.chkBatchNoControlled.Name = "chkBatchNoControlled"
        Me.chkBatchNoControlled.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlled.TabIndex = 136
        Me.chkBatchNoControlled.TabStop = False
        Me.chkBatchNoControlled.Text = "Batch No"
        Me.chkBatchNoControlled.UseVisualStyleBackColor = True
        Me.chkBatchNoControlled.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = True
        Me.lblLocation.Location = New System.Drawing.Point(114, 113)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(24, 13)
        Me.lblLocation.TabIndex = 135
        Me.lblLocation.Text = "n/a"
        '
        'txtPrimaryLocation
        '
        Me.txtPrimaryLocation.Location = New System.Drawing.Point(83, 110)
        Me.txtPrimaryLocation.Name = "txtPrimaryLocation"
        Me.txtPrimaryLocation.Size = New System.Drawing.Size(30, 20)
        Me.txtPrimaryLocation.TabIndex = 134
        Me.txtPrimaryLocation.TabStop = False
        Me.txtPrimaryLocation.Text = "-1"
        Me.txtPrimaryLocation.Visible = False
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(3, 113)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(48, 13)
        Me.Label32.TabIndex = 133
        Me.Label32.Text = "Location"
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Location = New System.Drawing.Point(114, 60)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNo.TabIndex = 132
        Me.lblBatchNo.Text = "n/a"
        '
        'lblExpiryDate
        '
        Me.lblExpiryDate.AutoSize = True
        Me.lblExpiryDate.Location = New System.Drawing.Point(114, 86)
        Me.lblExpiryDate.Name = "lblExpiryDate"
        Me.lblExpiryDate.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDate.TabIndex = 131
        Me.lblExpiryDate.Text = "n/a"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(3, 60)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(52, 13)
        Me.Label14.TabIndex = 130
        Me.Label14.Text = "Batch No"
        '
        'DtpExpired
        '
        Me.DtpExpired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpExpired.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpired.Enabled = False
        Me.DtpExpired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpired.Location = New System.Drawing.Point(114, 83)
        Me.DtpExpired.Name = "DtpExpired"
        Me.DtpExpired.Size = New System.Drawing.Size(300, 20)
        Me.DtpExpired.TabIndex = 127
        Me.DtpExpired.TabStop = False
        Me.DtpExpired.Visible = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(3, 86)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(61, 13)
        Me.Label16.TabIndex = 129
        Me.Label16.Text = "Expiry Date"
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(80, 29)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 14
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Text = "-1"
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'TxtQty
        '
        Me.TxtQty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQty.Location = New System.Drawing.Point(114, 3)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(300, 20)
        Me.TxtQty.TabIndex = 0
        Me.TxtQty.Text = "1"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(114, 29)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(237, 20)
        Me.TxtCodeNo.TabIndex = 1
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCodeNo.Location = New System.Drawing.Point(357, 29)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(57, 20)
        Me.btnFindCodeNo.TabIndex = 2
        Me.btnFindCodeNo.TabStop = False
        Me.btnFindCodeNo.Text = "Find"
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 6)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(23, 13)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Qty"
        '
        'LblCodeNo
        '
        Me.LblCodeNo.AutoSize = True
        Me.LblCodeNo.Location = New System.Drawing.Point(3, 32)
        Me.LblCodeNo.Name = "LblCodeNo"
        Me.LblCodeNo.Size = New System.Drawing.Size(49, 13)
        Me.LblCodeNo.TabIndex = 0
        Me.LblCodeNo.Text = "Code No"
        '
        'cboLocation
        '
        Me.cboLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.FormattingEnabled = True
        Me.cboLocation.Location = New System.Drawing.Point(114, 109)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Size = New System.Drawing.Size(300, 21)
        Me.cboLocation.TabIndex = 128
        Me.cboLocation.TabStop = False
        Me.cboLocation.Visible = False
        '
        'cboBatchNo
        '
        Me.cboBatchNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(114, 55)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(300, 21)
        Me.cboBatchNo.TabIndex = 126
        Me.cboBatchNo.TabStop = False
        Me.cboBatchNo.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(171, 59)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(77, 20)
        Me.TxtReadOnly.TabIndex = 9
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtDiscString
        '
        Me.TxtDiscString.Location = New System.Drawing.Point(115, 29)
        Me.TxtDiscString.Name = "TxtDiscString"
        Me.TxtDiscString.Size = New System.Drawing.Size(89, 20)
        Me.TxtDiscString.TabIndex = 1
        Me.TxtDiscString.TabStop = False
        Me.TxtDiscString.Text = "0"
        Me.TxtDiscString.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtNettPrice
        '
        Me.TxtNettPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNettPrice.Location = New System.Drawing.Point(115, 55)
        Me.TxtNettPrice.Name = "TxtNettPrice"
        Me.TxtNettPrice.ReadOnly = True
        Me.TxtNettPrice.Size = New System.Drawing.Size(300, 20)
        Me.TxtNettPrice.TabIndex = 3
        Me.TxtNettPrice.TabStop = False
        Me.TxtNettPrice.Text = "0"
        Me.TxtNettPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 58)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(54, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Nett Price"
        '
        'TxtDiscAmount
        '
        Me.TxtDiscAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDiscAmount.Location = New System.Drawing.Point(210, 29)
        Me.TxtDiscAmount.Name = "TxtDiscAmount"
        Me.TxtDiscAmount.ReadOnly = True
        Me.TxtDiscAmount.Size = New System.Drawing.Size(205, 20)
        Me.TxtDiscAmount.TabIndex = 2
        Me.TxtDiscAmount.TabStop = False
        Me.TxtDiscAmount.Text = "0"
        Me.TxtDiscAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblDiscount
        '
        Me.LblDiscount.AutoSize = True
        Me.LblDiscount.Location = New System.Drawing.Point(4, 32)
        Me.LblDiscount.Name = "LblDiscount"
        Me.LblDiscount.Size = New System.Drawing.Size(49, 13)
        Me.LblDiscount.TabIndex = 6
        Me.LblDiscount.Text = "Discount"
        '
        'TxtUnitPrice
        '
        Me.TxtUnitPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnitPrice.Location = New System.Drawing.Point(115, 3)
        Me.TxtUnitPrice.Name = "TxtUnitPrice"
        Me.TxtUnitPrice.Size = New System.Drawing.Size(300, 20)
        Me.TxtUnitPrice.TabIndex = 0
        Me.TxtUnitPrice.TabStop = False
        Me.TxtUnitPrice.Text = "0"
        Me.TxtUnitPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblUnitPrice
        '
        Me.LblUnitPrice.AutoSize = True
        Me.LblUnitPrice.Location = New System.Drawing.Point(4, 6)
        Me.LblUnitPrice.Name = "LblUnitPrice"
        Me.LblUnitPrice.Size = New System.Drawing.Size(53, 13)
        Me.LblUnitPrice.TabIndex = 4
        Me.LblUnitPrice.Text = "Unit Price"
        '
        'spc1
        '
        Me.spc1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.spc1.Location = New System.Drawing.Point(6, 13)
        Me.spc1.Name = "spc1"
        '
        'spc1.Panel1
        '
        Me.spc1.Panel1.Controls.Add(Me.TxtPrimarySession)
        Me.spc1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.spc1.Panel1.Controls.Add(Me.Label1)
        Me.spc1.Panel1.Controls.Add(Me.TxtIdNumber)
        Me.spc1.Panel1.Controls.Add(Me.dtpTanggalTransaksi)
        '
        'spc1.Panel2
        '
        Me.spc1.Panel2.Controls.Add(Me.btnFindMember)
        Me.spc1.Panel2.Controls.Add(Me.TxtPrimaryMember)
        Me.spc1.Panel2.Controls.Add(Me.TxtMember)
        Me.spc1.Panel2.Controls.Add(Me.Label2)
        Me.spc1.Size = New System.Drawing.Size(839, 28)
        Me.spc1.SplitterDistance = 417
        Me.spc1.TabIndex = 0
        Me.spc1.TabStop = False
        '
        'TxtPrimarySession
        '
        Me.TxtPrimarySession.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimarySession.Location = New System.Drawing.Point(46, 3)
        Me.TxtPrimarySession.Name = "TxtPrimarySession"
        Me.TxtPrimarySession.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimarySession.TabIndex = 16
        Me.TxtPrimarySession.TabStop = False
        Me.TxtPrimarySession.Text = "-1"
        Me.TxtPrimarySession.Visible = False
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimary.Location = New System.Drawing.Point(80, 3)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimary.TabIndex = 15
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ID Number"
        '
        'TxtIdNumber
        '
        Me.TxtIdNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtIdNumber.Location = New System.Drawing.Point(114, 3)
        Me.TxtIdNumber.Name = "TxtIdNumber"
        Me.TxtIdNumber.ReadOnly = True
        Me.TxtIdNumber.Size = New System.Drawing.Size(300, 20)
        Me.TxtIdNumber.TabIndex = 1
        Me.TxtIdNumber.TabStop = False
        '
        'dtpTanggalTransaksi
        '
        Me.dtpTanggalTransaksi.CustomFormat = "dd MMM yyyy"
        Me.dtpTanggalTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTanggalTransaksi.Location = New System.Drawing.Point(117, 4)
        Me.dtpTanggalTransaksi.Name = "dtpTanggalTransaksi"
        Me.dtpTanggalTransaksi.Size = New System.Drawing.Size(297, 20)
        Me.dtpTanggalTransaksi.TabIndex = 17
        Me.dtpTanggalTransaksi.Visible = False
        '
        'btnFindMember
        '
        Me.btnFindMember.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindMember.Location = New System.Drawing.Point(358, 2)
        Me.btnFindMember.Name = "btnFindMember"
        Me.btnFindMember.Size = New System.Drawing.Size(57, 20)
        Me.btnFindMember.TabIndex = 15
        Me.btnFindMember.TabStop = False
        Me.btnFindMember.Text = "Find"
        Me.btnFindMember.UseVisualStyleBackColor = True
        '
        'TxtPrimaryMember
        '
        Me.TxtPrimaryMember.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryMember.Location = New System.Drawing.Point(81, 3)
        Me.TxtPrimaryMember.Name = "TxtPrimaryMember"
        Me.TxtPrimaryMember.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryMember.TabIndex = 16
        Me.TxtPrimaryMember.TabStop = False
        Me.TxtPrimaryMember.Text = "-1"
        Me.TxtPrimaryMember.Visible = False
        '
        'TxtMember
        '
        Me.TxtMember.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMember.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMember.Location = New System.Drawing.Point(115, 4)
        Me.TxtMember.Name = "TxtMember"
        Me.TxtMember.Size = New System.Drawing.Size(237, 20)
        Me.TxtMember.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Member"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(12, 346)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(91, 56)
        Me.Button1.TabIndex = 3
        Me.Button1.TabStop = False
        Me.Button1.Text = "Next"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(109, 346)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(91, 56)
        Me.Button3.TabIndex = 4
        Me.Button3.TabStop = False
        Me.Button3.Text = "Next"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColLocation, Me.ColBatchNumber, Me.ColExpiry, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColDiscString, Me.ColDisc_Amount, Me.ColNettPrice, Me.ColSubTotal, Me.ColPrimaryCodeNo, Me.ColPrimaryUnit, Me.ColPrimaryDisc1, Me.ColDiscAmount1, Me.ColPrimaryDisc2, Me.ColDiscAmount2, Me.ColPrimaryDisc3, Me.ColDiscAmount3, Me.ColPrimaryLocation, Me.ColDiscString1, Me.ColDiscString2, Me.ColDiscString3, Me.ColPrimaryAutoNumber})
        Me.dgvItem.Location = New System.Drawing.Point(6, 185)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(839, 217)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 2
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 55
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.ReadOnly = True
        Me.ColLocation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocation.Visible = False
        Me.ColLocation.Width = 54
        '
        'ColBatchNumber
        '
        Me.ColBatchNumber.HeaderText = "Lot No"
        Me.ColBatchNumber.Name = "ColBatchNumber"
        Me.ColBatchNumber.ReadOnly = True
        Me.ColBatchNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBatchNumber.Visible = False
        Me.ColBatchNumber.Width = 45
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        Me.ColExpiry.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExpiry.Visible = False
        Me.ColExpiry.Width = 41
        '
        'ColQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 29
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 32
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 59
        '
        'ColDiscString
        '
        Me.ColDiscString.HeaderText = "Discount String"
        Me.ColDiscString.Name = "ColDiscString"
        Me.ColDiscString.ReadOnly = True
        Me.ColDiscString.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscString.Visible = False
        Me.ColDiscString.Width = 85
        '
        'ColDisc_Amount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDisc_Amount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDisc_Amount.HeaderText = "Discount"
        Me.ColDisc_Amount.Name = "ColDisc_Amount"
        Me.ColDisc_Amount.ReadOnly = True
        Me.ColDisc_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDisc_Amount.Width = 55
        '
        'ColNettPrice
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNettPrice.Width = 60
        '
        'ColSubTotal
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSubTotal.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColSubTotal.HeaderText = "Sub Total"
        Me.ColSubTotal.Name = "ColSubTotal"
        Me.ColSubTotal.ReadOnly = True
        Me.ColSubTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSubTotal.Width = 59
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 111
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 88
        '
        'ColPrimaryDisc1
        '
        Me.ColPrimaryDisc1.HeaderText = "PrimaryDisc1"
        Me.ColPrimaryDisc1.Name = "ColPrimaryDisc1"
        Me.ColPrimaryDisc1.ReadOnly = True
        Me.ColPrimaryDisc1.Visible = False
        Me.ColPrimaryDisc1.Width = 93
        '
        'ColDiscAmount1
        '
        Me.ColDiscAmount1.HeaderText = "DiscAmount1"
        Me.ColDiscAmount1.Name = "ColDiscAmount1"
        Me.ColDiscAmount1.ReadOnly = True
        Me.ColDiscAmount1.Visible = False
        Me.ColDiscAmount1.Width = 95
        '
        'ColPrimaryDisc2
        '
        Me.ColPrimaryDisc2.HeaderText = "Primary Disc 2"
        Me.ColPrimaryDisc2.Name = "ColPrimaryDisc2"
        Me.ColPrimaryDisc2.ReadOnly = True
        Me.ColPrimaryDisc2.Visible = False
        Me.ColPrimaryDisc2.Width = 99
        '
        'ColDiscAmount2
        '
        Me.ColDiscAmount2.HeaderText = "Disc Amount2"
        Me.ColDiscAmount2.Name = "ColDiscAmount2"
        Me.ColDiscAmount2.ReadOnly = True
        Me.ColDiscAmount2.Visible = False
        Me.ColDiscAmount2.Width = 98
        '
        'ColPrimaryDisc3
        '
        Me.ColPrimaryDisc3.HeaderText = "Primary Disc 3"
        Me.ColPrimaryDisc3.Name = "ColPrimaryDisc3"
        Me.ColPrimaryDisc3.ReadOnly = True
        Me.ColPrimaryDisc3.Visible = False
        Me.ColPrimaryDisc3.Width = 99
        '
        'ColDiscAmount3
        '
        Me.ColDiscAmount3.HeaderText = "Disc Amount 3"
        Me.ColDiscAmount3.Name = "ColDiscAmount3"
        Me.ColDiscAmount3.ReadOnly = True
        Me.ColDiscAmount3.Visible = False
        Me.ColDiscAmount3.Width = 101
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.ReadOnly = True
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 110
        '
        'ColDiscString1
        '
        Me.ColDiscString1.HeaderText = "DiscString1"
        Me.ColDiscString1.Name = "ColDiscString1"
        Me.ColDiscString1.ReadOnly = True
        Me.ColDiscString1.Visible = False
        Me.ColDiscString1.Width = 86
        '
        'ColDiscString2
        '
        Me.ColDiscString2.HeaderText = "DiscString2"
        Me.ColDiscString2.Name = "ColDiscString2"
        Me.ColDiscString2.ReadOnly = True
        Me.ColDiscString2.Visible = False
        Me.ColDiscString2.Width = 86
        '
        'ColDiscString3
        '
        Me.ColDiscString3.HeaderText = "DiscString3"
        Me.ColDiscString3.Name = "ColDiscString3"
        Me.ColDiscString3.ReadOnly = True
        Me.ColDiscString3.Visible = False
        Me.ColDiscString3.Width = 86
        '
        'ColPrimaryAutoNumber
        '
        Me.ColPrimaryAutoNumber.HeaderText = "PrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.Name = "ColPrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.ReadOnly = True
        Me.ColPrimaryAutoNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAutoNumber.Visible = False
        Me.ColPrimaryAutoNumber.Width = 106
        '
        'PbxViod
        '
        Me.PbxViod.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PbxViod.BackgroundImage = Global.SMESS_POS.My.Resources.Resources.Paid
        Me.PbxViod.Image = Global.SMESS_POS.My.Resources.Resources.Void_stamp
        Me.PbxViod.Location = New System.Drawing.Point(716, 414)
        Me.PbxViod.Name = "PbxViod"
        Me.PbxViod.Size = New System.Drawing.Size(147, 122)
        Me.PbxViod.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PbxViod.TabIndex = 8
        Me.PbxViod.TabStop = False
        Me.PbxViod.Visible = False
        '
        'PbxPaid
        '
        Me.PbxPaid.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PbxPaid.BackgroundImage = Global.SMESS_POS.My.Resources.Resources.Paid
        Me.PbxPaid.Image = Global.SMESS_POS.My.Resources.Resources.Paid
        Me.PbxPaid.Location = New System.Drawing.Point(716, 414)
        Me.PbxPaid.Name = "PbxPaid"
        Me.PbxPaid.Size = New System.Drawing.Size(147, 123)
        Me.PbxPaid.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PbxPaid.TabIndex = 7
        Me.PbxPaid.TabStop = False
        Me.PbxPaid.Visible = False
        '
        'gboButton
        '
        Me.gboButton.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboButton.Controls.Add(Me.btnClose)
        Me.gboButton.Controls.Add(Me.btnReprint)
        Me.gboButton.Controls.Add(Me.btnVoid)
        Me.gboButton.Controls.Add(Me.btnPayOutDropCash)
        Me.gboButton.Controls.Add(Me.btnNext)
        Me.gboButton.Controls.Add(Me.btnNew)
        Me.gboButton.Controls.Add(Me.btnFind)
        Me.gboButton.Controls.Add(Me.btnPayment)
        Me.gboButton.Location = New System.Drawing.Point(869, -1)
        Me.gboButton.Name = "gboButton"
        Me.gboButton.Size = New System.Drawing.Size(103, 538)
        Me.gboButton.TabIndex = 2
        Me.gboButton.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(6, 415)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(91, 50)
        Me.btnClose.TabIndex = 7
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "F8-Log Out"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnReprint
        '
        Me.btnReprint.Location = New System.Drawing.Point(6, 359)
        Me.btnReprint.Name = "btnReprint"
        Me.btnReprint.Size = New System.Drawing.Size(91, 50)
        Me.btnReprint.TabIndex = 6
        Me.btnReprint.TabStop = False
        Me.btnReprint.Text = "F7-Reprint Receipt"
        Me.btnReprint.UseVisualStyleBackColor = True
        '
        'btnVoid
        '
        Me.btnVoid.Location = New System.Drawing.Point(6, 303)
        Me.btnVoid.Name = "btnVoid"
        Me.btnVoid.Size = New System.Drawing.Size(91, 50)
        Me.btnVoid.TabIndex = 5
        Me.btnVoid.TabStop = False
        Me.btnVoid.Text = "F6-Void"
        Me.btnVoid.UseVisualStyleBackColor = True
        '
        'btnPayOutDropCash
        '
        Me.btnPayOutDropCash.Location = New System.Drawing.Point(6, 71)
        Me.btnPayOutDropCash.Name = "btnPayOutDropCash"
        Me.btnPayOutDropCash.Size = New System.Drawing.Size(91, 50)
        Me.btnPayOutDropCash.TabIndex = 1
        Me.btnPayOutDropCash.TabStop = False
        Me.btnPayOutDropCash.Text = "F2-Payout/Drop Cash"
        Me.btnPayOutDropCash.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(6, 245)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(91, 50)
        Me.btnNext.TabIndex = 4
        Me.btnNext.TabStop = False
        Me.btnNext.Text = "F5-Next Transaction"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(6, 187)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(91, 50)
        Me.btnNew.TabIndex = 3
        Me.btnNew.TabStop = False
        Me.btnNew.Text = "F4-New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(6, 129)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(91, 50)
        Me.btnFind.TabIndex = 2
        Me.btnFind.TabStop = False
        Me.btnFind.Text = "F3-Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnPayment
        '
        Me.btnPayment.Location = New System.Drawing.Point(6, 13)
        Me.btnPayment.Name = "btnPayment"
        Me.btnPayment.Size = New System.Drawing.Size(91, 50)
        Me.btnPayment.TabIndex = 0
        Me.btnPayment.TabStop = False
        Me.btnPayment.Text = "F1-Payment"
        Me.btnPayment.UseVisualStyleBackColor = True
        '
        'TxtDisplay
        '
        Me.TxtDisplay.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!)
        Me.TxtDisplay.ForeColor = System.Drawing.Color.Blue
        Me.TxtDisplay.Location = New System.Drawing.Point(12, 413)
        Me.TxtDisplay.Multiline = True
        Me.TxtDisplay.Name = "TxtDisplay"
        Me.TxtDisplay.Size = New System.Drawing.Size(851, 124)
        Me.TxtDisplay.TabIndex = 1
        Me.TxtDisplay.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(114, 3)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(300, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "ID Number"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(115, 4)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(300, 20)
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Member"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 540)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(984, 22)
        Me.StatusStrip2.TabIndex = 61
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(242, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(242, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(242, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(242, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'Timer1
        '
        Me.Timer1.Interval = 9000
        '
        'FrmTrans_POS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(984, 562)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.gboButton)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.PbxPaid)
        Me.Controls.Add(Me.PbxViod)
        Me.Controls.Add(Me.TxtDisplay)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "FrmTrans_POS"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Point Of Sales"
        Me.gboDetail.ResumeLayout(False)
        Me.spc2.Panel1.ResumeLayout(False)
        Me.spc2.Panel1.PerformLayout()
        Me.spc2.Panel2.ResumeLayout(False)
        Me.spc2.Panel2.PerformLayout()
        Me.spc2.ResumeLayout(False)
        Me.spc1.Panel1.ResumeLayout(False)
        Me.spc1.Panel1.PerformLayout()
        Me.spc1.Panel2.ResumeLayout(False)
        Me.spc1.Panel2.PerformLayout()
        Me.spc1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PbxViod, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PbxPaid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboButton.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents gboButton As System.Windows.Forms.GroupBox
    Friend WithEvents TxtDisplay As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents spc1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtIdNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtMember As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnPayment As System.Windows.Forms.Button
    Friend WithEvents btnPayOutDropCash As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnVoid As System.Windows.Forms.Button
    Friend WithEvents spc2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents LblCodeNo As System.Windows.Forms.Label
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryMember As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnReprint As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySession As System.Windows.Forms.TextBox
    Friend WithEvents btnFindMember As System.Windows.Forms.Button
    Friend WithEvents TxtDiscString As System.Windows.Forms.TextBox
    Friend WithEvents TxtNettPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtDiscAmount As System.Windows.Forms.TextBox
    Friend WithEvents LblDiscount As System.Windows.Forms.Label
    Friend WithEvents TxtUnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents LblUnitPrice As System.Windows.Forms.Label
    Friend WithEvents chkLocationControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDateControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatchNoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDate As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents DtpExpired As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboLocation As System.Windows.Forms.ComboBox
    Friend WithEvents dtpTanggalTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDisc_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSubTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAutoNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PbxPaid As System.Windows.Forms.PictureBox
    Friend WithEvents PbxViod As System.Windows.Forms.PictureBox
End Class
