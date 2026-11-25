<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMast_Discount
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
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtUrutanDiskon = New System.Windows.Forms.TextBox
        Me.ChkMultidiskon = New System.Windows.Forms.CheckBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.cboDisc_Type = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboMember = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.cboStatus = New System.Windows.Forms.ComboBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.ChkShowDiscount = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.tcoProperties = New System.Windows.Forms.TabControl
        Me.tbp_Produk = New System.Windows.Forms.TabPage
        Me.TxtPrimaryQuickAdd = New System.Windows.Forms.TextBox
        Me.BtnInputItem = New System.Windows.Forms.Button
        Me.btnFindQuickAdd = New System.Windows.Forms.Button
        Me.cboQuickAdd = New System.Windows.Forms.ComboBox
        Me.dgvItemProd = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cboProduct = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtQuickAdd = New System.Windows.Forms.TextBox
        Me.tbp_Branch = New System.Windows.Forms.TabPage
        Me.Label5 = New System.Windows.Forms.Label
        Me.dgvBranch = New System.Windows.Forms.DataGridView
        Me.ColBranch_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBranch_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBranch_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBranc_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBranch_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cboBranch = New System.Windows.Forms.ComboBox
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.clbDate = New System.Windows.Forms.CheckedListBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.cboTimeFilter = New System.Windows.Forms.ComboBox
        Me.DtpDateTo = New System.Windows.Forms.DateTimePicker
        Me.DtpDateFrom = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.cboTime_Advance = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.DtpTimeTo = New System.Windows.Forms.DateTimePicker
        Me.DtpTimeFrom = New System.Windows.Forms.DateTimePicker
        Me.Label10 = New System.Windows.Forms.Label
        Me.clbDayOfWeek = New System.Windows.Forms.CheckedListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.gboNominal = New System.Windows.Forms.GroupBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.cboRounding = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtMinQty1 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtDiscountRate = New System.Windows.Forms.TextBox
        Me.ChkAllowToSell = New System.Windows.Forms.CheckBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtMaxQty1 = New System.Windows.Forms.TextBox
        Me.gboGimmick = New System.Windows.Forms.GroupBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtFreeQty = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.TxtBaseQty = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.TxtMinQty2 = New System.Windows.Forms.TextBox
        Me.ChkMultiply = New System.Windows.Forms.CheckBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.TxtMaxQty2 = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.LblAcuanKiri = New System.Windows.Forms.Label
        Me.TxtAcuanKiri = New System.Windows.Forms.TextBox
        Me.LblAcuanKanan = New System.Windows.Forms.Label
        Me.TxtAcuanKanan = New System.Windows.Forms.TextBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.tcoProperties.SuspendLayout()
        Me.tbp_Produk.SuspendLayout()
        CType(Me.dgvItemProd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbp_Branch.SuspendLayout()
        CType(Me.dgvBranch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        Me.gboNominal.SuspendLayout()
        Me.gboGimmick.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(320, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(86, 19)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimary.TabIndex = 23
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(320, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(220, 125)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(77, 20)
        Me.TxtReadOnly.TabIndex = 5
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(86, 45)
        Me.TxtDesc.MaxLength = 150
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(220, 20)
        Me.TxtDesc.TabIndex = 1
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 518)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(656, 22)
        Me.StatusStrip2.TabIndex = 46
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(320, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 540)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(656, 22)
        Me.StatusStrip1.TabIndex = 45
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(320, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtUrutanDiskon)
        Me.gboDetail.Controls.Add(Me.ChkMultidiskon)
        Me.gboDetail.Controls.Add(Me.Label18)
        Me.gboDetail.Controls.Add(Me.cboDisc_Type)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.cboMember)
        Me.gboDetail.Controls.Add(Me.Label14)
        Me.gboDetail.Controls.Add(Me.cboStatus)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.ChkShowDiscount)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.GroupBox1)
        Me.gboDetail.Controls.Add(Me.TxtDesc)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtCode)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.gboNominal)
        Me.gboDetail.Controls.Add(Me.gboGimmick)
        Me.gboDetail.Location = New System.Drawing.Point(17, 12)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(624, 503)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'TxtUrutanDiskon
        '
        Me.TxtUrutanDiskon.Location = New System.Drawing.Point(274, 19)
        Me.TxtUrutanDiskon.Name = "TxtUrutanDiskon"
        Me.TxtUrutanDiskon.Size = New System.Drawing.Size(32, 20)
        Me.TxtUrutanDiskon.TabIndex = 47
        Me.TxtUrutanDiskon.TabStop = False
        Me.TxtUrutanDiskon.Text = "-1"
        Me.TxtUrutanDiskon.Visible = False
        '
        'ChkMultidiskon
        '
        Me.ChkMultidiskon.AutoSize = True
        Me.ChkMultidiskon.Location = New System.Drawing.Point(215, 155)
        Me.ChkMultidiskon.Name = "ChkMultidiskon"
        Me.ChkMultidiskon.Size = New System.Drawing.Size(91, 17)
        Me.ChkMultidiskon.TabIndex = 6
        Me.ChkMultidiskon.Text = "Multi discount"
        Me.ChkMultidiskon.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(16, 74)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(31, 13)
        Me.Label18.TabIndex = 75
        Me.Label18.Text = "Type"
        '
        'cboDisc_Type
        '
        Me.cboDisc_Type.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDisc_Type.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDisc_Type.FormattingEnabled = True
        Me.cboDisc_Type.Items.AddRange(New Object() {"Nominal & Percentage", "Gimmick"})
        Me.cboDisc_Type.Location = New System.Drawing.Point(86, 71)
        Me.cboDisc_Type.Name = "cboDisc_Type"
        Me.cboDisc_Type.Size = New System.Drawing.Size(220, 21)
        Me.cboDisc_Type.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 101)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 13)
        Me.Label3.TabIndex = 71
        Me.Label3.Text = "Membership"
        '
        'cboMember
        '
        Me.cboMember.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboMember.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMember.FormattingEnabled = True
        Me.cboMember.Items.AddRange(New Object() {"Apply For All", "Apply For Member Only"})
        Me.cboMember.Location = New System.Drawing.Point(86, 98)
        Me.cboMember.Name = "cboMember"
        Me.cboMember.Size = New System.Drawing.Size(220, 21)
        Me.cboMember.TabIndex = 3
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(16, 128)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(37, 13)
        Me.Label14.TabIndex = 60
        Me.Label14.Text = "Status"
        '
        'cboStatus
        '
        Me.cboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Items.AddRange(New Object() {"Active", "Not Active"})
        Me.cboStatus.Location = New System.Drawing.Point(86, 125)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(220, 21)
        Me.cboStatus.TabIndex = 4
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Location = New System.Drawing.Point(86, 152)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(123, 20)
        Me.TxtRemarks.TabIndex = 5
        '
        'ChkShowDiscount
        '
        Me.ChkShowDiscount.AutoSize = True
        Me.ChkShowDiscount.Location = New System.Drawing.Point(86, 178)
        Me.ChkShowDiscount.Name = "ChkShowDiscount"
        Me.ChkShowDiscount.Size = New System.Drawing.Size(155, 17)
        Me.ChkShowDiscount.TabIndex = 7
        Me.ChkShowDiscount.Text = "Show Discount On Receipt"
        Me.ChkShowDiscount.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(16, 152)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 59
        Me.Label7.Text = "Remarks"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.tcoProperties)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 203)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(609, 264)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Discount Validity"
        '
        'tcoProperties
        '
        Me.tcoProperties.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tcoProperties.Controls.Add(Me.tbp_Produk)
        Me.tcoProperties.Controls.Add(Me.tbp_Branch)
        Me.tcoProperties.Controls.Add(Me.TabPage1)
        Me.tcoProperties.Location = New System.Drawing.Point(6, 18)
        Me.tcoProperties.Name = "tcoProperties"
        Me.tcoProperties.SelectedIndex = 0
        Me.tcoProperties.Size = New System.Drawing.Size(591, 234)
        Me.tcoProperties.TabIndex = 0
        '
        'tbp_Produk
        '
        Me.tbp_Produk.Controls.Add(Me.TxtPrimaryQuickAdd)
        Me.tbp_Produk.Controls.Add(Me.BtnInputItem)
        Me.tbp_Produk.Controls.Add(Me.btnFindQuickAdd)
        Me.tbp_Produk.Controls.Add(Me.cboQuickAdd)
        Me.tbp_Produk.Controls.Add(Me.dgvItemProd)
        Me.tbp_Produk.Controls.Add(Me.cboProduct)
        Me.tbp_Produk.Controls.Add(Me.Label4)
        Me.tbp_Produk.Controls.Add(Me.TxtQuickAdd)
        Me.tbp_Produk.Location = New System.Drawing.Point(4, 22)
        Me.tbp_Produk.Name = "tbp_Produk"
        Me.tbp_Produk.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp_Produk.Size = New System.Drawing.Size(583, 208)
        Me.tbp_Produk.TabIndex = 0
        Me.tbp_Produk.Text = "Product Data"
        Me.tbp_Produk.UseVisualStyleBackColor = True
        '
        'TxtPrimaryQuickAdd
        '
        Me.TxtPrimaryQuickAdd.Location = New System.Drawing.Point(376, 6)
        Me.TxtPrimaryQuickAdd.Name = "TxtPrimaryQuickAdd"
        Me.TxtPrimaryQuickAdd.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryQuickAdd.TabIndex = 76
        Me.TxtPrimaryQuickAdd.TabStop = False
        Me.TxtPrimaryQuickAdd.Text = "-1"
        Me.TxtPrimaryQuickAdd.Visible = False
        '
        'BtnInputItem
        '
        Me.BtnInputItem.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInputItem.Location = New System.Drawing.Point(538, 6)
        Me.BtnInputItem.Name = "BtnInputItem"
        Me.BtnInputItem.Size = New System.Drawing.Size(42, 22)
        Me.BtnInputItem.TabIndex = 4
        Me.BtnInputItem.Text = "Input"
        Me.BtnInputItem.UseVisualStyleBackColor = True
        '
        'btnFindQuickAdd
        '
        Me.btnFindQuickAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindQuickAdd.Location = New System.Drawing.Point(496, 6)
        Me.btnFindQuickAdd.Name = "btnFindQuickAdd"
        Me.btnFindQuickAdd.Size = New System.Drawing.Size(42, 22)
        Me.btnFindQuickAdd.TabIndex = 3
        Me.btnFindQuickAdd.TabStop = False
        Me.btnFindQuickAdd.Text = "Find"
        Me.btnFindQuickAdd.UseVisualStyleBackColor = True
        '
        'cboQuickAdd
        '
        Me.cboQuickAdd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboQuickAdd.FormattingEnabled = True
        Me.cboQuickAdd.Items.AddRange(New Object() {"Add By Product Group", "Add By Brand", "Add By Sub Brand", "Add By Category"})
        Me.cboQuickAdd.Location = New System.Drawing.Point(232, 6)
        Me.cboQuickAdd.Name = "cboQuickAdd"
        Me.cboQuickAdd.Size = New System.Drawing.Size(138, 21)
        Me.cboQuickAdd.TabIndex = 1
        '
        'dgvItemProd
        '
        Me.dgvItemProd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemProd.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProd.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColRemarks, Me.ColPrimaryProduct})
        Me.dgvItemProd.Location = New System.Drawing.Point(3, 33)
        Me.dgvItemProd.Name = "dgvItemProd"
        Me.dgvItemProd.Size = New System.Drawing.Size(577, 169)
        Me.dgvItemProd.TabIndex = 5
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.Width = 85
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryProduct
        '
        Me.ColPrimaryProduct.HeaderText = "Primary"
        Me.ColPrimaryProduct.Name = "ColPrimaryProduct"
        Me.ColPrimaryProduct.Visible = False
        Me.ColPrimaryProduct.Width = 66
        '
        'cboProduct
        '
        Me.cboProduct.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProduct.FormattingEnabled = True
        Me.cboProduct.Items.AddRange(New Object() {"Apply For All Item", "Apply For Spesific Item Only"})
        Me.cboProduct.Location = New System.Drawing.Point(70, 6)
        Me.cboProduct.Name = "cboProduct"
        Me.cboProduct.Size = New System.Drawing.Size(156, 21)
        Me.cboProduct.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(0, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 41
        Me.Label4.Text = "Product"
        '
        'TxtQuickAdd
        '
        Me.TxtQuickAdd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQuickAdd.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtQuickAdd.Location = New System.Drawing.Point(376, 6)
        Me.TxtQuickAdd.MaxLength = 150
        Me.TxtQuickAdd.Name = "TxtQuickAdd"
        Me.TxtQuickAdd.Size = New System.Drawing.Size(114, 20)
        Me.TxtQuickAdd.TabIndex = 2
        '
        'tbp_Branch
        '
        Me.tbp_Branch.Controls.Add(Me.Label5)
        Me.tbp_Branch.Controls.Add(Me.dgvBranch)
        Me.tbp_Branch.Controls.Add(Me.cboBranch)
        Me.tbp_Branch.Location = New System.Drawing.Point(4, 22)
        Me.tbp_Branch.Name = "tbp_Branch"
        Me.tbp_Branch.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp_Branch.Size = New System.Drawing.Size(583, 208)
        Me.tbp_Branch.TabIndex = 1
        Me.tbp_Branch.Text = "Branch Data"
        Me.tbp_Branch.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(0, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 13)
        Me.Label5.TabIndex = 43
        Me.Label5.Text = "Branch"
        '
        'dgvBranch
        '
        Me.dgvBranch.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvBranch.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvBranch.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBranch_No, Me.ColBranch_Code, Me.ColBranch_Desc, Me.ColBranc_Remarks, Me.ColBranch_Primary})
        Me.dgvBranch.Location = New System.Drawing.Point(3, 33)
        Me.dgvBranch.Name = "dgvBranch"
        Me.dgvBranch.Size = New System.Drawing.Size(577, 169)
        Me.dgvBranch.TabIndex = 2
        '
        'ColBranch_No
        '
        Me.ColBranch_No.HeaderText = "No"
        Me.ColBranch_No.Name = "ColBranch_No"
        Me.ColBranch_No.Width = 46
        '
        'ColBranch_Code
        '
        Me.ColBranch_Code.HeaderText = "Branch Code"
        Me.ColBranch_Code.Name = "ColBranch_Code"
        Me.ColBranch_Code.Width = 94
        '
        'ColBranch_Desc
        '
        Me.ColBranch_Desc.HeaderText = "Description"
        Me.ColBranch_Desc.Name = "ColBranch_Desc"
        Me.ColBranch_Desc.Width = 85
        '
        'ColBranc_Remarks
        '
        Me.ColBranc_Remarks.HeaderText = "Remarks"
        Me.ColBranc_Remarks.Name = "ColBranc_Remarks"
        Me.ColBranc_Remarks.Width = 74
        '
        'ColBranch_Primary
        '
        Me.ColBranch_Primary.HeaderText = "Primary"
        Me.ColBranch_Primary.Name = "ColBranch_Primary"
        Me.ColBranch_Primary.Visible = False
        Me.ColBranch_Primary.Width = 66
        '
        'cboBranch
        '
        Me.cboBranch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBranch.FormattingEnabled = True
        Me.cboBranch.Items.AddRange(New Object() {"Apply For All Branch", "Apply For Specific Branch"})
        Me.cboBranch.Location = New System.Drawing.Point(95, 6)
        Me.cboBranch.Name = "cboBranch"
        Me.cboBranch.Size = New System.Drawing.Size(195, 21)
        Me.cboBranch.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.clbDate)
        Me.TabPage1.Controls.Add(Me.Label13)
        Me.TabPage1.Controls.Add(Me.cboTimeFilter)
        Me.TabPage1.Controls.Add(Me.DtpDateTo)
        Me.TabPage1.Controls.Add(Me.DtpDateFrom)
        Me.TabPage1.Controls.Add(Me.Label12)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.cboTime_Advance)
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.Label9)
        Me.TabPage1.Controls.Add(Me.DtpTimeTo)
        Me.TabPage1.Controls.Add(Me.DtpTimeFrom)
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.clbDayOfWeek)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(583, 208)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "Time"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'clbDate
        '
        Me.clbDate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbDate.CheckOnClick = True
        Me.clbDate.ColumnWidth = 75
        Me.clbDate.FormattingEnabled = True
        Me.clbDate.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"})
        Me.clbDate.Location = New System.Drawing.Point(332, 7)
        Me.clbDate.MultiColumn = True
        Me.clbDate.Name = "clbDate"
        Me.clbDate.Size = New System.Drawing.Size(245, 169)
        Me.clbDate.TabIndex = 55
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(1, 9)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(37, 13)
        Me.Label13.TabIndex = 57
        Me.Label13.Text = "Period"
        '
        'cboTimeFilter
        '
        Me.cboTimeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTimeFilter.FormattingEnabled = True
        Me.cboTimeFilter.Items.AddRange(New Object() {"Always Valid", "Valid For Certain Period"})
        Me.cboTimeFilter.Location = New System.Drawing.Point(96, 6)
        Me.cboTimeFilter.Name = "cboTimeFilter"
        Me.cboTimeFilter.Size = New System.Drawing.Size(230, 21)
        Me.cboTimeFilter.TabIndex = 0
        '
        'DtpDateTo
        '
        Me.DtpDateTo.CustomFormat = "dd MMM yy"
        Me.DtpDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDateTo.Location = New System.Drawing.Point(226, 34)
        Me.DtpDateTo.Name = "DtpDateTo"
        Me.DtpDateTo.Size = New System.Drawing.Size(99, 20)
        Me.DtpDateTo.TabIndex = 2
        '
        'DtpDateFrom
        '
        Me.DtpDateFrom.CustomFormat = "dd MMM yy"
        Me.DtpDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDateFrom.Location = New System.Drawing.Point(96, 33)
        Me.DtpDateFrom.Name = "DtpDateFrom"
        Me.DtpDateFrom.Size = New System.Drawing.Size(98, 20)
        Me.DtpDateFrom.TabIndex = 1
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(200, 36)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(20, 13)
        Me.Label12.TabIndex = 52
        Me.Label12.Text = "To"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(4, 90)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(89, 13)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "Advance Filtering"
        '
        'cboTime_Advance
        '
        Me.cboTime_Advance.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTime_Advance.FormattingEnabled = True
        Me.cboTime_Advance.Items.AddRange(New Object() {"Apply Always", "Apply On Specific Day Of The Week", "Apply On Specific Date Of The Month"})
        Me.cboTime_Advance.Location = New System.Drawing.Point(96, 87)
        Me.cboTime_Advance.Name = "cboTime_Advance"
        Me.cboTime_Advance.Size = New System.Drawing.Size(230, 21)
        Me.cboTime_Advance.TabIndex = 5
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 36)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(30, 13)
        Me.Label11.TabIndex = 50
        Me.Label11.Text = "Date"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 63)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 13)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "Time"
        '
        'DtpTimeTo
        '
        Me.DtpTimeTo.CustomFormat = "HH:mm"
        Me.DtpTimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTimeTo.Location = New System.Drawing.Point(226, 61)
        Me.DtpTimeTo.Name = "DtpTimeTo"
        Me.DtpTimeTo.ShowUpDown = True
        Me.DtpTimeTo.Size = New System.Drawing.Size(99, 20)
        Me.DtpTimeTo.TabIndex = 4
        '
        'DtpTimeFrom
        '
        Me.DtpTimeFrom.CustomFormat = "HH:mm"
        Me.DtpTimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTimeFrom.Location = New System.Drawing.Point(96, 60)
        Me.DtpTimeFrom.Name = "DtpTimeFrom"
        Me.DtpTimeFrom.ShowUpDown = True
        Me.DtpTimeFrom.Size = New System.Drawing.Size(98, 20)
        Me.DtpTimeFrom.TabIndex = 3
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(200, 63)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(20, 13)
        Me.Label10.TabIndex = 48
        Me.Label10.Text = "To"
        '
        'clbDayOfWeek
        '
        Me.clbDayOfWeek.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbDayOfWeek.CheckOnClick = True
        Me.clbDayOfWeek.FormattingEnabled = True
        Me.clbDayOfWeek.Items.AddRange(New Object() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"})
        Me.clbDayOfWeek.Location = New System.Drawing.Point(332, 7)
        Me.clbDayOfWeek.Name = "clbDayOfWeek"
        Me.clbDayOfWeek.Size = New System.Drawing.Size(245, 169)
        Me.clbDayOfWeek.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Description"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Code"
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(199, 471)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(86, 19)
        Me.TxtCode.MaxLength = 150
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(220, 20)
        Me.TxtCode.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(315, 471)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'gboNominal
        '
        Me.gboNominal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboNominal.Controls.Add(Me.Label16)
        Me.gboNominal.Controls.Add(Me.cboRounding)
        Me.gboNominal.Controls.Add(Me.Label17)
        Me.gboNominal.Controls.Add(Me.TxtMinQty1)
        Me.gboNominal.Controls.Add(Me.Label6)
        Me.gboNominal.Controls.Add(Me.TxtDiscountRate)
        Me.gboNominal.Controls.Add(Me.ChkAllowToSell)
        Me.gboNominal.Controls.Add(Me.Label15)
        Me.gboNominal.Controls.Add(Me.TxtMaxQty1)
        Me.gboNominal.Controls.Add(Me.TxtReadOnly)
        Me.gboNominal.Location = New System.Drawing.Point(312, 19)
        Me.gboNominal.Name = "gboNominal"
        Me.gboNominal.Size = New System.Drawing.Size(306, 176)
        Me.gboNominal.TabIndex = 9
        Me.gboNominal.TabStop = False
        Me.gboNominal.Text = "Detail"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(11, 74)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(53, 13)
        Me.Label16.TabIndex = 83
        Me.Label16.Text = "Rounding"
        '
        'cboRounding
        '
        Me.cboRounding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRounding.FormattingEnabled = True
        Me.cboRounding.Items.AddRange(New Object() {"Up", "Down", "Middle"})
        Me.cboRounding.Location = New System.Drawing.Point(106, 71)
        Me.cboRounding.Name = "cboRounding"
        Me.cboRounding.Size = New System.Drawing.Size(191, 21)
        Me.cboRounding.TabIndex = 4
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(184, 48)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(43, 13)
        Me.Label17.TabIndex = 82
        Me.Label17.Text = "Min Qty"
        '
        'TxtMinQty1
        '
        Me.TxtMinQty1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMinQty1.Location = New System.Drawing.Point(239, 45)
        Me.TxtMinQty1.MaxLength = 12
        Me.TxtMinQty1.Name = "TxtMinQty1"
        Me.TxtMinQty1.Size = New System.Drawing.Size(59, 20)
        Me.TxtMinQty1.TabIndex = 2
        Me.TxtMinQty1.Text = "0"
        Me.TxtMinQty1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 13)
        Me.Label6.TabIndex = 80
        Me.Label6.Text = "Discount Rate"
        '
        'TxtDiscountRate
        '
        Me.TxtDiscountRate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDiscountRate.Location = New System.Drawing.Point(106, 19)
        Me.TxtDiscountRate.MaxLength = 12
        Me.TxtDiscountRate.Name = "TxtDiscountRate"
        Me.TxtDiscountRate.Size = New System.Drawing.Size(191, 20)
        Me.TxtDiscountRate.TabIndex = 0
        Me.TxtDiscountRate.Text = "0"
        Me.TxtDiscountRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkAllowToSell
        '
        Me.ChkAllowToSell.AutoSize = True
        Me.ChkAllowToSell.Location = New System.Drawing.Point(106, 70)
        Me.ChkAllowToSell.Name = "ChkAllowToSell"
        Me.ChkAllowToSell.Size = New System.Drawing.Size(184, 17)
        Me.ChkAllowToSell.TabIndex = 3
        Me.ChkAllowToSell.TabStop = False
        Me.ChkAllowToSell.Text = "Allow To Sell More Than Max Qty"
        Me.ChkAllowToSell.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(10, 48)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(49, 13)
        Me.Label15.TabIndex = 77
        Me.Label15.Text = "Max Qty "
        '
        'TxtMaxQty1
        '
        Me.TxtMaxQty1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMaxQty1.Location = New System.Drawing.Point(106, 45)
        Me.TxtMaxQty1.MaxLength = 12
        Me.TxtMaxQty1.Name = "TxtMaxQty1"
        Me.TxtMaxQty1.Size = New System.Drawing.Size(64, 20)
        Me.TxtMaxQty1.TabIndex = 1
        Me.TxtMaxQty1.Text = "0"
        Me.TxtMaxQty1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'gboGimmick
        '
        Me.gboGimmick.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboGimmick.Controls.Add(Me.Label21)
        Me.gboGimmick.Controls.Add(Me.TxtFreeQty)
        Me.gboGimmick.Controls.Add(Me.Label23)
        Me.gboGimmick.Controls.Add(Me.TxtBaseQty)
        Me.gboGimmick.Controls.Add(Me.Label20)
        Me.gboGimmick.Controls.Add(Me.TxtMinQty2)
        Me.gboGimmick.Controls.Add(Me.ChkMultiply)
        Me.gboGimmick.Controls.Add(Me.Label22)
        Me.gboGimmick.Controls.Add(Me.TxtMaxQty2)
        Me.gboGimmick.Location = New System.Drawing.Point(312, 19)
        Me.gboGimmick.Name = "gboGimmick"
        Me.gboGimmick.Size = New System.Drawing.Size(306, 176)
        Me.gboGimmick.TabIndex = 8
        Me.gboGimmick.TabStop = False
        Me.gboGimmick.Text = "Detail"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(184, 22)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(47, 13)
        Me.Label21.TabIndex = 88
        Me.Label21.Text = "Free Qty"
        '
        'TxtFreeQty
        '
        Me.TxtFreeQty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFreeQty.Location = New System.Drawing.Point(239, 19)
        Me.TxtFreeQty.MaxLength = 12
        Me.TxtFreeQty.Name = "TxtFreeQty"
        Me.TxtFreeQty.Size = New System.Drawing.Size(64, 20)
        Me.TxtFreeQty.TabIndex = 1
        Me.TxtFreeQty.Text = "0"
        Me.TxtFreeQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(10, 22)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(53, 13)
        Me.Label23.TabIndex = 86
        Me.Label23.Text = "Base Qty "
        '
        'TxtBaseQty
        '
        Me.TxtBaseQty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBaseQty.Location = New System.Drawing.Point(106, 19)
        Me.TxtBaseQty.MaxLength = 12
        Me.TxtBaseQty.Name = "TxtBaseQty"
        Me.TxtBaseQty.Size = New System.Drawing.Size(64, 20)
        Me.TxtBaseQty.TabIndex = 0
        Me.TxtBaseQty.Text = "0"
        Me.TxtBaseQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(184, 48)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(43, 13)
        Me.Label20.TabIndex = 82
        Me.Label20.Text = "Min Qty"
        '
        'TxtMinQty2
        '
        Me.TxtMinQty2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMinQty2.Location = New System.Drawing.Point(239, 45)
        Me.TxtMinQty2.MaxLength = 12
        Me.TxtMinQty2.Name = "TxtMinQty2"
        Me.TxtMinQty2.Size = New System.Drawing.Size(64, 20)
        Me.TxtMinQty2.TabIndex = 3
        Me.TxtMinQty2.Text = "0"
        Me.TxtMinQty2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkMultiply
        '
        Me.ChkMultiply.AutoSize = True
        Me.ChkMultiply.Location = New System.Drawing.Point(106, 70)
        Me.ChkMultiply.Name = "ChkMultiply"
        Me.ChkMultiply.Size = New System.Drawing.Size(61, 17)
        Me.ChkMultiply.TabIndex = 4
        Me.ChkMultiply.Text = "Multiply"
        Me.ChkMultiply.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(10, 48)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(49, 13)
        Me.Label22.TabIndex = 77
        Me.Label22.Text = "Max Qty "
        '
        'TxtMaxQty2
        '
        Me.TxtMaxQty2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMaxQty2.Location = New System.Drawing.Point(106, 45)
        Me.TxtMaxQty2.MaxLength = 12
        Me.TxtMaxQty2.Name = "TxtMaxQty2"
        Me.TxtMaxQty2.Size = New System.Drawing.Size(64, 20)
        Me.TxtMaxQty2.TabIndex = 2
        Me.TxtMaxQty2.Text = "0"
        Me.TxtMaxQty2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(23, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblAcuanKiri)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtAcuanKiri)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblAcuanKanan)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAcuanKanan)
        Me.SplitContainer1.Size = New System.Drawing.Size(609, 32)
        Me.SplitContainer1.SplitterDistance = 303
        Me.SplitContainer1.TabIndex = 38
        Me.SplitContainer1.TabStop = False
        Me.SplitContainer1.Visible = False
        '
        'LblAcuanKiri
        '
        Me.LblAcuanKiri.AutoSize = True
        Me.LblAcuanKiri.Location = New System.Drawing.Point(9, 6)
        Me.LblAcuanKiri.Name = "LblAcuanKiri"
        Me.LblAcuanKiri.Size = New System.Drawing.Size(52, 13)
        Me.LblAcuanKiri.TabIndex = 40
        Me.LblAcuanKiri.Text = "LblAcuan"
        '
        'TxtAcuanKiri
        '
        Me.TxtAcuanKiri.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAcuanKiri.Location = New System.Drawing.Point(105, 3)
        Me.TxtAcuanKiri.MaxLength = 150
        Me.TxtAcuanKiri.Name = "TxtAcuanKiri"
        Me.TxtAcuanKiri.Size = New System.Drawing.Size(195, 20)
        Me.TxtAcuanKiri.TabIndex = 39
        Me.TxtAcuanKiri.TabStop = False
        Me.TxtAcuanKiri.Visible = False
        '
        'LblAcuanKanan
        '
        Me.LblAcuanKanan.AutoSize = True
        Me.LblAcuanKanan.Location = New System.Drawing.Point(3, 6)
        Me.LblAcuanKanan.Name = "LblAcuanKanan"
        Me.LblAcuanKanan.Size = New System.Drawing.Size(52, 13)
        Me.LblAcuanKanan.TabIndex = 41
        Me.LblAcuanKanan.Text = "LblAcuan"
        '
        'TxtAcuanKanan
        '
        Me.TxtAcuanKanan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAcuanKanan.Location = New System.Drawing.Point(98, 3)
        Me.TxtAcuanKanan.MaxLength = 150
        Me.TxtAcuanKanan.Name = "TxtAcuanKanan"
        Me.TxtAcuanKanan.Size = New System.Drawing.Size(198, 20)
        Me.TxtAcuanKanan.TabIndex = 40
        Me.TxtAcuanKanan.TabStop = False
        Me.TxtAcuanKanan.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'FrmMast_Discount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(656, 562)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(672, 600)
        Me.Name = "FrmMast_Discount"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Discount"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.tcoProperties.ResumeLayout(False)
        Me.tbp_Produk.ResumeLayout(False)
        Me.tbp_Produk.PerformLayout()
        CType(Me.dgvItemProd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbp_Branch.ResumeLayout(False)
        Me.tbp_Branch.PerformLayout()
        CType(Me.dgvBranch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.gboNominal.ResumeLayout(False)
        Me.gboNominal.PerformLayout()
        Me.gboGimmick.ResumeLayout(False)
        Me.gboGimmick.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents LblAcuanKiri As System.Windows.Forms.Label
    Friend WithEvents TxtAcuanKiri As System.Windows.Forms.TextBox
    Friend WithEvents LblAcuanKanan As System.Windows.Forms.Label
    Friend WithEvents TxtAcuanKanan As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpDateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents DtpDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DtpTimeTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents DtpTimeFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboTime_Advance As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboBranch As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboProduct As System.Windows.Forms.ComboBox
    Friend WithEvents tcoProperties As System.Windows.Forms.TabControl
    Friend WithEvents tbp_Produk As System.Windows.Forms.TabPage
    Friend WithEvents tbp_Branch As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemProd As System.Windows.Forms.DataGridView
    Friend WithEvents dgvBranch As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents clbDate As System.Windows.Forms.CheckedListBox
    Friend WithEvents clbDayOfWeek As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboTimeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cboDisc_Type As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboMember As System.Windows.Forms.ComboBox
    Friend WithEvents gboNominal As System.Windows.Forms.GroupBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cboRounding As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtMinQty1 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtDiscountRate As System.Windows.Forms.TextBox
    Friend WithEvents ChkAllowToSell As System.Windows.Forms.CheckBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtMaxQty1 As System.Windows.Forms.TextBox
    Friend WithEvents ChkShowDiscount As System.Windows.Forms.CheckBox
    Friend WithEvents gboGimmick As System.Windows.Forms.GroupBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtFreeQty As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TxtBaseQty As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TxtMinQty2 As System.Windows.Forms.TextBox
    Friend WithEvents ChkMultiply As System.Windows.Forms.CheckBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtMaxQty2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtQuickAdd As System.Windows.Forms.TextBox
    Friend WithEvents cboQuickAdd As System.Windows.Forms.ComboBox
    Friend WithEvents btnFindQuickAdd As System.Windows.Forms.Button
    Friend WithEvents ChkMultidiskon As System.Windows.Forms.CheckBox
    Friend WithEvents BtnInputItem As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryQuickAdd As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBranch_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBranch_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBranch_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBranc_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBranch_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtUrutanDiskon As System.Windows.Forms.TextBox
End Class
