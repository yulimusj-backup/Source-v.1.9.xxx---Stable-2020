<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseReturnDetail
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.txtProductCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtProductName = New System.Windows.Forms.TextBox
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.txtSupplierName = New System.Windows.Forms.TextBox
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.ChkAllProductPRTN = New System.Windows.Forms.CheckBox
        Me.ChkShowRatio = New System.Windows.Forms.CheckBox
        Me.TabUserKelompok = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        Me.rboProduct = New System.Windows.Forms.RadioButton
        Me.rboPRTNNumber = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.CboTipePRTN = New System.Windows.Forms.ComboBox
        Me.ChkAOPRevision = New System.Windows.Forms.CheckBox
        Me.ChkNotaRetur = New System.Windows.Forms.CheckBox
        Me.ChkAOPCancellation = New System.Windows.Forms.CheckBox
        Me.TabUserKelompok.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(806, 37)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(50, 23)
        Me.btnClose.TabIndex = 27
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(806, 8)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(50, 23)
        Me.btnFilter.TabIndex = 26
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(238, 4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(126, 20)
        Me.dtpTo.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(209, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(78, 4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(125, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(0, 195)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(866, 368)
        Me.crv.TabIndex = 28
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'txtProductCode
        '
        Me.txtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProductCode.Location = New System.Drawing.Point(79, 54)
        Me.txtProductCode.Name = "txtProductCode"
        Me.txtProductCode.Size = New System.Drawing.Size(103, 20)
        Me.txtProductCode.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Product"
        '
        'txtProductName
        '
        Me.txtProductName.Location = New System.Drawing.Point(225, 54)
        Me.txtProductName.Name = "txtProductName"
        Me.txtProductName.ReadOnly = True
        Me.txtProductName.Size = New System.Drawing.Size(213, 20)
        Me.txtProductName.TabIndex = 12
        Me.txtProductName.TabStop = False
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Location = New System.Drawing.Point(188, 55)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(31, 18)
        Me.btnFindProduct.TabIndex = 11
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "..."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Supplier"
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(188, 80)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(31, 18)
        Me.btnFindSupplier.TabIndex = 16
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "..."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'txtSupplierName
        '
        Me.txtSupplierName.Location = New System.Drawing.Point(225, 79)
        Me.txtSupplierName.Name = "txtSupplierName"
        Me.txtSupplierName.ReadOnly = True
        Me.txtSupplierName.Size = New System.Drawing.Size(213, 20)
        Me.txtSupplierName.TabIndex = 17
        Me.txtSupplierName.TabStop = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(79, 79)
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.Size = New System.Drawing.Size(103, 20)
        Me.txtSupplierCode.TabIndex = 15
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Location = New System.Drawing.Point(79, 29)
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(103, 20)
        Me.txtCurrencyCode.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 33)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Currency"
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(188, 30)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(31, 18)
        Me.btnFindCurrency.TabIndex = 7
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "..."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'txtPrimaryProduct
        '
        Me.txtPrimaryProduct.Location = New System.Drawing.Point(42, 54)
        Me.txtPrimaryProduct.Name = "txtPrimaryProduct"
        Me.txtPrimaryProduct.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimaryProduct.TabIndex = 9
        Me.txtPrimaryProduct.Text = "-1"
        Me.txtPrimaryProduct.Visible = False
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(42, 79)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimarySupplier.TabIndex = 14
        Me.txtPrimarySupplier.Text = "-1"
        Me.txtPrimarySupplier.Visible = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(42, 29)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimaryCurrency.TabIndex = 5
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'ChkAllProductPRTN
        '
        Me.ChkAllProductPRTN.AutoSize = True
        Me.ChkAllProductPRTN.Location = New System.Drawing.Point(130, 105)
        Me.ChkAllProductPRTN.Name = "ChkAllProductPRTN"
        Me.ChkAllProductPRTN.Size = New System.Drawing.Size(167, 17)
        Me.ChkAllProductPRTN.TabIndex = 20
        Me.ChkAllProductPRTN.Text = "Always display all PRTN items"
        Me.ChkAllProductPRTN.UseVisualStyleBackColor = True
        '
        'ChkShowRatio
        '
        Me.ChkShowRatio.AutoSize = True
        Me.ChkShowRatio.Location = New System.Drawing.Point(130, 128)
        Me.ChkShowRatio.Name = "ChkShowRatio"
        Me.ChkShowRatio.Size = New System.Drawing.Size(142, 17)
        Me.ChkShowRatio.TabIndex = 21
        Me.ChkShowRatio.Text = "Show Ratio (If Available)"
        Me.ChkShowRatio.UseVisualStyleBackColor = True
        '
        'TabUserKelompok
        '
        Me.TabUserKelompok.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabUserKelompok.Controls.Add(Me.TabPage1)
        Me.TabUserKelompok.Controls.Add(Me.TabPage2)
        Me.TabUserKelompok.Location = New System.Drawing.Point(444, 3)
        Me.TabUserKelompok.Name = "TabUserKelompok"
        Me.TabUserKelompok.SelectedIndex = 0
        Me.TabUserKelompok.Size = New System.Drawing.Size(356, 175)
        Me.TabUserKelompok.TabIndex = 25
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvUser)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(348, 149)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Created By"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_Primary, Me.ColCreatedBy, Me.ColCreatedBy_Name})
        Me.DgvUser.Location = New System.Drawing.Point(2, 2)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(344, 145)
        Me.DgvUser.TabIndex = 0
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
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(348, 149)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Product Group"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(0, 0)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(348, 181)
        Me.dgvItemProductGroup.TabIndex = 177
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 27
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 70
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 66
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 55
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'BackgroundWorker3
        '
        '
        'BgwUser
        '
        '
        'rboProduct
        '
        Me.rboProduct.AutoSize = True
        Me.rboProduct.Location = New System.Drawing.Point(297, 128)
        Me.rboProduct.Margin = New System.Windows.Forms.Padding(2)
        Me.rboProduct.Name = "rboProduct"
        Me.rboProduct.Size = New System.Drawing.Size(109, 17)
        Me.rboProduct.TabIndex = 24
        Me.rboProduct.Text = "Group By Product"
        Me.rboProduct.UseVisualStyleBackColor = True
        '
        'rboPRTNNumber
        '
        Me.rboPRTNNumber.AutoSize = True
        Me.rboPRTNNumber.Checked = True
        Me.rboPRTNNumber.Location = New System.Drawing.Point(297, 105)
        Me.rboPRTNNumber.Margin = New System.Windows.Forms.Padding(2)
        Me.rboPRTNNumber.Name = "rboPRTNNumber"
        Me.rboPRTNNumber.Size = New System.Drawing.Size(142, 17)
        Me.rboPRTNNumber.TabIndex = 23
        Me.rboPRTNNumber.TabStop = True
        Me.rboPRTNNumber.Text = "Group By PRTN Number"
        Me.rboPRTNNumber.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(158, 161)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 13)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Tipe PRTN"
        Me.Label3.Visible = False
        '
        'CboTipePRTN
        '
        Me.CboTipePRTN.FormattingEnabled = True
        Me.CboTipePRTN.Items.AddRange(New Object() {"All", "AOP Revision", "Tax Return Notes", "AOP Cancellation"})
        Me.CboTipePRTN.Location = New System.Drawing.Point(225, 153)
        Me.CboTipePRTN.Name = "CboTipePRTN"
        Me.CboTipePRTN.Size = New System.Drawing.Size(104, 21)
        Me.CboTipePRTN.TabIndex = 30
        Me.CboTipePRTN.Text = "All"
        Me.CboTipePRTN.Visible = False
        '
        'ChkAOPRevision
        '
        Me.ChkAOPRevision.AutoSize = True
        Me.ChkAOPRevision.Checked = True
        Me.ChkAOPRevision.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkAOPRevision.Location = New System.Drawing.Point(11, 105)
        Me.ChkAOPRevision.Name = "ChkAOPRevision"
        Me.ChkAOPRevision.Size = New System.Drawing.Size(92, 17)
        Me.ChkAOPRevision.TabIndex = 31
        Me.ChkAOPRevision.Text = "AOP Revision"
        Me.ChkAOPRevision.UseVisualStyleBackColor = True
        '
        'ChkNotaRetur
        '
        Me.ChkNotaRetur.AutoSize = True
        Me.ChkNotaRetur.Checked = True
        Me.ChkNotaRetur.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkNotaRetur.Location = New System.Drawing.Point(11, 151)
        Me.ChkNotaRetur.Name = "ChkNotaRetur"
        Me.ChkNotaRetur.Size = New System.Drawing.Size(110, 17)
        Me.ChkNotaRetur.TabIndex = 32
        Me.ChkNotaRetur.Text = "Tax Return Notes"
        Me.ChkNotaRetur.UseVisualStyleBackColor = True
        '
        'ChkAOPCancellation
        '
        Me.ChkAOPCancellation.AutoSize = True
        Me.ChkAOPCancellation.Checked = True
        Me.ChkAOPCancellation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkAOPCancellation.Location = New System.Drawing.Point(11, 128)
        Me.ChkAOPCancellation.Name = "ChkAOPCancellation"
        Me.ChkAOPCancellation.Size = New System.Drawing.Size(109, 17)
        Me.ChkAOPCancellation.TabIndex = 33
        Me.ChkAOPCancellation.Text = "AOP Cancellation"
        Me.ChkAOPCancellation.UseVisualStyleBackColor = True
        '
        'FrmLaporanPurchaseReturnDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(868, 562)
        Me.Controls.Add(Me.ChkAOPCancellation)
        Me.Controls.Add(Me.ChkNotaRetur)
        Me.Controls.Add(Me.ChkAOPRevision)
        Me.Controls.Add(Me.CboTipePRTN)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.rboPRTNNumber)
        Me.Controls.Add(Me.rboProduct)
        Me.Controls.Add(Me.TabUserKelompok)
        Me.Controls.Add(Me.ChkShowRatio)
        Me.Controls.Add(Me.ChkAllProductPRTN)
        Me.Controls.Add(Me.txtPrimaryCurrency)
        Me.Controls.Add(Me.txtPrimarySupplier)
        Me.Controls.Add(Me.txtPrimaryProduct)
        Me.Controls.Add(Me.btnFindCurrency)
        Me.Controls.Add(Me.txtCurrencyCode)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnFindSupplier)
        Me.Controls.Add(Me.txtSupplierName)
        Me.Controls.Add(Me.txtSupplierCode)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnFindProduct)
        Me.Controls.Add(Me.txtProductName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtProductCode)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.crv)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanPurchaseReturnDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Laporan Purchase Return Detail"
        Me.TabUserKelompok.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtProductName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents txtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents ChkAllProductPRTN As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowRatio As System.Windows.Forms.CheckBox
    Friend WithEvents TabUserKelompok As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents rboProduct As System.Windows.Forms.RadioButton
    Friend WithEvents rboPRTNNumber As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CboTipePRTN As System.Windows.Forms.ComboBox
    Friend WithEvents ChkAOPRevision As System.Windows.Forms.CheckBox
    Friend WithEvents ChkNotaRetur As System.Windows.Forms.CheckBox
    Friend WithEvents ChkAOPCancellation As System.Windows.Forms.CheckBox
End Class
