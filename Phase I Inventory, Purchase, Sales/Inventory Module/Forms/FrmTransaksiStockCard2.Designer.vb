<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockCard2
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.CboGroup = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CboScope = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DgvLocation = New System.Windows.Forms.DataGridView
        Me.Col1No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Col1Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkShowItemRemarks = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.txtProductCode = New System.Windows.Forms.TextBox
        Me.ChkShowBatch = New System.Windows.Forms.CheckBox
        Me.txtProductGroupCode = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.txtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.lblProductGroup = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblProductCode = New System.Windows.Forms.Label
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.txtPrimaryBrand = New System.Windows.Forms.TextBox
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.lblSubBrand = New System.Windows.Forms.Label
        Me.btnFindBrand = New System.Windows.Forms.Button
        Me.txtSubBrandCode = New System.Windows.Forms.TextBox
        Me.txtBrandCode = New System.Windows.Forms.TextBox
        Me.lblCategory = New System.Windows.Forms.Label
        Me.btnFindSubBrand = New System.Windows.Forms.Button
        Me.lblBrand = New System.Windows.Forms.Label
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.txtCategoryCode = New System.Windows.Forms.TextBox
        Me.txtPrimarySubBrand = New System.Windows.Forms.TextBox
        Me.btnFindCategory = New System.Windows.Forms.Button
        Me.txtPrimaryProductGroup = New System.Windows.Forms.TextBox
        Me.btnFindProductGroup = New System.Windows.Forms.Button
        Me.txtBrandDesc = New System.Windows.Forms.TextBox
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.txtCategoryDesc = New System.Windows.Forms.TextBox
        Me.txtSubBrandDesc = New System.Windows.Forms.TextBox
        Me.txtProductGroupDesc = New System.Windows.Forms.TextBox
        Me.crvStockCard = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BgwLocation = New System.ComponentModel.BackgroundWorker
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.gboDetail.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.LinkLabel2)
        Me.gboDetail.Controls.Add(Me.LinkLabel1)
        Me.gboDetail.Controls.Add(Me.CboGroup)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.CboScope)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.GroupBox1)
        Me.gboDetail.Controls.Add(Me.ChkShowItemRemarks)
        Me.gboDetail.Controls.Add(Me.CheckBox2)
        Me.gboDetail.Controls.Add(Me.txtProductCode)
        Me.gboDetail.Controls.Add(Me.ChkShowBatch)
        Me.gboDetail.Controls.Add(Me.txtProductGroupCode)
        Me.gboDetail.Controls.Add(Me.btnClose)
        Me.gboDetail.Controls.Add(Me.txtPrimaryProduct)
        Me.gboDetail.Controls.Add(Me.btnFilter)
        Me.gboDetail.Controls.Add(Me.lblProductGroup)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.lblProductCode)
        Me.gboDetail.Controls.Add(Me.DtpTo)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.DtpFrom)
        Me.gboDetail.Controls.Add(Me.txtPrimaryBrand)
        Me.gboDetail.Controls.Add(Me.btnFindProduct)
        Me.gboDetail.Controls.Add(Me.lblSubBrand)
        Me.gboDetail.Controls.Add(Me.btnFindBrand)
        Me.gboDetail.Controls.Add(Me.txtSubBrandCode)
        Me.gboDetail.Controls.Add(Me.txtBrandCode)
        Me.gboDetail.Controls.Add(Me.lblCategory)
        Me.gboDetail.Controls.Add(Me.btnFindSubBrand)
        Me.gboDetail.Controls.Add(Me.lblBrand)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCategory)
        Me.gboDetail.Controls.Add(Me.txtCategoryCode)
        Me.gboDetail.Controls.Add(Me.txtPrimarySubBrand)
        Me.gboDetail.Controls.Add(Me.btnFindCategory)
        Me.gboDetail.Controls.Add(Me.txtPrimaryProductGroup)
        Me.gboDetail.Controls.Add(Me.btnFindProductGroup)
        Me.gboDetail.Controls.Add(Me.txtBrandDesc)
        Me.gboDetail.Controls.Add(Me.txtDescription)
        Me.gboDetail.Controls.Add(Me.txtCategoryDesc)
        Me.gboDetail.Controls.Add(Me.txtSubBrandDesc)
        Me.gboDetail.Controls.Add(Me.txtProductGroupDesc)
        Me.gboDetail.Location = New System.Drawing.Point(6, -1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(735, 167)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(616, 145)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(119, 13)
        Me.LinkLabel1.TabIndex = 40
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Update Monthly Closing"
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'CboGroup
        '
        Me.CboGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroup.FormattingEnabled = True
        Me.CboGroup.Items.AddRange(New Object() {"Product Code", "Category", "Sub Brand", "Brand", "Product Group"})
        Me.CboGroup.Location = New System.Drawing.Point(85, 89)
        Me.CboGroup.Margin = New System.Windows.Forms.Padding(2)
        Me.CboGroup.Name = "CboGroup"
        Me.CboGroup.Size = New System.Drawing.Size(207, 21)
        Me.CboGroup.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Group By"
        '
        'CboScope
        '
        Me.CboScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboScope.FormattingEnabled = True
        Me.CboScope.Items.AddRange(New Object() {"Whole", "Warehouse", "Location"})
        Me.CboScope.Location = New System.Drawing.Point(85, 64)
        Me.CboScope.Margin = New System.Windows.Forms.Padding(2)
        Me.CboScope.Name = "CboScope"
        Me.CboScope.Size = New System.Drawing.Size(207, 21)
        Me.CboScope.TabIndex = 38
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 68)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "Scope"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DgvLocation)
        Me.GroupBox1.Location = New System.Drawing.Point(298, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(358, 128)
        Me.GroupBox1.TabIndex = 37
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter Location | Warehouse"
        '
        'DgvLocation
        '
        Me.DgvLocation.AllowUserToResizeColumns = False
        Me.DgvLocation.AllowUserToResizeRows = False
        Me.DgvLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1No, Me.Col1Level, Me.Col1Code, Me.Col1Description, Me.Col1Primary})
        Me.DgvLocation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvLocation.Location = New System.Drawing.Point(3, 16)
        Me.DgvLocation.Name = "DgvLocation"
        Me.DgvLocation.RowHeadersWidth = 24
        Me.DgvLocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvLocation.Size = New System.Drawing.Size(352, 109)
        Me.DgvLocation.TabIndex = 36
        '
        'Col1No
        '
        Me.Col1No.HeaderText = "No"
        Me.Col1No.Name = "Col1No"
        Me.Col1No.ReadOnly = True
        Me.Col1No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col1No.Width = 27
        '
        'Col1Level
        '
        Me.Col1Level.HeaderText = "Level"
        Me.Col1Level.Items.AddRange(New Object() {"Location", "Warehouse"})
        Me.Col1Level.Name = "Col1Level"
        Me.Col1Level.Width = 39
        '
        'Col1Code
        '
        Me.Col1Code.HeaderText = "Code"
        Me.Col1Code.Name = "Col1Code"
        Me.Col1Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col1Code.Width = 38
        '
        'Col1Description
        '
        Me.Col1Description.HeaderText = "Description"
        Me.Col1Description.Name = "Col1Description"
        Me.Col1Description.ReadOnly = True
        Me.Col1Description.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col1Description.Width = 66
        '
        'Col1Primary
        '
        Me.Col1Primary.HeaderText = "Primary"
        Me.Col1Primary.Name = "Col1Primary"
        Me.Col1Primary.ReadOnly = True
        Me.Col1Primary.Visible = False
        Me.Col1Primary.Width = 66
        '
        'ChkShowItemRemarks
        '
        Me.ChkShowItemRemarks.AutoSize = True
        Me.ChkShowItemRemarks.Location = New System.Drawing.Point(502, 143)
        Me.ChkShowItemRemarks.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowItemRemarks.Name = "ChkShowItemRemarks"
        Me.ChkShowItemRemarks.Size = New System.Drawing.Size(115, 17)
        Me.ChkShowItemRemarks.TabIndex = 5
        Me.ChkShowItemRemarks.Text = "Show item remarks"
        Me.ChkShowItemRemarks.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(502, 143)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(67, 17)
        Me.CheckBox2.TabIndex = 35
        Me.CheckBox2.Text = "Show All"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'txtProductCode
        '
        Me.txtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProductCode.Location = New System.Drawing.Point(85, 115)
        Me.txtProductCode.Name = "txtProductCode"
        Me.txtProductCode.Size = New System.Drawing.Size(157, 20)
        Me.txtProductCode.TabIndex = 3
        '
        'ChkShowBatch
        '
        Me.ChkShowBatch.AutoSize = True
        Me.ChkShowBatch.Location = New System.Drawing.Point(298, 143)
        Me.ChkShowBatch.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowBatch.Name = "ChkShowBatch"
        Me.ChkShowBatch.Size = New System.Drawing.Size(196, 17)
        Me.ChkShowBatch.TabIndex = 4
        Me.ChkShowBatch.Text = "Show batch number and expiry date"
        Me.ChkShowBatch.UseVisualStyleBackColor = True
        '
        'txtProductGroupCode
        '
        Me.txtProductGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProductGroupCode.Location = New System.Drawing.Point(85, 115)
        Me.txtProductGroupCode.Name = "txtProductGroupCode"
        Me.txtProductGroupCode.Size = New System.Drawing.Size(157, 20)
        Me.txtProductGroupCode.TabIndex = 3
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(662, 40)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(66, 21)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtPrimaryProduct
        '
        Me.txtPrimaryProduct.Location = New System.Drawing.Point(85, 115)
        Me.txtPrimaryProduct.Name = "txtPrimaryProduct"
        Me.txtPrimaryProduct.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryProduct.TabIndex = 12
        Me.txtPrimaryProduct.TabStop = False
        Me.txtPrimaryProduct.Text = "-1"
        Me.txtPrimaryProduct.Visible = False
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(662, 13)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(66, 21)
        Me.btnFilter.TabIndex = 6
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'lblProductGroup
        '
        Me.lblProductGroup.AutoSize = True
        Me.lblProductGroup.Location = New System.Drawing.Point(7, 119)
        Me.lblProductGroup.Name = "lblProductGroup"
        Me.lblProductGroup.Size = New System.Drawing.Size(76, 13)
        Me.lblProductGroup.TabIndex = 30
        Me.lblProductGroup.Text = "Product Group"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "To"
        '
        'lblProductCode
        '
        Me.lblProductCode.AutoSize = True
        Me.lblProductCode.Location = New System.Drawing.Point(7, 119)
        Me.lblProductCode.Name = "lblProductCode"
        Me.lblProductCode.Size = New System.Drawing.Size(72, 13)
        Me.lblProductCode.TabIndex = 1
        Me.lblProductCode.Text = "Product Code"
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(85, 39)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(207, 20)
        Me.DtpTo.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "From"
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(85, 13)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(207, 20)
        Me.DtpFrom.TabIndex = 0
        '
        'txtPrimaryBrand
        '
        Me.txtPrimaryBrand.Location = New System.Drawing.Point(85, 115)
        Me.txtPrimaryBrand.Name = "txtPrimaryBrand"
        Me.txtPrimaryBrand.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryBrand.TabIndex = 28
        Me.txtPrimaryBrand.Text = "-1"
        Me.txtPrimaryBrand.Visible = False
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Location = New System.Drawing.Point(248, 115)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(44, 20)
        Me.btnFindProduct.TabIndex = 8
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "...."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'lblSubBrand
        '
        Me.lblSubBrand.AutoSize = True
        Me.lblSubBrand.Location = New System.Drawing.Point(7, 119)
        Me.lblSubBrand.Name = "lblSubBrand"
        Me.lblSubBrand.Size = New System.Drawing.Size(57, 13)
        Me.lblSubBrand.TabIndex = 20
        Me.lblSubBrand.Text = "Sub Brand"
        '
        'btnFindBrand
        '
        Me.btnFindBrand.Location = New System.Drawing.Point(248, 115)
        Me.btnFindBrand.Name = "btnFindBrand"
        Me.btnFindBrand.Size = New System.Drawing.Size(44, 20)
        Me.btnFindBrand.TabIndex = 11
        Me.btnFindBrand.TabStop = False
        Me.btnFindBrand.Text = "...."
        Me.btnFindBrand.UseVisualStyleBackColor = True
        '
        'txtSubBrandCode
        '
        Me.txtSubBrandCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSubBrandCode.Location = New System.Drawing.Point(85, 115)
        Me.txtSubBrandCode.Name = "txtSubBrandCode"
        Me.txtSubBrandCode.Size = New System.Drawing.Size(157, 20)
        Me.txtSubBrandCode.TabIndex = 5
        '
        'txtBrandCode
        '
        Me.txtBrandCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBrandCode.Location = New System.Drawing.Point(85, 115)
        Me.txtBrandCode.Name = "txtBrandCode"
        Me.txtBrandCode.Size = New System.Drawing.Size(157, 20)
        Me.txtBrandCode.TabIndex = 6
        '
        'lblCategory
        '
        Me.lblCategory.AutoSize = True
        Me.lblCategory.Location = New System.Drawing.Point(7, 119)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(49, 13)
        Me.lblCategory.TabIndex = 15
        Me.lblCategory.Text = "Category"
        '
        'btnFindSubBrand
        '
        Me.btnFindSubBrand.Location = New System.Drawing.Point(248, 115)
        Me.btnFindSubBrand.Name = "btnFindSubBrand"
        Me.btnFindSubBrand.Size = New System.Drawing.Size(44, 20)
        Me.btnFindSubBrand.TabIndex = 10
        Me.btnFindSubBrand.TabStop = False
        Me.btnFindSubBrand.Text = "...."
        Me.btnFindSubBrand.UseVisualStyleBackColor = True
        '
        'lblBrand
        '
        Me.lblBrand.AutoSize = True
        Me.lblBrand.Location = New System.Drawing.Point(7, 119)
        Me.lblBrand.Name = "lblBrand"
        Me.lblBrand.Size = New System.Drawing.Size(35, 13)
        Me.lblBrand.TabIndex = 25
        Me.lblBrand.Text = "Brand"
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(85, 115)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryCategory.TabIndex = 18
        Me.txtPrimaryCategory.Text = "-1"
        Me.txtPrimaryCategory.Visible = False
        '
        'txtCategoryCode
        '
        Me.txtCategoryCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCategoryCode.Location = New System.Drawing.Point(85, 115)
        Me.txtCategoryCode.Name = "txtCategoryCode"
        Me.txtCategoryCode.Size = New System.Drawing.Size(157, 20)
        Me.txtCategoryCode.TabIndex = 4
        '
        'txtPrimarySubBrand
        '
        Me.txtPrimarySubBrand.Location = New System.Drawing.Point(85, 115)
        Me.txtPrimarySubBrand.Name = "txtPrimarySubBrand"
        Me.txtPrimarySubBrand.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimarySubBrand.TabIndex = 23
        Me.txtPrimarySubBrand.Text = "-1"
        Me.txtPrimarySubBrand.Visible = False
        '
        'btnFindCategory
        '
        Me.btnFindCategory.Location = New System.Drawing.Point(248, 115)
        Me.btnFindCategory.Name = "btnFindCategory"
        Me.btnFindCategory.Size = New System.Drawing.Size(44, 20)
        Me.btnFindCategory.TabIndex = 9
        Me.btnFindCategory.TabStop = False
        Me.btnFindCategory.Text = "...."
        Me.btnFindCategory.UseVisualStyleBackColor = True
        '
        'txtPrimaryProductGroup
        '
        Me.txtPrimaryProductGroup.Location = New System.Drawing.Point(85, 115)
        Me.txtPrimaryProductGroup.Name = "txtPrimaryProductGroup"
        Me.txtPrimaryProductGroup.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryProductGroup.TabIndex = 33
        Me.txtPrimaryProductGroup.Text = "-1"
        Me.txtPrimaryProductGroup.Visible = False
        '
        'btnFindProductGroup
        '
        Me.btnFindProductGroup.Location = New System.Drawing.Point(248, 115)
        Me.btnFindProductGroup.Name = "btnFindProductGroup"
        Me.btnFindProductGroup.Size = New System.Drawing.Size(44, 20)
        Me.btnFindProductGroup.TabIndex = 10
        Me.btnFindProductGroup.TabStop = False
        Me.btnFindProductGroup.Text = "...."
        Me.btnFindProductGroup.UseVisualStyleBackColor = True
        '
        'txtBrandDesc
        '
        Me.txtBrandDesc.Location = New System.Drawing.Point(85, 141)
        Me.txtBrandDesc.Name = "txtBrandDesc"
        Me.txtBrandDesc.ReadOnly = True
        Me.txtBrandDesc.Size = New System.Drawing.Size(207, 20)
        Me.txtBrandDesc.TabIndex = 29
        Me.txtBrandDesc.TabStop = False
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(85, 141)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReadOnly = True
        Me.txtDescription.Size = New System.Drawing.Size(207, 20)
        Me.txtDescription.TabIndex = 1
        Me.txtDescription.TabStop = False
        '
        'txtCategoryDesc
        '
        Me.txtCategoryDesc.Location = New System.Drawing.Point(85, 141)
        Me.txtCategoryDesc.Name = "txtCategoryDesc"
        Me.txtCategoryDesc.ReadOnly = True
        Me.txtCategoryDesc.Size = New System.Drawing.Size(207, 20)
        Me.txtCategoryDesc.TabIndex = 19
        Me.txtCategoryDesc.TabStop = False
        '
        'txtSubBrandDesc
        '
        Me.txtSubBrandDesc.Location = New System.Drawing.Point(85, 141)
        Me.txtSubBrandDesc.Name = "txtSubBrandDesc"
        Me.txtSubBrandDesc.ReadOnly = True
        Me.txtSubBrandDesc.Size = New System.Drawing.Size(207, 20)
        Me.txtSubBrandDesc.TabIndex = 24
        Me.txtSubBrandDesc.TabStop = False
        '
        'txtProductGroupDesc
        '
        Me.txtProductGroupDesc.Location = New System.Drawing.Point(85, 141)
        Me.txtProductGroupDesc.Name = "txtProductGroupDesc"
        Me.txtProductGroupDesc.ReadOnly = True
        Me.txtProductGroupDesc.Size = New System.Drawing.Size(207, 20)
        Me.txtProductGroupDesc.TabIndex = 9
        Me.txtProductGroupDesc.TabStop = False
        '
        'crvStockCard
        '
        Me.crvStockCard.ActiveViewIndex = -1
        Me.crvStockCard.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvStockCard.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvStockCard.DisplayGroupTree = False
        Me.crvStockCard.Location = New System.Drawing.Point(6, 172)
        Me.crvStockCard.Name = "crvStockCard"
        Me.crvStockCard.SelectionFormula = ""
        Me.crvStockCard.ShowCloseButton = False
        Me.crvStockCard.ShowGroupTreeButton = False
        Me.crvStockCard.ShowRefreshButton = False
        Me.crvStockCard.Size = New System.Drawing.Size(735, 338)
        Me.crvStockCard.TabIndex = 1
        Me.crvStockCard.ViewTimeSelectionFormula = ""
        '
        'BgwLocation
        '
        '
        'LinkLabel2
        '
        Me.LinkLabel2.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(680, 128)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(55, 13)
        Me.LinkLabel2.TabIndex = 41
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Get Stock"
        Me.LinkLabel2.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTransaksiStockCard2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(746, 515)
        Me.Controls.Add(Me.crvStockCard)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(762, 554)
        Me.Name = "FrmTransaksiStockCard2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Card"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DgvLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblProductCode As System.Windows.Forms.Label
    Friend WithEvents txtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents crvStockCard As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CboGroup As System.Windows.Forms.ComboBox
    Friend WithEvents ChkShowBatch As System.Windows.Forms.CheckBox
    Friend WithEvents btnFindCategory As System.Windows.Forms.Button
    Friend WithEvents txtCategoryCode As System.Windows.Forms.TextBox
    Friend WithEvents lblCategory As System.Windows.Forms.Label
    Friend WithEvents txtCategoryDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtSubBrandDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySubBrand As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSubBrand As System.Windows.Forms.Button
    Friend WithEvents txtSubBrandCode As System.Windows.Forms.TextBox
    Friend WithEvents lblSubBrand As System.Windows.Forms.Label
    Friend WithEvents txtBrandDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryBrand As System.Windows.Forms.TextBox
    Friend WithEvents btnFindBrand As System.Windows.Forms.Button
    Friend WithEvents txtBrandCode As System.Windows.Forms.TextBox
    Friend WithEvents lblBrand As System.Windows.Forms.Label
    Friend WithEvents txtProductGroupDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryProductGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnFindProductGroup As System.Windows.Forms.Button
    Friend WithEvents txtProductGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents lblProductGroup As System.Windows.Forms.Label
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowItemRemarks As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents Col1No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Col1Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwLocation As System.ComponentModel.BackgroundWorker
    Friend WithEvents CboScope As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
End Class
