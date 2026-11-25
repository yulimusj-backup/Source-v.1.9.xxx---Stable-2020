<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseOrderBalance
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
        Me.chkPeriod = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.crvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.txtFormatDecimal = New System.Windows.Forms.TextBox
        Me.txtFormatYears = New System.Windows.Forms.TextBox
        Me.txtFormatDay = New System.Windows.Forms.TextBox
        Me.txtFormatMonth = New System.Windows.Forms.TextBox
        Me.txtSupplierName = New System.Windows.Forms.TextBox
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.TxtProductDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtProductCode = New System.Windows.Forms.TextBox
        Me.rboSupplier = New System.Windows.Forms.RadioButton
        Me.rboProduct = New System.Windows.Forms.RadioButton
        Me.ChkProduct = New System.Windows.Forms.CheckBox
        Me.btnFindProductCode1 = New System.Windows.Forms.Button
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.ChkShowRatio = New System.Windows.Forms.CheckBox
        Me.ChkIgnoreLowTolerance = New System.Windows.Forms.CheckBox
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
        Me.txtPONumber = New System.Windows.Forms.TextBox
        Me.btnFindPO = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPrimaryPO = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.TabUserKelompok.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkPeriod
        '
        Me.chkPeriod.AutoSize = True
        Me.chkPeriod.Location = New System.Drawing.Point(28, 9)
        Me.chkPeriod.Name = "chkPeriod"
        Me.chkPeriod.Size = New System.Drawing.Size(96, 17)
        Me.chkPeriod.TabIndex = 1
        Me.chkPeriod.Text = "Filter By Period"
        Me.chkPeriod.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 11)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(409, 47)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Enabled = False
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(84, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(123, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "PO Date from:"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Enabled = False
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(248, 19)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(123, 20)
        Me.dtpTo.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(220, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "to: "
        '
        'crvReport
        '
        Me.crvReport.ActiveViewIndex = -1
        Me.crvReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport.DisplayGroupTree = False
        Me.crvReport.EnableDrillDown = False
        Me.crvReport.Location = New System.Drawing.Point(16, 197)
        Me.crvReport.Name = "crvReport"
        Me.crvReport.SelectionFormula = ""
        Me.crvReport.ShowCloseButton = False
        Me.crvReport.ShowRefreshButton = False
        Me.crvReport.Size = New System.Drawing.Size(833, 356)
        Me.crvReport.TabIndex = 21
        Me.crvReport.TabStop = False
        Me.crvReport.ViewTimeSelectionFormula = ""
        '
        'txtFormatDecimal
        '
        Me.txtFormatDecimal.Location = New System.Drawing.Point(726, 180)
        Me.txtFormatDecimal.Name = "txtFormatDecimal"
        Me.txtFormatDecimal.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatDecimal.TabIndex = 26
        Me.txtFormatDecimal.TabStop = False
        Me.txtFormatDecimal.Text = "2"
        Me.txtFormatDecimal.Visible = False
        '
        'txtFormatYears
        '
        Me.txtFormatYears.Location = New System.Drawing.Point(650, 233)
        Me.txtFormatYears.Name = "txtFormatYears"
        Me.txtFormatYears.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatYears.TabIndex = 25
        Me.txtFormatYears.TabStop = False
        Me.txtFormatYears.Text = "1"
        Me.txtFormatYears.Visible = False
        '
        'txtFormatDay
        '
        Me.txtFormatDay.Location = New System.Drawing.Point(650, 207)
        Me.txtFormatDay.Name = "txtFormatDay"
        Me.txtFormatDay.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatDay.TabIndex = 24
        Me.txtFormatDay.TabStop = False
        Me.txtFormatDay.Text = "1"
        Me.txtFormatDay.Visible = False
        '
        'txtFormatMonth
        '
        Me.txtFormatMonth.Location = New System.Drawing.Point(650, 180)
        Me.txtFormatMonth.Name = "txtFormatMonth"
        Me.txtFormatMonth.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatMonth.TabIndex = 23
        Me.txtFormatMonth.TabStop = False
        Me.txtFormatMonth.Text = "3"
        Me.txtFormatMonth.Visible = False
        '
        'txtSupplierName
        '
        Me.txtSupplierName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSupplierName.Location = New System.Drawing.Point(216, 91)
        Me.txtSupplierName.MaxLength = 50
        Me.txtSupplierName.Name = "txtSupplierName"
        Me.txtSupplierName.ReadOnly = True
        Me.txtSupplierName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtSupplierName.Size = New System.Drawing.Size(209, 20)
        Me.txtSupplierName.TabIndex = 8
        Me.txtSupplierName.TabStop = False
        Me.txtSupplierName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(98, 91)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(27, 20)
        Me.txtPrimarySupplier.TabIndex = 5
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Text = "-1"
        Me.txtPrimarySupplier.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimarySupplier.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Supplier Code"
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSupplierCode.Location = New System.Drawing.Point(98, 91)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.Size = New System.Drawing.Size(74, 20)
        Me.txtSupplierCode.TabIndex = 6
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(790, 18)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(57, 21)
        Me.btnFilter.TabIndex = 19
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(790, 45)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(57, 21)
        Me.btnClose.TabIndex = 20
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'TxtProductDesc
        '
        Me.TxtProductDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtProductDesc.Location = New System.Drawing.Point(216, 117)
        Me.TxtProductDesc.MaxLength = 50
        Me.TxtProductDesc.Name = "TxtProductDesc"
        Me.TxtProductDesc.ReadOnly = True
        Me.TxtProductDesc.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TxtProductDesc.Size = New System.Drawing.Size(209, 20)
        Me.TxtProductDesc.TabIndex = 12
        Me.TxtProductDesc.TabStop = False
        Me.TxtProductDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.Cursor = System.Windows.Forms.Cursors.Default
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(98, 117)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(27, 20)
        Me.TxtPrimaryProduct.TabIndex = 9
        Me.TxtPrimaryProduct.TabStop = False
        Me.TxtPrimaryProduct.Text = "-1"
        Me.TxtPrimaryProduct.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPrimaryProduct.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 119)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Product Code"
        '
        'TxtProductCode
        '
        Me.TxtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtProductCode.Location = New System.Drawing.Point(98, 117)
        Me.TxtProductCode.MaxLength = 50
        Me.TxtProductCode.Name = "TxtProductCode"
        Me.TxtProductCode.Size = New System.Drawing.Size(74, 20)
        Me.TxtProductCode.TabIndex = 10
        '
        'rboSupplier
        '
        Me.rboSupplier.AutoSize = True
        Me.rboSupplier.Checked = True
        Me.rboSupplier.Location = New System.Drawing.Point(23, 142)
        Me.rboSupplier.Margin = New System.Windows.Forms.Padding(2)
        Me.rboSupplier.Name = "rboSupplier"
        Me.rboSupplier.Size = New System.Drawing.Size(110, 17)
        Me.rboSupplier.TabIndex = 13
        Me.rboSupplier.TabStop = True
        Me.rboSupplier.Text = "Group By Supplier"
        Me.rboSupplier.UseVisualStyleBackColor = True
        '
        'rboProduct
        '
        Me.rboProduct.AutoSize = True
        Me.rboProduct.Location = New System.Drawing.Point(23, 167)
        Me.rboProduct.Margin = New System.Windows.Forms.Padding(2)
        Me.rboProduct.Name = "rboProduct"
        Me.rboProduct.Size = New System.Drawing.Size(109, 17)
        Me.rboProduct.TabIndex = 14
        Me.rboProduct.Text = "Group By Product"
        Me.rboProduct.UseVisualStyleBackColor = True
        '
        'ChkProduct
        '
        Me.ChkProduct.AutoSize = True
        Me.ChkProduct.Location = New System.Drawing.Point(138, 143)
        Me.ChkProduct.Name = "ChkProduct"
        Me.ChkProduct.Size = New System.Drawing.Size(127, 17)
        Me.ChkProduct.TabIndex = 15
        Me.ChkProduct.Text = "Display All Item In PO"
        Me.ChkProduct.UseVisualStyleBackColor = True
        '
        'btnFindProductCode1
        '
        Me.btnFindProductCode1.Location = New System.Drawing.Point(177, 116)
        Me.btnFindProductCode1.Name = "btnFindProductCode1"
        Me.btnFindProductCode1.Size = New System.Drawing.Size(32, 20)
        Me.btnFindProductCode1.TabIndex = 11
        Me.btnFindProductCode1.TabStop = False
        Me.btnFindProductCode1.Text = "..."
        Me.btnFindProductCode1.UseVisualStyleBackColor = True
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(177, 93)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(32, 20)
        Me.btnFindSupplier.TabIndex = 7
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "..."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'ChkShowRatio
        '
        Me.ChkShowRatio.AutoSize = True
        Me.ChkShowRatio.Location = New System.Drawing.Point(283, 168)
        Me.ChkShowRatio.Name = "ChkShowRatio"
        Me.ChkShowRatio.Size = New System.Drawing.Size(142, 17)
        Me.ChkShowRatio.TabIndex = 17
        Me.ChkShowRatio.Text = "Show Ratio (If Available)"
        Me.ChkShowRatio.UseVisualStyleBackColor = True
        '
        'ChkIgnoreLowTolerance
        '
        Me.ChkIgnoreLowTolerance.AutoSize = True
        Me.ChkIgnoreLowTolerance.Location = New System.Drawing.Point(138, 167)
        Me.ChkIgnoreLowTolerance.Name = "ChkIgnoreLowTolerance"
        Me.ChkIgnoreLowTolerance.Size = New System.Drawing.Size(139, 17)
        Me.ChkIgnoreLowTolerance.TabIndex = 16
        Me.ChkIgnoreLowTolerance.Text = "Ignore Lower Tolerance"
        Me.ChkIgnoreLowTolerance.UseVisualStyleBackColor = True
        '
        'TabUserKelompok
        '
        Me.TabUserKelompok.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabUserKelompok.Controls.Add(Me.TabPage1)
        Me.TabUserKelompok.Controls.Add(Me.TabPage2)
        Me.TabUserKelompok.Location = New System.Drawing.Point(437, 18)
        Me.TabUserKelompok.Name = "TabUserKelompok"
        Me.TabUserKelompok.SelectedIndex = 0
        Me.TabUserKelompok.Size = New System.Drawing.Size(347, 173)
        Me.TabUserKelompok.TabIndex = 18
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvUser)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(339, 147)
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
        Me.DgvUser.Size = New System.Drawing.Size(335, 143)
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
        Me.TabPage2.Size = New System.Drawing.Size(339, 147)
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
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(339, 144)
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
        'txtPONumber
        '
        Me.txtPONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPONumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPONumber.Location = New System.Drawing.Point(100, 65)
        Me.txtPONumber.MaxLength = 50
        Me.txtPONumber.Name = "txtPONumber"
        Me.txtPONumber.Size = New System.Drawing.Size(287, 20)
        Me.txtPONumber.TabIndex = 3
        '
        'btnFindPO
        '
        Me.btnFindPO.Location = New System.Drawing.Point(393, 65)
        Me.btnFindPO.Name = "btnFindPO"
        Me.btnFindPO.Size = New System.Drawing.Size(32, 20)
        Me.btnFindPO.TabIndex = 4
        Me.btnFindPO.TabStop = False
        Me.btnFindPO.Text = "..."
        Me.btnFindPO.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(19, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "PO Number"
        '
        'txtPrimaryPO
        '
        Me.txtPrimaryPO.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryPO.Location = New System.Drawing.Point(97, 65)
        Me.txtPrimaryPO.Name = "txtPrimaryPO"
        Me.txtPrimaryPO.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryPO.Size = New System.Drawing.Size(27, 20)
        Me.txtPrimaryPO.TabIndex = 2
        Me.txtPrimaryPO.TabStop = False
        Me.txtPrimaryPO.Text = "-1"
        Me.txtPrimaryPO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryPO.Visible = False
        '
        'FrmLaporanPurchaseOrderBalance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(858, 571)
        Me.Controls.Add(Me.txtPrimaryPO)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnFindPO)
        Me.Controls.Add(Me.txtPONumber)
        Me.Controls.Add(Me.TabUserKelompok)
        Me.Controls.Add(Me.ChkIgnoreLowTolerance)
        Me.Controls.Add(Me.ChkShowRatio)
        Me.Controls.Add(Me.ChkProduct)
        Me.Controls.Add(Me.rboProduct)
        Me.Controls.Add(Me.rboSupplier)
        Me.Controls.Add(Me.TxtProductDesc)
        Me.Controls.Add(Me.btnFindProductCode1)
        Me.Controls.Add(Me.TxtPrimaryProduct)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtProductCode)
        Me.Controls.Add(Me.chkPeriod)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.crvReport)
        Me.Controls.Add(Me.txtFormatDecimal)
        Me.Controls.Add(Me.txtFormatYears)
        Me.Controls.Add(Me.txtFormatDay)
        Me.Controls.Add(Me.txtFormatMonth)
        Me.Controls.Add(Me.txtSupplierName)
        Me.Controls.Add(Me.btnFindSupplier)
        Me.Controls.Add(Me.txtPrimarySupplier)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSupplierCode)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.btnClose)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 550)
        Me.Name = "FrmLaporanPurchaseOrderBalance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Order Balance"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabUserKelompok.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkPeriod As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents crvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtFormatDecimal As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatYears As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatDay As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatMonth As System.Windows.Forms.TextBox
    Friend WithEvents txtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents TxtProductDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents rboSupplier As System.Windows.Forms.RadioButton
    Friend WithEvents rboProduct As System.Windows.Forms.RadioButton
    Friend WithEvents ChkProduct As System.Windows.Forms.CheckBox
    Friend WithEvents btnFindProductCode1 As System.Windows.Forms.Button
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents ChkShowRatio As System.Windows.Forms.CheckBox
    Friend WithEvents ChkIgnoreLowTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents TabUserKelompok As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents btnFindPO As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPO As System.Windows.Forms.TextBox
End Class
