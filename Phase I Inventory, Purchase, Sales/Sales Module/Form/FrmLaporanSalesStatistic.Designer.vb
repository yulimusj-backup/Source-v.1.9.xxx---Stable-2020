<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesStatistic
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
        Me.rboByCustGroupCode = New System.Windows.Forms.RadioButton
        Me.rboByCustCode = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DgvMarketSegmentation = New System.Windows.Forms.DataGridView
        Me.ColMS_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_MarketSeg = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnExit = New System.Windows.Forms.Button
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.btnShowReport = New System.Windows.Forms.Button
        Me.dtpSalesInvoicePeriod = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.rboSortByTotalSales = New System.Windows.Forms.RadioButton
        Me.rboSortByCode = New System.Windows.Forms.RadioButton
        Me.rboDescending = New System.Windows.Forms.RadioButton
        Me.rboAscending = New System.Windows.Forms.RadioButton
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.DgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColC_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DgvBrand = New System.Windows.Forms.DataGridView
        Me.ColB_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColB_Brand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColB_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColB_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.pbr1 = New System.Windows.Forms.ProgressBar
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.DgvCustomerType = New System.Windows.Forms.DataGridView
        Me.ColCT_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_CustType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.DgvSalesArea = New System.Windows.Forms.DataGridView
        Me.ColSA_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_SalesArea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox6.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvBrand, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rboByCustGroupCode
        '
        Me.rboByCustGroupCode.AutoSize = True
        Me.rboByCustGroupCode.Location = New System.Drawing.Point(7, 42)
        Me.rboByCustGroupCode.Name = "rboByCustGroupCode"
        Me.rboByCustGroupCode.Size = New System.Drawing.Size(144, 17)
        Me.rboByCustGroupCode.TabIndex = 1
        Me.rboByCustGroupCode.Text = "By Customer Group Code"
        Me.rboByCustGroupCode.UseVisualStyleBackColor = True
        '
        'rboByCustCode
        '
        Me.rboByCustCode.AutoSize = True
        Me.rboByCustCode.Checked = True
        Me.rboByCustCode.Location = New System.Drawing.Point(7, 19)
        Me.rboByCustCode.Name = "rboByCustCode"
        Me.rboByCustCode.Size = New System.Drawing.Size(112, 17)
        Me.rboByCustCode.TabIndex = 0
        Me.rboByCustCode.TabStop = True
        Me.rboByCustCode.Text = "By Customer Code"
        Me.rboByCustCode.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DgvMarketSegmentation)
        Me.GroupBox2.Controls.Add(Me.btnExit)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(427, 145)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Market Segmentation"
        '
        'DgvMarketSegmentation
        '
        Me.DgvMarketSegmentation.AllowUserToResizeColumns = False
        Me.DgvMarketSegmentation.AllowUserToResizeRows = False
        Me.DgvMarketSegmentation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvMarketSegmentation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvMarketSegmentation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMS_No, Me.ColMS_MarketSeg, Me.ColMS_Desc, Me.ColMS_Primary})
        Me.DgvMarketSegmentation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvMarketSegmentation.Location = New System.Drawing.Point(3, 16)
        Me.DgvMarketSegmentation.Name = "DgvMarketSegmentation"
        Me.DgvMarketSegmentation.RowHeadersWidth = 24
        Me.DgvMarketSegmentation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvMarketSegmentation.RowTemplate.Height = 24
        Me.DgvMarketSegmentation.Size = New System.Drawing.Size(421, 126)
        Me.DgvMarketSegmentation.TabIndex = 6
        '
        'ColMS_No
        '
        Me.ColMS_No.HeaderText = "No"
        Me.ColMS_No.Name = "ColMS_No"
        Me.ColMS_No.ReadOnly = True
        Me.ColMS_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMS_No.Width = 27
        '
        'ColMS_MarketSeg
        '
        Me.ColMS_MarketSeg.HeaderText = "Market Segmentation"
        Me.ColMS_MarketSeg.Name = "ColMS_MarketSeg"
        Me.ColMS_MarketSeg.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMS_MarketSeg.Width = 114
        '
        'ColMS_Desc
        '
        Me.ColMS_Desc.HeaderText = "Description"
        Me.ColMS_Desc.Name = "ColMS_Desc"
        Me.ColMS_Desc.ReadOnly = True
        Me.ColMS_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMS_Desc.Width = 66
        '
        'ColMS_Primary
        '
        Me.ColMS_Primary.HeaderText = "Primary"
        Me.ColMS_Primary.Name = "ColMS_Primary"
        Me.ColMS_Primary.Width = 66
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(73, 52)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(81, 30)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox6.Controls.Add(Me.rboByCustGroupCode)
        Me.GroupBox6.Controls.Add(Me.rboByCustCode)
        Me.GroupBox6.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(425, 63)
        Me.GroupBox6.TabIndex = 0
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Grouping"
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.DisplayStatusBar = False
        Me.crv1.DisplayToolbar = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.Location = New System.Drawing.Point(3, 3)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(903, 483)
        Me.crv1.TabIndex = 1
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crv1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(909, 489)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Report View"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 528)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(904, 22)
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnShowReport
        '
        Me.btnShowReport.Location = New System.Drawing.Point(286, 13)
        Me.btnShowReport.Name = "btnShowReport"
        Me.btnShowReport.Size = New System.Drawing.Size(81, 30)
        Me.btnShowReport.TabIndex = 1
        Me.btnShowReport.Text = "Show Report"
        Me.btnShowReport.UseVisualStyleBackColor = True
        '
        'dtpSalesInvoicePeriod
        '
        Me.dtpSalesInvoicePeriod.CustomFormat = "MMM yyyy"
        Me.dtpSalesInvoicePeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSalesInvoicePeriod.Location = New System.Drawing.Point(80, 13)
        Me.dtpSalesInvoicePeriod.Name = "dtpSalesInvoicePeriod"
        Me.dtpSalesInvoicePeriod.ShowUpDown = True
        Me.dtpSalesInvoicePeriod.Size = New System.Drawing.Size(200, 20)
        Me.dtpSalesInvoicePeriod.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Sales Period"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 49)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox6)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(856, 145)
        Me.SplitContainer1.SplitterDistance = 425
        Me.SplitContainer1.TabIndex = 5
        '
        'GroupBox7
        '
        Me.GroupBox7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox7.Controls.Add(Me.SplitContainer4)
        Me.GroupBox7.Location = New System.Drawing.Point(0, 69)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(425, 63)
        Me.GroupBox7.TabIndex = 0
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Sort"
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.rboSortByTotalSales)
        Me.SplitContainer4.Panel1.Controls.Add(Me.rboSortByCode)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.rboDescending)
        Me.SplitContainer4.Panel2.Controls.Add(Me.rboAscending)
        Me.SplitContainer4.Size = New System.Drawing.Size(419, 44)
        Me.SplitContainer4.SplitterDistance = 204
        Me.SplitContainer4.TabIndex = 0
        '
        'rboSortByTotalSales
        '
        Me.rboSortByTotalSales.AutoSize = True
        Me.rboSortByTotalSales.Location = New System.Drawing.Point(4, 25)
        Me.rboSortByTotalSales.Name = "rboSortByTotalSales"
        Me.rboSortByTotalSales.Size = New System.Drawing.Size(78, 17)
        Me.rboSortByTotalSales.TabIndex = 3
        Me.rboSortByTotalSales.Text = "Total Sales"
        Me.rboSortByTotalSales.UseVisualStyleBackColor = True
        '
        'rboSortByCode
        '
        Me.rboSortByCode.AutoSize = True
        Me.rboSortByCode.Checked = True
        Me.rboSortByCode.Location = New System.Drawing.Point(4, 2)
        Me.rboSortByCode.Name = "rboSortByCode"
        Me.rboSortByCode.Size = New System.Drawing.Size(112, 17)
        Me.rboSortByCode.TabIndex = 2
        Me.rboSortByCode.TabStop = True
        Me.rboSortByCode.Text = "By Customer Code"
        Me.rboSortByCode.UseVisualStyleBackColor = True
        '
        'rboDescending
        '
        Me.rboDescending.AutoSize = True
        Me.rboDescending.Location = New System.Drawing.Point(11, 25)
        Me.rboDescending.Name = "rboDescending"
        Me.rboDescending.Size = New System.Drawing.Size(82, 17)
        Me.rboDescending.TabIndex = 3
        Me.rboDescending.Text = "Descending"
        Me.rboDescending.UseVisualStyleBackColor = True
        '
        'rboAscending
        '
        Me.rboAscending.AutoSize = True
        Me.rboAscending.Checked = True
        Me.rboAscending.Location = New System.Drawing.Point(11, 2)
        Me.rboAscending.Name = "rboAscending"
        Me.rboAscending.Size = New System.Drawing.Size(75, 17)
        Me.rboAscending.TabIndex = 2
        Me.rboAscending.TabStop = True
        Me.rboAscending.Text = "Ascending"
        Me.rboAscending.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.DgvCustomer)
        Me.GroupBox5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox5.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(425, 106)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Customer"
        '
        'DgvCustomer
        '
        Me.DgvCustomer.AllowUserToResizeColumns = False
        Me.DgvCustomer.AllowUserToResizeRows = False
        Me.DgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColC_No, Me.ColC_Code, Me.ColC_Desc, Me.ColC_Primary})
        Me.DgvCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomer.Location = New System.Drawing.Point(3, 16)
        Me.DgvCustomer.Name = "DgvCustomer"
        Me.DgvCustomer.RowHeadersWidth = 24
        Me.DgvCustomer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomer.RowTemplate.Height = 24
        Me.DgvCustomer.Size = New System.Drawing.Size(419, 87)
        Me.DgvCustomer.TabIndex = 2
        '
        'ColC_No
        '
        Me.ColC_No.HeaderText = "No"
        Me.ColC_No.Name = "ColC_No"
        Me.ColC_No.ReadOnly = True
        Me.ColC_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColC_No.Width = 27
        '
        'ColC_Code
        '
        Me.ColC_Code.HeaderText = "Customer Code"
        Me.ColC_Code.Name = "ColC_Code"
        Me.ColC_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColC_Code.Width = 85
        '
        'ColC_Desc
        '
        Me.ColC_Desc.HeaderText = "Customer Name"
        Me.ColC_Desc.Name = "ColC_Desc"
        Me.ColC_Desc.ReadOnly = True
        Me.ColC_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColC_Desc.Width = 88
        '
        'ColC_Primary
        '
        Me.ColC_Primary.HeaderText = "Primary"
        Me.ColC_Primary.Name = "ColC_Primary"
        Me.ColC_Primary.Visible = False
        Me.ColC_Primary.Width = 81
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DgvBrand)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(427, 106)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Brand"
        '
        'DgvBrand
        '
        Me.DgvBrand.AllowUserToResizeColumns = False
        Me.DgvBrand.AllowUserToResizeRows = False
        Me.DgvBrand.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvBrand.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvBrand.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColB_No, Me.ColB_Brand, Me.ColB_Desc, Me.ColB_Primary})
        Me.DgvBrand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvBrand.Location = New System.Drawing.Point(3, 16)
        Me.DgvBrand.Name = "DgvBrand"
        Me.DgvBrand.RowHeadersWidth = 24
        Me.DgvBrand.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvBrand.RowTemplate.Height = 24
        Me.DgvBrand.Size = New System.Drawing.Size(421, 87)
        Me.DgvBrand.TabIndex = 3
        '
        'ColB_No
        '
        Me.ColB_No.HeaderText = "No"
        Me.ColB_No.Name = "ColB_No"
        Me.ColB_No.ReadOnly = True
        Me.ColB_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColB_No.Width = 27
        '
        'ColB_Brand
        '
        Me.ColB_Brand.HeaderText = "Brand"
        Me.ColB_Brand.Name = "ColB_Brand"
        Me.ColB_Brand.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColB_Brand.Width = 41
        '
        'ColB_Desc
        '
        Me.ColB_Desc.HeaderText = "Description"
        Me.ColB_Desc.Name = "ColB_Desc"
        Me.ColB_Desc.ReadOnly = True
        Me.ColB_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColB_Desc.Width = 66
        '
        'ColB_Primary
        '
        Me.ColB_Primary.HeaderText = "Primary"
        Me.ColB_Primary.Name = "ColB_Primary"
        Me.ColB_Primary.ReadOnly = True
        Me.ColB_Primary.Visible = False
        Me.ColB_Primary.Width = 81
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(8, 351)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox5)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(856, 106)
        Me.SplitContainer3.SplitterDistance = 425
        Me.SplitContainer3.TabIndex = 7
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(880, 513)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.pbr1)
        Me.TabPage1.Controls.Add(Me.SplitContainer3)
        Me.TabPage1.Controls.Add(Me.SplitContainer2)
        Me.TabPage1.Controls.Add(Me.SplitContainer1)
        Me.TabPage1.Controls.Add(Me.btnShowReport)
        Me.TabPage1.Controls.Add(Me.dtpSalesInvoicePeriod)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(872, 487)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Filter && Grouping"
        '
        'pbr1
        '
        Me.pbr1.Location = New System.Drawing.Point(373, 18)
        Me.pbr1.Name = "pbr1"
        Me.pbr1.Size = New System.Drawing.Size(110, 20)
        Me.pbr1.TabIndex = 8
        Me.pbr1.Visible = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(8, 200)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox4)
        Me.SplitContainer2.Size = New System.Drawing.Size(856, 145)
        Me.SplitContainer2.SplitterDistance = 425
        Me.SplitContainer2.TabIndex = 6
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DgvCustomerType)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(425, 145)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Customer Type"
        '
        'DgvCustomerType
        '
        Me.DgvCustomerType.AllowUserToResizeColumns = False
        Me.DgvCustomerType.AllowUserToResizeRows = False
        Me.DgvCustomerType.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomerType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomerType.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCT_No, Me.ColCT_CustType, Me.ColCT_Desc, Me.ColCT_Primary})
        Me.DgvCustomerType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomerType.Location = New System.Drawing.Point(3, 16)
        Me.DgvCustomerType.Name = "DgvCustomerType"
        Me.DgvCustomerType.RowHeadersWidth = 24
        Me.DgvCustomerType.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomerType.RowTemplate.Height = 24
        Me.DgvCustomerType.Size = New System.Drawing.Size(419, 126)
        Me.DgvCustomerType.TabIndex = 3
        '
        'ColCT_No
        '
        Me.ColCT_No.HeaderText = "No"
        Me.ColCT_No.Name = "ColCT_No"
        Me.ColCT_No.ReadOnly = True
        Me.ColCT_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_No.Width = 27
        '
        'ColCT_CustType
        '
        Me.ColCT_CustType.HeaderText = "Customer Type"
        Me.ColCT_CustType.Name = "ColCT_CustType"
        Me.ColCT_CustType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_CustType.Width = 84
        '
        'ColCT_Desc
        '
        Me.ColCT_Desc.HeaderText = "Description"
        Me.ColCT_Desc.Name = "ColCT_Desc"
        Me.ColCT_Desc.ReadOnly = True
        Me.ColCT_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_Desc.Width = 66
        '
        'ColCT_Primary
        '
        Me.ColCT_Primary.HeaderText = "Primary"
        Me.ColCT_Primary.Name = "ColCT_Primary"
        Me.ColCT_Primary.ReadOnly = True
        Me.ColCT_Primary.Visible = False
        Me.ColCT_Primary.Width = 81
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.DgvSalesArea)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(427, 145)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Sales Area"
        '
        'DgvSalesArea
        '
        Me.DgvSalesArea.AllowUserToResizeColumns = False
        Me.DgvSalesArea.AllowUserToResizeRows = False
        Me.DgvSalesArea.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalesArea.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSalesArea.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSA_No, Me.ColSA_SalesArea, Me.ColSA_Desc, Me.ColSA_Primary})
        Me.DgvSalesArea.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalesArea.Location = New System.Drawing.Point(3, 16)
        Me.DgvSalesArea.Name = "DgvSalesArea"
        Me.DgvSalesArea.RowHeadersWidth = 24
        Me.DgvSalesArea.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSalesArea.RowTemplate.Height = 24
        Me.DgvSalesArea.Size = New System.Drawing.Size(421, 126)
        Me.DgvSalesArea.TabIndex = 3
        '
        'ColSA_No
        '
        Me.ColSA_No.HeaderText = "No"
        Me.ColSA_No.Name = "ColSA_No"
        Me.ColSA_No.ReadOnly = True
        Me.ColSA_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSA_No.Width = 27
        '
        'ColSA_SalesArea
        '
        Me.ColSA_SalesArea.HeaderText = "Sales Area"
        Me.ColSA_SalesArea.Name = "ColSA_SalesArea"
        Me.ColSA_SalesArea.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSA_SalesArea.Width = 64
        '
        'ColSA_Desc
        '
        Me.ColSA_Desc.HeaderText = "Description"
        Me.ColSA_Desc.Name = "ColSA_Desc"
        Me.ColSA_Desc.ReadOnly = True
        Me.ColSA_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSA_Desc.Width = 66
        '
        'ColSA_Primary
        '
        Me.ColSA_Primary.HeaderText = "Primary"
        Me.ColSA_Primary.Name = "ColSA_Primary"
        Me.ColSA_Primary.ReadOnly = True
        Me.ColSA_Primary.Visible = False
        Me.ColSA_Primary.Width = 81
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker3
        '
        '
        'BackgroundWorker4
        '
        '
        'BackgroundWorker5
        '
        '
        'FrmLaporanSalesStatistic
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(904, 550)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(920, 588)
        Me.Name = "FrmLaporanSalesStatistic"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Statistic"
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DgvBrand, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rboByCustGroupCode As System.Windows.Forms.RadioButton
    Friend WithEvents rboByCustCode As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents dtpSalesInvoicePeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents pbr1 As System.Windows.Forms.ProgressBar
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents rboSortByTotalSales As System.Windows.Forms.RadioButton
    Friend WithEvents rboSortByCode As System.Windows.Forms.RadioButton
    Friend WithEvents rboDescending As System.Windows.Forms.RadioButton
    Friend WithEvents rboAscending As System.Windows.Forms.RadioButton
    Friend WithEvents DgvMarketSegmentation As System.Windows.Forms.DataGridView
    Friend WithEvents DgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents DgvBrand As System.Windows.Forms.DataGridView
    Friend WithEvents DgvCustomerType As System.Windows.Forms.DataGridView
    Friend WithEvents DgvSalesArea As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColMS_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_MarketSeg As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_Brand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_CustType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_SalesArea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
