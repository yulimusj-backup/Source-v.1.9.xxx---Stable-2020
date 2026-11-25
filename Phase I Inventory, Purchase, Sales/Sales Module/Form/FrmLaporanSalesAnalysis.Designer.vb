<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesAnalysis
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
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtGMPercentage = New System.Windows.Forms.TextBox
        Me.CboFilterGM = New System.Windows.Forms.ComboBox
        Me.ChkMarkGM = New System.Windows.Forms.CheckBox
        Me.chkIncludeSalesReturn = New System.Windows.Forms.CheckBox
        Me.cboDeliveryType = New System.Windows.Forms.ComboBox
        Me.chkService = New System.Windows.Forms.CheckBox
        Me.chkGoods = New System.Windows.Forms.CheckBox
        Me.pbr1 = New System.Windows.Forms.ProgressBar
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.rboProductCodeWithReturn = New System.Windows.Forms.RadioButton
        Me.rboByProductCategorywithReturn = New System.Windows.Forms.RadioButton
        Me.rboByMarketSegmentationWithReturn = New System.Windows.Forms.RadioButton
        Me.rboByCustomerTypeWithReturn = New System.Windows.Forms.RadioButton
        Me.rboBySalesAreaWithReturn = New System.Windows.Forms.RadioButton
        Me.rboByBrandWithReturn = New System.Windows.Forms.RadioButton
        Me.rboByInvoiceNumber = New System.Windows.Forms.RadioButton
        Me.rboProductCode = New System.Windows.Forms.RadioButton
        Me.rboByProductCategory = New System.Windows.Forms.RadioButton
        Me.rboSalesAreaDetailCategory = New System.Windows.Forms.RadioButton
        Me.rboSalesAreaDetail = New System.Windows.Forms.RadioButton
        Me.rboSalesAreaRekap = New System.Windows.Forms.RadioButton
        Me.rboByCustomerType = New System.Windows.Forms.RadioButton
        Me.rboByMarketSegmentation = New System.Windows.Forms.RadioButton
        Me.rboByBrand = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DgvMarketSegmentation = New System.Windows.Forms.DataGridView
        Me.ColMS_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_MarketSeg = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnShowReport = New System.Windows.Forms.Button
        Me.dtpSalesInvoicePeriod = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvBrand, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
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
        Me.TabControl1.Size = New System.Drawing.Size(880, 523)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.txtGMPercentage)
        Me.TabPage1.Controls.Add(Me.CboFilterGM)
        Me.TabPage1.Controls.Add(Me.ChkMarkGM)
        Me.TabPage1.Controls.Add(Me.chkIncludeSalesReturn)
        Me.TabPage1.Controls.Add(Me.cboDeliveryType)
        Me.TabPage1.Controls.Add(Me.chkService)
        Me.TabPage1.Controls.Add(Me.chkGoods)
        Me.TabPage1.Controls.Add(Me.pbr1)
        Me.TabPage1.Controls.Add(Me.SplitContainer3)
        Me.TabPage1.Controls.Add(Me.SplitContainer2)
        Me.TabPage1.Controls.Add(Me.SplitContainer1)
        Me.TabPage1.Controls.Add(Me.btnShowReport)
        Me.TabPage1.Controls.Add(Me.dtpSalesInvoicePeriod)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.btnExit)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(872, 497)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "CboFilterGM"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(862, 29)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(15, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "%"
        '
        'txtGMPercentage
        '
        Me.txtGMPercentage.Location = New System.Drawing.Point(815, 27)
        Me.txtGMPercentage.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtGMPercentage.Name = "txtGMPercentage"
        Me.txtGMPercentage.Size = New System.Drawing.Size(47, 20)
        Me.txtGMPercentage.TabIndex = 13
        Me.txtGMPercentage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CboFilterGM
        '
        Me.CboFilterGM.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFilterGM.FormattingEnabled = True
        Me.CboFilterGM.Items.AddRange(New Object() {"<", "<=", "=", ">=", ">"})
        Me.CboFilterGM.Location = New System.Drawing.Point(752, 26)
        Me.CboFilterGM.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.CboFilterGM.Name = "CboFilterGM"
        Me.CboFilterGM.Size = New System.Drawing.Size(56, 21)
        Me.CboFilterGM.TabIndex = 12
        '
        'ChkMarkGM
        '
        Me.ChkMarkGM.AutoSize = True
        Me.ChkMarkGM.Checked = True
        Me.ChkMarkGM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkMarkGM.Location = New System.Drawing.Point(635, 28)
        Me.ChkMarkGM.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkMarkGM.Name = "ChkMarkGM"
        Me.ChkMarkGM.Size = New System.Drawing.Size(115, 17)
        Me.ChkMarkGM.TabIndex = 11
        Me.ChkMarkGM.Text = "Mark Gross Margin"
        Me.ChkMarkGM.UseVisualStyleBackColor = True
        '
        'chkIncludeSalesReturn
        '
        Me.chkIncludeSalesReturn.AutoSize = True
        Me.chkIncludeSalesReturn.Checked = True
        Me.chkIncludeSalesReturn.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeSalesReturn.Location = New System.Drawing.Point(453, 6)
        Me.chkIncludeSalesReturn.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkIncludeSalesReturn.Name = "chkIncludeSalesReturn"
        Me.chkIncludeSalesReturn.Size = New System.Drawing.Size(96, 17)
        Me.chkIncludeSalesReturn.TabIndex = 10
        Me.chkIncludeSalesReturn.Text = "Include Return"
        Me.chkIncludeSalesReturn.UseVisualStyleBackColor = True
        '
        'cboDeliveryType
        '
        Me.cboDeliveryType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDeliveryType.FormattingEnabled = True
        Me.cboDeliveryType.Items.AddRange(New Object() {"Use Customer Sales Area", "Use Delivery Sales Area"})
        Me.cboDeliveryType.Location = New System.Drawing.Point(453, 26)
        Me.cboDeliveryType.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboDeliveryType.Name = "cboDeliveryType"
        Me.cboDeliveryType.Size = New System.Drawing.Size(159, 21)
        Me.cboDeliveryType.TabIndex = 9
        '
        'chkService
        '
        Me.chkService.AutoSize = True
        Me.chkService.Checked = True
        Me.chkService.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkService.Location = New System.Drawing.Point(391, 28)
        Me.chkService.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkService.Name = "chkService"
        Me.chkService.Size = New System.Drawing.Size(62, 17)
        Me.chkService.TabIndex = 3
        Me.chkService.Text = "Service"
        Me.chkService.UseVisualStyleBackColor = True
        '
        'chkGoods
        '
        Me.chkGoods.AutoSize = True
        Me.chkGoods.Checked = True
        Me.chkGoods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkGoods.Location = New System.Drawing.Point(390, 6)
        Me.chkGoods.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkGoods.Name = "chkGoods"
        Me.chkGoods.Size = New System.Drawing.Size(57, 17)
        Me.chkGoods.TabIndex = 2
        Me.chkGoods.Text = "Goods"
        Me.chkGoods.UseVisualStyleBackColor = True
        '
        'pbr1
        '
        Me.pbr1.Location = New System.Drawing.Point(391, 21)
        Me.pbr1.Name = "pbr1"
        Me.pbr1.Size = New System.Drawing.Size(110, 20)
        Me.pbr1.TabIndex = 8
        Me.pbr1.Visible = False
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(9, 376)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox5)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(856, 117)
        Me.SplitContainer3.SplitterDistance = 420
        Me.SplitContainer3.TabIndex = 6
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.DgvCustomer)
        Me.GroupBox5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox5.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(420, 117)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
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
        Me.DgvCustomer.Size = New System.Drawing.Size(414, 98)
        Me.DgvCustomer.TabIndex = 3
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
        Me.GroupBox1.Size = New System.Drawing.Size(432, 117)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
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
        Me.DgvBrand.Size = New System.Drawing.Size(426, 98)
        Me.DgvBrand.TabIndex = 4
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
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(8, 242)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox4)
        Me.SplitContainer2.Size = New System.Drawing.Size(856, 128)
        Me.SplitContainer2.SplitterDistance = 420
        Me.SplitContainer2.TabIndex = 5
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DgvCustomerType)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(420, 128)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
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
        Me.DgvCustomerType.Size = New System.Drawing.Size(414, 109)
        Me.DgvCustomerType.TabIndex = 4
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
        Me.GroupBox4.Size = New System.Drawing.Size(432, 128)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
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
        Me.DgvSalesArea.Size = New System.Drawing.Size(426, 109)
        Me.DgvSalesArea.TabIndex = 4
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox6)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(856, 187)
        Me.SplitContainer1.SplitterDistance = 420
        Me.SplitContainer1.TabIndex = 4
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.rboProductCodeWithReturn)
        Me.GroupBox6.Controls.Add(Me.rboByProductCategorywithReturn)
        Me.GroupBox6.Controls.Add(Me.rboByMarketSegmentationWithReturn)
        Me.GroupBox6.Controls.Add(Me.rboByCustomerTypeWithReturn)
        Me.GroupBox6.Controls.Add(Me.rboBySalesAreaWithReturn)
        Me.GroupBox6.Controls.Add(Me.rboByBrandWithReturn)
        Me.GroupBox6.Controls.Add(Me.rboByInvoiceNumber)
        Me.GroupBox6.Controls.Add(Me.rboProductCode)
        Me.GroupBox6.Controls.Add(Me.rboByProductCategory)
        Me.GroupBox6.Controls.Add(Me.rboSalesAreaDetailCategory)
        Me.GroupBox6.Controls.Add(Me.rboSalesAreaDetail)
        Me.GroupBox6.Controls.Add(Me.rboSalesAreaRekap)
        Me.GroupBox6.Controls.Add(Me.rboByCustomerType)
        Me.GroupBox6.Controls.Add(Me.rboByMarketSegmentation)
        Me.GroupBox6.Controls.Add(Me.rboByBrand)
        Me.GroupBox6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox6.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(420, 187)
        Me.GroupBox6.TabIndex = 0
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Grouping"
        '
        'rboProductCodeWithReturn
        '
        Me.rboProductCodeWithReturn.AutoSize = True
        Me.rboProductCodeWithReturn.Location = New System.Drawing.Point(232, 158)
        Me.rboProductCodeWithReturn.Name = "rboProductCodeWithReturn"
        Me.rboProductCodeWithReturn.Size = New System.Drawing.Size(156, 17)
        Me.rboProductCodeWithReturn.TabIndex = 13
        Me.rboProductCodeWithReturn.Text = "By Product Code w/ Return"
        Me.rboProductCodeWithReturn.UseVisualStyleBackColor = True
        '
        'rboByProductCategorywithReturn
        '
        Me.rboByProductCategorywithReturn.AutoSize = True
        Me.rboByProductCategorywithReturn.Location = New System.Drawing.Point(7, 42)
        Me.rboByProductCategorywithReturn.Name = "rboByProductCategorywithReturn"
        Me.rboByProductCategorywithReturn.Size = New System.Drawing.Size(173, 17)
        Me.rboByProductCategorywithReturn.TabIndex = 1
        Me.rboByProductCategorywithReturn.Text = "By Product Category w/ Return"
        Me.rboByProductCategorywithReturn.UseVisualStyleBackColor = True
        '
        'rboByMarketSegmentationWithReturn
        '
        Me.rboByMarketSegmentationWithReturn.AutoSize = True
        Me.rboByMarketSegmentationWithReturn.Location = New System.Drawing.Point(7, 88)
        Me.rboByMarketSegmentationWithReturn.Name = "rboByMarketSegmentationWithReturn"
        Me.rboByMarketSegmentationWithReturn.Size = New System.Drawing.Size(192, 17)
        Me.rboByMarketSegmentationWithReturn.TabIndex = 3
        Me.rboByMarketSegmentationWithReturn.Text = "By Market Segmentation w/ Return"
        Me.rboByMarketSegmentationWithReturn.UseVisualStyleBackColor = True
        '
        'rboByCustomerTypeWithReturn
        '
        Me.rboByCustomerTypeWithReturn.AutoSize = True
        Me.rboByCustomerTypeWithReturn.Location = New System.Drawing.Point(7, 136)
        Me.rboByCustomerTypeWithReturn.Name = "rboByCustomerTypeWithReturn"
        Me.rboByCustomerTypeWithReturn.Size = New System.Drawing.Size(162, 17)
        Me.rboByCustomerTypeWithReturn.TabIndex = 5
        Me.rboByCustomerTypeWithReturn.Text = "By Customer Type w/ Return"
        Me.rboByCustomerTypeWithReturn.UseVisualStyleBackColor = True
        '
        'rboBySalesAreaWithReturn
        '
        Me.rboBySalesAreaWithReturn.AutoSize = True
        Me.rboBySalesAreaWithReturn.Location = New System.Drawing.Point(232, 66)
        Me.rboBySalesAreaWithReturn.Name = "rboBySalesAreaWithReturn"
        Me.rboBySalesAreaWithReturn.Size = New System.Drawing.Size(142, 17)
        Me.rboBySalesAreaWithReturn.TabIndex = 9
        Me.rboBySalesAreaWithReturn.Text = "By Sales Area w/ Return"
        Me.rboBySalesAreaWithReturn.UseVisualStyleBackColor = True
        '
        'rboByBrandWithReturn
        '
        Me.rboByBrandWithReturn.AutoSize = True
        Me.rboByBrandWithReturn.Location = New System.Drawing.Point(232, 113)
        Me.rboByBrandWithReturn.Name = "rboByBrandWithReturn"
        Me.rboByBrandWithReturn.Size = New System.Drawing.Size(159, 17)
        Me.rboByBrandWithReturn.TabIndex = 11
        Me.rboByBrandWithReturn.Text = "By Product Brand w/ Return"
        Me.rboByBrandWithReturn.UseVisualStyleBackColor = True
        '
        'rboByInvoiceNumber
        '
        Me.rboByInvoiceNumber.AutoSize = True
        Me.rboByInvoiceNumber.Location = New System.Drawing.Point(232, 111)
        Me.rboByInvoiceNumber.Name = "rboByInvoiceNumber"
        Me.rboByInvoiceNumber.Size = New System.Drawing.Size(115, 17)
        Me.rboByInvoiceNumber.TabIndex = 8
        Me.rboByInvoiceNumber.Text = "By Invoice Number"
        Me.rboByInvoiceNumber.UseVisualStyleBackColor = True
        Me.rboByInvoiceNumber.Visible = False
        '
        'rboProductCode
        '
        Me.rboProductCode.AutoSize = True
        Me.rboProductCode.Location = New System.Drawing.Point(232, 135)
        Me.rboProductCode.Name = "rboProductCode"
        Me.rboProductCode.Size = New System.Drawing.Size(105, 17)
        Me.rboProductCode.TabIndex = 12
        Me.rboProductCode.Text = "By Product Code"
        Me.rboProductCode.UseVisualStyleBackColor = True
        '
        'rboByProductCategory
        '
        Me.rboByProductCategory.AutoSize = True
        Me.rboByProductCategory.Checked = True
        Me.rboByProductCategory.Location = New System.Drawing.Point(7, 19)
        Me.rboByProductCategory.Name = "rboByProductCategory"
        Me.rboByProductCategory.Size = New System.Drawing.Size(122, 17)
        Me.rboByProductCategory.TabIndex = 0
        Me.rboByProductCategory.TabStop = True
        Me.rboByProductCategory.Text = "By Product Category"
        Me.rboByProductCategory.UseVisualStyleBackColor = True
        '
        'rboSalesAreaDetailCategory
        '
        Me.rboSalesAreaDetailCategory.AutoSize = True
        Me.rboSalesAreaDetailCategory.Location = New System.Drawing.Point(232, 19)
        Me.rboSalesAreaDetailCategory.Name = "rboSalesAreaDetailCategory"
        Me.rboSalesAreaDetailCategory.Size = New System.Drawing.Size(172, 17)
        Me.rboSalesAreaDetailCategory.TabIndex = 7
        Me.rboSalesAreaDetailCategory.Text = "By Sales Area - Detail Category"
        Me.rboSalesAreaDetailCategory.UseVisualStyleBackColor = True
        '
        'rboSalesAreaDetail
        '
        Me.rboSalesAreaDetail.AutoSize = True
        Me.rboSalesAreaDetail.Location = New System.Drawing.Point(232, 42)
        Me.rboSalesAreaDetail.Name = "rboSalesAreaDetail"
        Me.rboSalesAreaDetail.Size = New System.Drawing.Size(158, 17)
        Me.rboSalesAreaDetail.TabIndex = 8
        Me.rboSalesAreaDetail.Text = "By Sales Area - Detail Brand"
        Me.rboSalesAreaDetail.UseVisualStyleBackColor = True
        '
        'rboSalesAreaRekap
        '
        Me.rboSalesAreaRekap.AutoSize = True
        Me.rboSalesAreaRekap.Location = New System.Drawing.Point(7, 158)
        Me.rboSalesAreaRekap.Name = "rboSalesAreaRekap"
        Me.rboSalesAreaRekap.Size = New System.Drawing.Size(132, 17)
        Me.rboSalesAreaRekap.TabIndex = 6
        Me.rboSalesAreaRekap.Text = "By Sales Area - Rekap"
        Me.rboSalesAreaRekap.UseVisualStyleBackColor = True
        '
        'rboByCustomerType
        '
        Me.rboByCustomerType.AutoSize = True
        Me.rboByCustomerType.Location = New System.Drawing.Point(7, 111)
        Me.rboByCustomerType.Name = "rboByCustomerType"
        Me.rboByCustomerType.Size = New System.Drawing.Size(111, 17)
        Me.rboByCustomerType.TabIndex = 4
        Me.rboByCustomerType.Text = "By Customer Type"
        Me.rboByCustomerType.UseVisualStyleBackColor = True
        '
        'rboByMarketSegmentation
        '
        Me.rboByMarketSegmentation.AutoSize = True
        Me.rboByMarketSegmentation.Location = New System.Drawing.Point(7, 66)
        Me.rboByMarketSegmentation.Name = "rboByMarketSegmentation"
        Me.rboByMarketSegmentation.Size = New System.Drawing.Size(141, 17)
        Me.rboByMarketSegmentation.TabIndex = 2
        Me.rboByMarketSegmentation.Text = "By Market Segmentation"
        Me.rboByMarketSegmentation.UseVisualStyleBackColor = True
        '
        'rboByBrand
        '
        Me.rboByBrand.AutoSize = True
        Me.rboByBrand.Location = New System.Drawing.Point(232, 89)
        Me.rboByBrand.Name = "rboByBrand"
        Me.rboByBrand.Size = New System.Drawing.Size(108, 17)
        Me.rboByBrand.TabIndex = 10
        Me.rboByBrand.Text = "By Product Brand"
        Me.rboByBrand.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DgvMarketSegmentation)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(432, 187)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
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
        Me.DgvMarketSegmentation.Size = New System.Drawing.Size(426, 168)
        Me.DgvMarketSegmentation.TabIndex = 7
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
        'btnShowReport
        '
        Me.btnShowReport.Location = New System.Drawing.Point(286, 15)
        Me.btnShowReport.Name = "btnShowReport"
        Me.btnShowReport.Size = New System.Drawing.Size(99, 30)
        Me.btnShowReport.TabIndex = 1
        Me.btnShowReport.Text = "Show Report"
        Me.btnShowReport.UseVisualStyleBackColor = True
        '
        'dtpSalesInvoicePeriod
        '
        Me.dtpSalesInvoicePeriod.CustomFormat = "MMM yyyy"
        Me.dtpSalesInvoicePeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSalesInvoicePeriod.Location = New System.Drawing.Point(80, 18)
        Me.dtpSalesInvoicePeriod.Name = "dtpSalesInvoicePeriod"
        Me.dtpSalesInvoicePeriod.ShowUpDown = True
        Me.dtpSalesInvoicePeriod.Size = New System.Drawing.Size(200, 20)
        Me.dtpSalesInvoicePeriod.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Sales Period"
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(286, 15)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(99, 30)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crv1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(878, 497)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Report View"
        Me.TabPage2.UseVisualStyleBackColor = True
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
        Me.crv1.Size = New System.Drawing.Size(872, 491)
        Me.crv1.TabIndex = 1
        Me.crv1.ViewTimeSelectionFormula = ""
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
        'FrmLaporanSalesAnalysis
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(912, 549)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimumSize = New System.Drawing.Size(928, 586)
        Me.Name = "FrmLaporanSalesAnalysis"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Analysis"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DgvBrand, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents dtpSalesInvoicePeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents rboByBrand As System.Windows.Forms.RadioButton
    Friend WithEvents rboSalesAreaDetail As System.Windows.Forms.RadioButton
    Friend WithEvents rboSalesAreaRekap As System.Windows.Forms.RadioButton
    Friend WithEvents rboByCustomerType As System.Windows.Forms.RadioButton
    Friend WithEvents rboByMarketSegmentation As System.Windows.Forms.RadioButton
    Friend WithEvents pbr1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents rboSalesAreaDetailCategory As System.Windows.Forms.RadioButton
    Friend WithEvents rboByProductCategory As System.Windows.Forms.RadioButton
    Friend WithEvents rboProductCode As System.Windows.Forms.RadioButton
    Friend WithEvents rboByInvoiceNumber As System.Windows.Forms.RadioButton
    Friend WithEvents chkService As System.Windows.Forms.CheckBox
    Friend WithEvents chkGoods As System.Windows.Forms.CheckBox
    Friend WithEvents rboByBrandWithReturn As System.Windows.Forms.RadioButton
    Friend WithEvents rboByCustomerTypeWithReturn As System.Windows.Forms.RadioButton
    Friend WithEvents rboBySalesAreaWithReturn As System.Windows.Forms.RadioButton
    Friend WithEvents cboDeliveryType As System.Windows.Forms.ComboBox
    Friend WithEvents rboByMarketSegmentationWithReturn As System.Windows.Forms.RadioButton
    Friend WithEvents rboByProductCategorywithReturn As System.Windows.Forms.RadioButton
    Friend WithEvents chkIncludeSalesReturn As System.Windows.Forms.CheckBox
    Friend WithEvents rboProductCodeWithReturn As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtGMPercentage As System.Windows.Forms.TextBox
    Friend WithEvents CboFilterGM As System.Windows.Forms.ComboBox
    Friend WithEvents ChkMarkGM As System.Windows.Forms.CheckBox
    Friend WithEvents DgvCustomerType As System.Windows.Forms.DataGridView
    Friend WithEvents ColCT_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_CustType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents ColC_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvMarketSegmentation As System.Windows.Forms.DataGridView
    Friend WithEvents ColMS_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_MarketSeg As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvSalesArea As System.Windows.Forms.DataGridView
    Friend WithEvents ColSA_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_SalesArea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvBrand As System.Windows.Forms.DataGridView
    Friend WithEvents ColB_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_Brand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColB_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
End Class
