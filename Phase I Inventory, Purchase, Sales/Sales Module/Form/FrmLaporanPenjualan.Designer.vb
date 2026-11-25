<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPenjualan
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
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.chkCustomer = New System.Windows.Forms.CheckBox
        Me.clbCustomer = New System.Windows.Forms.CheckedListBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkProduct = New System.Windows.Forms.CheckBox
        Me.clbProduct = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboitemDetail = New System.Windows.Forms.ComboBox
        Me.ChkProductDesc = New System.Windows.Forms.CheckBox
        Me.cboSort1 = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.cboGroupBy = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.btnShowReport = New System.Windows.Forms.Button
        Me.dtpTo2 = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom2 = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvSortRekap = New System.Windows.Forms.DataGridView
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSortNumber = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColSrtType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.dgvSortDetail = New System.Windows.Forms.DataGridView
        Me.ColName2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewComboBoxColumn1 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.DataGridViewComboBoxColumn2 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.dgvSortRekap2PeriodComp = New System.Windows.Forms.DataGridView
        Me.ColName1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewComboBoxColumn3 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.DataGridViewComboBoxColumn4 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SplitContainer8 = New System.Windows.Forms.SplitContainer
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvMarketSegmentation = New System.Windows.Forms.DataGridView
        Me.ColMS_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_MarketSeg = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.DgvSalesArea = New System.Windows.Forms.DataGridView
        Me.ColSA_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_SalesArea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.DgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColC_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.dgvCustomerGroup = New System.Windows.Forms.DataGridView
        Me.ColSG_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSG_SalesGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSG_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSG_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.DgvCustomerType = New System.Windows.Forms.DataGridView
        Me.ColCT_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_CustType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label8 = New System.Windows.Forms.Label
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BgwProduct = New System.ComponentModel.BackgroundWorker
        Me.bgwCustomer = New System.ComponentModel.BackgroundWorker
        Me.bgwCustSegmen = New System.ComponentModel.BackgroundWorker
        Me.bgwSalesArea = New System.ComponentModel.BackgroundWorker
        Me.bgwCustType = New System.ComponentModel.BackgroundWorker
        Me.bgwCustGroup = New System.ComponentModel.BackgroundWorker
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvSortRekap, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSortDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSortRekap2PeriodComp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer8.Panel1.SuspendLayout()
        Me.SplitContainer8.Panel2.SuspendLayout()
        Me.SplitContainer8.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage8.SuspendLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgvCustomerGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(141, 331)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(776, 499)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.SplitContainer3)
        Me.TabPage1.Controls.Add(Me.SplitContainer1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(768, 473)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Filter && Grouping"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(859, 472)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox5)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(752, 267)
        Me.SplitContainer3.SplitterDistance = 370
        Me.SplitContainer3.TabIndex = 7
        Me.SplitContainer3.Visible = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.chkCustomer)
        Me.GroupBox5.Controls.Add(Me.clbCustomer)
        Me.GroupBox5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox5.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(370, 267)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        '
        'chkCustomer
        '
        Me.chkCustomer.AutoSize = True
        Me.chkCustomer.Location = New System.Drawing.Point(7, -1)
        Me.chkCustomer.Name = "chkCustomer"
        Me.chkCustomer.Size = New System.Drawing.Size(84, 17)
        Me.chkCustomer.TabIndex = 0
        Me.chkCustomer.Text = "All Customer"
        Me.chkCustomer.UseVisualStyleBackColor = True
        '
        'clbCustomer
        '
        Me.clbCustomer.CheckOnClick = True
        Me.clbCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbCustomer.FormattingEnabled = True
        Me.clbCustomer.Location = New System.Drawing.Point(3, 16)
        Me.clbCustomer.Name = "clbCustomer"
        Me.clbCustomer.Size = New System.Drawing.Size(364, 244)
        Me.clbCustomer.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkProduct)
        Me.GroupBox1.Controls.Add(Me.clbProduct)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(378, 267)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'chkProduct
        '
        Me.chkProduct.AutoSize = True
        Me.chkProduct.Location = New System.Drawing.Point(7, -1)
        Me.chkProduct.Name = "chkProduct"
        Me.chkProduct.Size = New System.Drawing.Size(77, 17)
        Me.chkProduct.TabIndex = 0
        Me.chkProduct.Text = "All Product"
        Me.chkProduct.UseVisualStyleBackColor = True
        '
        'clbProduct
        '
        Me.clbProduct.CheckOnClick = True
        Me.clbProduct.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbProduct.FormattingEnabled = True
        Me.clbProduct.Location = New System.Drawing.Point(3, 16)
        Me.clbProduct.Name = "clbProduct"
        Me.clbProduct.Size = New System.Drawing.Size(372, 244)
        Me.clbProduct.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 6)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Size = New System.Drawing.Size(755, 174)
        Me.SplitContainer1.SplitterDistance = 352
        Me.SplitContainer1.TabIndex = 18
        Me.SplitContainer1.TabStop = False
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(847, 475)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Report View"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 13)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Item Detail"
        '
        'cboitemDetail
        '
        Me.cboitemDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboitemDetail.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboitemDetail.FormattingEnabled = True
        Me.cboitemDetail.Items.AddRange(New Object() {"Code No", "Category", "Sub Brand", "Brand", "Product Group"})
        Me.cboitemDetail.Location = New System.Drawing.Point(66, 84)
        Me.cboitemDetail.Name = "cboitemDetail"
        Me.cboitemDetail.Size = New System.Drawing.Size(244, 21)
        Me.cboitemDetail.TabIndex = 20
        '
        'ChkProductDesc
        '
        Me.ChkProductDesc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkProductDesc.AutoSize = True
        Me.ChkProductDesc.Location = New System.Drawing.Point(7, 510)
        Me.ChkProductDesc.Name = "ChkProductDesc"
        Me.ChkProductDesc.Size = New System.Drawing.Size(149, 17)
        Me.ChkProductDesc.TabIndex = 19
        Me.ChkProductDesc.Text = "Show Product Description"
        Me.ChkProductDesc.UseVisualStyleBackColor = True
        '
        'cboSort1
        '
        Me.cboSort1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSort1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSort1.FormattingEnabled = True
        Me.cboSort1.Items.AddRange(New Object() {"Ascending", "Descending", "Qty - Ascending", "Qty - Descending", "Amount - Ascending", "Amount - Descending"})
        Me.cboSort1.Location = New System.Drawing.Point(66, 57)
        Me.cboSort1.Name = "cboSort1"
        Me.cboSort1.Size = New System.Drawing.Size(244, 21)
        Me.cboSort1.TabIndex = 18
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 34)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 13)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Group By"
        '
        'cboFormat
        '
        Me.cboFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Recap", "Recap-2 Period Comparison", "Detail"})
        Me.cboFormat.Location = New System.Drawing.Point(66, 3)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(244, 21)
        Me.cboFormat.TabIndex = 14
        '
        'cboGroupBy
        '
        Me.cboGroupBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboGroupBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupBy.FormattingEnabled = True
        Me.cboGroupBy.Items.AddRange(New Object() {"Rekap", "Detail"})
        Me.cboGroupBy.Location = New System.Drawing.Point(66, 30)
        Me.cboGroupBy.Name = "cboGroupBy"
        Me.cboGroupBy.Size = New System.Drawing.Size(244, 21)
        Me.cboGroupBy.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 115)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Period I"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 7)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Format"
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(66, 111)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(244, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'btnShowReport
        '
        Me.btnShowReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShowReport.Location = New System.Drawing.Point(216, 510)
        Me.btnShowReport.Name = "btnShowReport"
        Me.btnShowReport.Size = New System.Drawing.Size(89, 21)
        Me.btnShowReport.TabIndex = 1
        Me.btnShowReport.Text = "Show Report"
        Me.btnShowReport.UseVisualStyleBackColor = True
        '
        'dtpTo2
        '
        Me.dtpTo2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo2.CustomFormat = "dd MMM yyyy"
        Me.dtpTo2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo2.Location = New System.Drawing.Point(66, 189)
        Me.dtpTo2.Name = "dtpTo2"
        Me.dtpTo2.Size = New System.Drawing.Size(244, 20)
        Me.dtpTo2.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 141)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "to"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 193)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "to"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(66, 137)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(244, 20)
        Me.dtpTo.TabIndex = 9
        '
        'dtpFrom2
        '
        Me.dtpFrom2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom2.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom2.Location = New System.Drawing.Point(66, 163)
        Me.dtpFrom2.Name = "dtpFrom2"
        Me.dtpFrom2.Size = New System.Drawing.Size(244, 20)
        Me.dtpFrom2.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 167)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Period II"
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(8, 63)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(99, 30)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvSortRekap)
        Me.GroupBox2.Controls.Add(Me.dgvSortDetail)
        Me.GroupBox2.Controls.Add(Me.btnExit)
        Me.GroupBox2.Controls.Add(Me.dgvSortRekap2PeriodComp)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 215)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(304, 159)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Item Detail Sort"
        '
        'dgvSortRekap
        '
        Me.dgvSortRekap.AllowUserToAddRows = False
        Me.dgvSortRekap.AllowUserToDeleteRows = False
        Me.dgvSortRekap.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSortRekap.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSortRekap.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColName, Me.ColSortNumber, Me.ColSrtType})
        Me.dgvSortRekap.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSortRekap.Location = New System.Drawing.Point(3, 16)
        Me.dgvSortRekap.Name = "dgvSortRekap"
        Me.dgvSortRekap.RowHeadersWidth = 24
        Me.dgvSortRekap.RowTemplate.Height = 24
        Me.dgvSortRekap.Size = New System.Drawing.Size(298, 140)
        Me.dgvSortRekap.TabIndex = 0
        '
        'ColName
        '
        Me.ColName.HeaderText = "Column Name"
        Me.ColName.Name = "ColName"
        Me.ColName.Width = 98
        '
        'ColSortNumber
        '
        Me.ColSortNumber.HeaderText = "Sort Number"
        Me.ColSortNumber.Items.AddRange(New Object() {"1", "2", "3"})
        Me.ColSortNumber.Name = "ColSortNumber"
        Me.ColSortNumber.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColSortNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColSortNumber.Width = 91
        '
        'ColSrtType
        '
        Me.ColSrtType.HeaderText = "Sort Type"
        Me.ColSrtType.Items.AddRange(New Object() {"Ascending", "Descending"})
        Me.ColSrtType.Name = "ColSrtType"
        Me.ColSrtType.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColSrtType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColSrtType.Width = 78
        '
        'dgvSortDetail
        '
        Me.dgvSortDetail.AllowUserToAddRows = False
        Me.dgvSortDetail.AllowUserToDeleteRows = False
        Me.dgvSortDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSortDetail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSortDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColName2, Me.DataGridViewComboBoxColumn1, Me.DataGridViewComboBoxColumn2})
        Me.dgvSortDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSortDetail.Location = New System.Drawing.Point(3, 16)
        Me.dgvSortDetail.Name = "dgvSortDetail"
        Me.dgvSortDetail.RowHeadersWidth = 24
        Me.dgvSortDetail.RowTemplate.Height = 24
        Me.dgvSortDetail.Size = New System.Drawing.Size(298, 140)
        Me.dgvSortDetail.TabIndex = 1
        '
        'ColName2
        '
        Me.ColName2.HeaderText = "Column Name"
        Me.ColName2.Name = "ColName2"
        Me.ColName2.Width = 98
        '
        'DataGridViewComboBoxColumn1
        '
        Me.DataGridViewComboBoxColumn1.HeaderText = "Sort Number"
        Me.DataGridViewComboBoxColumn1.Items.AddRange(New Object() {"1", "2", "3", "4", "5"})
        Me.DataGridViewComboBoxColumn1.Name = "DataGridViewComboBoxColumn1"
        Me.DataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewComboBoxColumn1.Width = 91
        '
        'DataGridViewComboBoxColumn2
        '
        Me.DataGridViewComboBoxColumn2.HeaderText = "Sort Type"
        Me.DataGridViewComboBoxColumn2.Items.AddRange(New Object() {"Ascending", "Descending"})
        Me.DataGridViewComboBoxColumn2.Name = "DataGridViewComboBoxColumn2"
        Me.DataGridViewComboBoxColumn2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewComboBoxColumn2.Width = 78
        '
        'dgvSortRekap2PeriodComp
        '
        Me.dgvSortRekap2PeriodComp.AllowUserToAddRows = False
        Me.dgvSortRekap2PeriodComp.AllowUserToDeleteRows = False
        Me.dgvSortRekap2PeriodComp.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSortRekap2PeriodComp.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSortRekap2PeriodComp.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColName1, Me.DataGridViewComboBoxColumn3, Me.DataGridViewComboBoxColumn4})
        Me.dgvSortRekap2PeriodComp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSortRekap2PeriodComp.Location = New System.Drawing.Point(3, 16)
        Me.dgvSortRekap2PeriodComp.Name = "dgvSortRekap2PeriodComp"
        Me.dgvSortRekap2PeriodComp.RowHeadersWidth = 24
        Me.dgvSortRekap2PeriodComp.RowTemplate.Height = 24
        Me.dgvSortRekap2PeriodComp.Size = New System.Drawing.Size(298, 140)
        Me.dgvSortRekap2PeriodComp.TabIndex = 2
        '
        'ColName1
        '
        Me.ColName1.HeaderText = "Column Name"
        Me.ColName1.Name = "ColName1"
        Me.ColName1.Width = 98
        '
        'DataGridViewComboBoxColumn3
        '
        Me.DataGridViewComboBoxColumn3.HeaderText = "Sort Number"
        Me.DataGridViewComboBoxColumn3.Items.AddRange(New Object() {"1", "2", "3", "4"})
        Me.DataGridViewComboBoxColumn3.Name = "DataGridViewComboBoxColumn3"
        Me.DataGridViewComboBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewComboBoxColumn3.Width = 91
        '
        'DataGridViewComboBoxColumn4
        '
        Me.DataGridViewComboBoxColumn4.HeaderText = "Sort Type"
        Me.DataGridViewComboBoxColumn4.Items.AddRange(New Object() {"Ascending", "Descending"})
        Me.DataGridViewComboBoxColumn4.Name = "DataGridViewComboBoxColumn4"
        Me.DataGridViewComboBoxColumn4.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewComboBoxColumn4.Width = 78
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 537)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(902, 22)
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'SplitContainer8
        '
        Me.SplitContainer8.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer8.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer8.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer8.Name = "SplitContainer8"
        '
        'SplitContainer8.Panel1
        '
        Me.SplitContainer8.Panel1.Controls.Add(Me.btnShowReport)
        Me.SplitContainer8.Panel1.Controls.Add(Me.TabControl2)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer8.Panel1.Controls.Add(Me.dtpTo2)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer8.Panel1.Controls.Add(Me.dtpFrom2)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer8.Panel1.Controls.Add(Me.dtpFrom)
        Me.SplitContainer8.Panel1.Controls.Add(Me.cboitemDetail)
        Me.SplitContainer8.Panel1.Controls.Add(Me.dtpTo)
        Me.SplitContainer8.Panel1.Controls.Add(Me.cboFormat)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer8.Panel1.Controls.Add(Me.cboGroupBy)
        Me.SplitContainer8.Panel1.Controls.Add(Me.cboSort1)
        Me.SplitContainer8.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer8.Panel1.Controls.Add(Me.ChkProductDesc)
        Me.SplitContainer8.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer8.Panel2
        '
        Me.SplitContainer8.Panel2.Controls.Add(Me.crv1)
        Me.SplitContainer8.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer8.Size = New System.Drawing.Size(902, 534)
        Me.SplitContainer8.SplitterDistance = 314
        Me.SplitContainer8.TabIndex = 4
        Me.SplitContainer8.TabStop = False
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage8)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Location = New System.Drawing.Point(4, 380)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(304, 124)
        Me.TabControl2.TabIndex = 3
        '
        'TabPage8
        '
        Me.TabPage8.BackColor = System.Drawing.Color.Transparent
        Me.TabPage8.Controls.Add(Me.dgvProduct)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage8.Size = New System.Drawing.Size(296, 98)
        Me.TabPage8.TabIndex = 5
        Me.TabPage8.Text = "Product"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'dgvProduct
        '
        Me.dgvProduct.AllowUserToResizeColumns = False
        Me.dgvProduct.AllowUserToResizeRows = False
        Me.dgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProduct_No, Me.colProduct_GroupLevel, Me.colProduct_Code, Me.colProduct_Desc, Me.colProduct_Primary})
        Me.dgvProduct.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProduct.Location = New System.Drawing.Point(3, 3)
        Me.dgvProduct.Name = "dgvProduct"
        Me.dgvProduct.RowHeadersWidth = 24
        Me.dgvProduct.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvProduct.RowTemplate.Height = 24
        Me.dgvProduct.Size = New System.Drawing.Size(290, 92)
        Me.dgvProduct.TabIndex = 5
        '
        'colProduct_No
        '
        Me.colProduct_No.HeaderText = "No"
        Me.colProduct_No.Name = "colProduct_No"
        Me.colProduct_No.ReadOnly = True
        Me.colProduct_No.Width = 46
        '
        'colProduct_GroupLevel
        '
        Me.colProduct_GroupLevel.HeaderText = "Level"
        Me.colProduct_GroupLevel.Items.AddRange(New Object() {"Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.colProduct_GroupLevel.Name = "colProduct_GroupLevel"
        Me.colProduct_GroupLevel.Width = 39
        '
        'colProduct_Code
        '
        Me.colProduct_Code.HeaderText = "Code"
        Me.colProduct_Code.Name = "colProduct_Code"
        Me.colProduct_Code.Width = 57
        '
        'colProduct_Desc
        '
        Me.colProduct_Desc.HeaderText = "Description"
        Me.colProduct_Desc.Name = "colProduct_Desc"
        Me.colProduct_Desc.ReadOnly = True
        Me.colProduct_Desc.Width = 85
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.ReadOnly = True
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 81
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.Transparent
        Me.TabPage3.Controls.Add(Me.DgvMarketSegmentation)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage3.Size = New System.Drawing.Size(218, 100)
        Me.TabPage3.TabIndex = 0
        Me.TabPage3.Text = "Segmentation"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DgvMarketSegmentation
        '
        Me.DgvMarketSegmentation.AllowUserToResizeColumns = False
        Me.DgvMarketSegmentation.AllowUserToResizeRows = False
        Me.DgvMarketSegmentation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvMarketSegmentation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvMarketSegmentation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMS_No, Me.ColMS_MarketSeg, Me.ColMS_Desc, Me.ColMS_Primary})
        Me.DgvMarketSegmentation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvMarketSegmentation.Location = New System.Drawing.Point(3, 3)
        Me.DgvMarketSegmentation.Name = "DgvMarketSegmentation"
        Me.DgvMarketSegmentation.RowHeadersWidth = 24
        Me.DgvMarketSegmentation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvMarketSegmentation.RowTemplate.Height = 24
        Me.DgvMarketSegmentation.Size = New System.Drawing.Size(212, 94)
        Me.DgvMarketSegmentation.TabIndex = 5
        '
        'ColMS_No
        '
        Me.ColMS_No.HeaderText = "No"
        Me.ColMS_No.Name = "ColMS_No"
        Me.ColMS_No.ReadOnly = True
        Me.ColMS_No.Width = 46
        '
        'ColMS_MarketSeg
        '
        Me.ColMS_MarketSeg.HeaderText = "Segmentation Code"
        Me.ColMS_MarketSeg.Name = "ColMS_MarketSeg"
        Me.ColMS_MarketSeg.Width = 125
        '
        'ColMS_Desc
        '
        Me.ColMS_Desc.HeaderText = "Description"
        Me.ColMS_Desc.Name = "ColMS_Desc"
        Me.ColMS_Desc.Width = 85
        '
        'ColMS_Primary
        '
        Me.ColMS_Primary.HeaderText = "Primary"
        Me.ColMS_Primary.Name = "ColMS_Primary"
        Me.ColMS_Primary.ReadOnly = True
        Me.ColMS_Primary.Visible = False
        Me.ColMS_Primary.Width = 81
        '
        'TabPage5
        '
        Me.TabPage5.BackColor = System.Drawing.Color.Transparent
        Me.TabPage5.Controls.Add(Me.DgvSalesArea)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage5.Size = New System.Drawing.Size(218, 100)
        Me.TabPage5.TabIndex = 2
        Me.TabPage5.Text = "Sales Area"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DgvSalesArea
        '
        Me.DgvSalesArea.AllowUserToResizeColumns = False
        Me.DgvSalesArea.AllowUserToResizeRows = False
        Me.DgvSalesArea.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalesArea.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSalesArea.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSA_No, Me.ColSA_SalesArea, Me.ColSA_Desc, Me.ColSA_Primary})
        Me.DgvSalesArea.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalesArea.Location = New System.Drawing.Point(3, 3)
        Me.DgvSalesArea.Name = "DgvSalesArea"
        Me.DgvSalesArea.RowHeadersWidth = 24
        Me.DgvSalesArea.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSalesArea.RowTemplate.Height = 24
        Me.DgvSalesArea.Size = New System.Drawing.Size(212, 94)
        Me.DgvSalesArea.TabIndex = 5
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
        Me.ColSA_SalesArea.HeaderText = "Sales Area Code"
        Me.ColSA_SalesArea.Name = "ColSA_SalesArea"
        Me.ColSA_SalesArea.Width = 111
        '
        'ColSA_Desc
        '
        Me.ColSA_Desc.HeaderText = "Description"
        Me.ColSA_Desc.Name = "ColSA_Desc"
        Me.ColSA_Desc.Width = 85
        '
        'ColSA_Primary
        '
        Me.ColSA_Primary.HeaderText = "Primary"
        Me.ColSA_Primary.Name = "ColSA_Primary"
        Me.ColSA_Primary.ReadOnly = True
        Me.ColSA_Primary.Visible = False
        Me.ColSA_Primary.Width = 81
        '
        'TabPage7
        '
        Me.TabPage7.BackColor = System.Drawing.Color.Transparent
        Me.TabPage7.Controls.Add(Me.DgvCustomer)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage7.Size = New System.Drawing.Size(218, 100)
        Me.TabPage7.TabIndex = 4
        Me.TabPage7.Text = "Customer"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'DgvCustomer
        '
        Me.DgvCustomer.AllowUserToResizeColumns = False
        Me.DgvCustomer.AllowUserToResizeRows = False
        Me.DgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColC_No, Me.ColC_Code, Me.ColC_Desc, Me.ColC_Primary})
        Me.DgvCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomer.Location = New System.Drawing.Point(3, 3)
        Me.DgvCustomer.Name = "DgvCustomer"
        Me.DgvCustomer.RowHeadersWidth = 24
        Me.DgvCustomer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomer.RowTemplate.Height = 24
        Me.DgvCustomer.Size = New System.Drawing.Size(212, 94)
        Me.DgvCustomer.TabIndex = 5
        '
        'ColC_No
        '
        Me.ColC_No.HeaderText = "No"
        Me.ColC_No.Name = "ColC_No"
        Me.ColC_No.ReadOnly = True
        Me.ColC_No.Width = 46
        '
        'ColC_Code
        '
        Me.ColC_Code.HeaderText = "Customer Code"
        Me.ColC_Code.Name = "ColC_Code"
        Me.ColC_Code.Width = 104
        '
        'ColC_Desc
        '
        Me.ColC_Desc.HeaderText = "Description"
        Me.ColC_Desc.Name = "ColC_Desc"
        Me.ColC_Desc.Width = 85
        '
        'ColC_Primary
        '
        Me.ColC_Primary.HeaderText = "Primary"
        Me.ColC_Primary.Name = "ColC_Primary"
        Me.ColC_Primary.ReadOnly = True
        Me.ColC_Primary.Visible = False
        Me.ColC_Primary.Width = 81
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.Transparent
        Me.TabPage4.Controls.Add(Me.dgvCustomerGroup)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage4.Size = New System.Drawing.Size(218, 100)
        Me.TabPage4.TabIndex = 1
        Me.TabPage4.Text = "Customer Group"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'dgvCustomerGroup
        '
        Me.dgvCustomerGroup.AllowUserToResizeColumns = False
        Me.dgvCustomerGroup.AllowUserToResizeRows = False
        Me.dgvCustomerGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCustomerGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCustomerGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSG_No, Me.ColSG_SalesGroup, Me.ColSG_Desc, Me.ColSG_Primary})
        Me.dgvCustomerGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomerGroup.Location = New System.Drawing.Point(3, 3)
        Me.dgvCustomerGroup.Name = "dgvCustomerGroup"
        Me.dgvCustomerGroup.RowHeadersWidth = 24
        Me.dgvCustomerGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCustomerGroup.RowTemplate.Height = 24
        Me.dgvCustomerGroup.Size = New System.Drawing.Size(212, 94)
        Me.dgvCustomerGroup.TabIndex = 5
        '
        'ColSG_No
        '
        Me.ColSG_No.HeaderText = "No"
        Me.ColSG_No.Name = "ColSG_No"
        Me.ColSG_No.ReadOnly = True
        Me.ColSG_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSG_No.Width = 27
        '
        'ColSG_SalesGroup
        '
        Me.ColSG_SalesGroup.HeaderText = "Group Code"
        Me.ColSG_SalesGroup.Name = "ColSG_SalesGroup"
        Me.ColSG_SalesGroup.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSG_SalesGroup.Width = 70
        '
        'ColSG_Desc
        '
        Me.ColSG_Desc.HeaderText = "Description"
        Me.ColSG_Desc.Name = "ColSG_Desc"
        Me.ColSG_Desc.ReadOnly = True
        Me.ColSG_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSG_Desc.Width = 66
        '
        'ColSG_Primary
        '
        Me.ColSG_Primary.HeaderText = "Primary"
        Me.ColSG_Primary.Name = "ColSG_Primary"
        Me.ColSG_Primary.ReadOnly = True
        Me.ColSG_Primary.Visible = False
        Me.ColSG_Primary.Width = 81
        '
        'TabPage6
        '
        Me.TabPage6.BackColor = System.Drawing.Color.Transparent
        Me.TabPage6.Controls.Add(Me.DgvCustomerType)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage6.Size = New System.Drawing.Size(218, 100)
        Me.TabPage6.TabIndex = 3
        Me.TabPage6.Text = "Customer Type"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'DgvCustomerType
        '
        Me.DgvCustomerType.AllowUserToResizeColumns = False
        Me.DgvCustomerType.AllowUserToResizeRows = False
        Me.DgvCustomerType.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomerType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomerType.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCT_No, Me.ColCT_CustType, Me.ColCT_Desc, Me.ColCT_Primary})
        Me.DgvCustomerType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomerType.Location = New System.Drawing.Point(3, 3)
        Me.DgvCustomerType.Name = "DgvCustomerType"
        Me.DgvCustomerType.RowHeadersWidth = 24
        Me.DgvCustomerType.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomerType.RowTemplate.Height = 24
        Me.DgvCustomerType.Size = New System.Drawing.Size(212, 94)
        Me.DgvCustomerType.TabIndex = 5
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
        Me.ColCT_CustType.HeaderText = "Type Code"
        Me.ColCT_CustType.Name = "ColCT_CustType"
        Me.ColCT_CustType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_CustType.Width = 65
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
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 61)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 13)
        Me.Label8.TabIndex = 19
        Me.Label8.Text = "Sort by"
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(0, 0)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(584, 534)
        Me.crv1.TabIndex = 3
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'BgwProduct
        '
        '
        'bgwCustomer
        '
        '
        'bgwCustSegmen
        '
        '
        'bgwSalesArea
        '
        '
        'bgwCustType
        '
        '
        'bgwCustGroup
        '
        '
        'FrmLaporanPenjualan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(902, 559)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.SplitContainer8)
        Me.MinimumSize = New System.Drawing.Size(918, 597)
        Me.Name = "FrmLaporanPenjualan"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Historical Sales"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvSortRekap, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSortDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSortRekap2PeriodComp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer8.Panel1.ResumeLayout(False)
        Me.SplitContainer8.Panel1.PerformLayout()
        Me.SplitContainer8.Panel2.ResumeLayout(False)
        Me.SplitContainer8.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage8.ResumeLayout(False)
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.dgvCustomerGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents chkCustomer As System.Windows.Forms.CheckBox
    Friend WithEvents clbCustomer As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkProduct As System.Windows.Forms.CheckBox
    Friend WithEvents clbProduct As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents dtpTo2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboGroupBy As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvSortRekap As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSortDetail As System.Windows.Forms.DataGridView
    Friend WithEvents cboSort1 As System.Windows.Forms.ComboBox
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSortNumber As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColSrtType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents dgvSortRekap2PeriodComp As System.Windows.Forms.DataGridView
    Friend WithEvents ChkProductDesc As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboitemDetail As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer8 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvMarketSegmentation As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents dgvCustomerGroup As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents DgvSalesArea As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents DgvCustomerType As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents DgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BgwProduct As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColC_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_MarketSeg As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_SalesGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_CustType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents bgwCustomer As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCustSegmen As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwSalesArea As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCustType As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCustGroup As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColName2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn1 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn2 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColName1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn3 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn4 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColSA_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_SalesArea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
