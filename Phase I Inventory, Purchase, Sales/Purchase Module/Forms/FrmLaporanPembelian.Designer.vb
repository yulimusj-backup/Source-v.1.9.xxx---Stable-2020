<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPembelian
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
        Me.Label8 = New System.Windows.Forms.Label
        Me.CboItemDetail = New System.Windows.Forms.ComboBox
        Me.ChkProductDesc = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.CboDataFrom = New System.Windows.Forms.ComboBox
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
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvSupplier = New System.Windows.Forms.DataGridView
        Me.ColSupp_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvSortRekap2PeriodComp = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewComboBoxColumn3 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.DataGridViewComboBoxColumn4 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.dgvSortRekap = New System.Windows.Forms.DataGridView
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSortNumber = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColSrtType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.dgvSortDetail = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewComboBoxColumn1 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.DataGridViewComboBoxColumn2 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Label9 = New System.Windows.Forms.Label
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BgwSupplier = New System.ComponentModel.BackgroundWorker
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvSupplier, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvSortRekap2PeriodComp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSortRekap, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSortDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(435, 410)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(853, 532)
        Me.TabControl1.TabIndex = 4
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.SplitContainer3)
        Me.TabPage1.Controls.Add(Me.SplitContainer1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(845, 506)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Filter && Grouping"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(858, 475)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox5)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(829, 300)
        Me.SplitContainer3.SplitterDistance = 408
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
        Me.GroupBox5.Size = New System.Drawing.Size(408, 300)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        '
        'chkCustomer
        '
        Me.chkCustomer.AutoSize = True
        Me.chkCustomer.Location = New System.Drawing.Point(7, -1)
        Me.chkCustomer.Name = "chkCustomer"
        Me.chkCustomer.Size = New System.Drawing.Size(78, 17)
        Me.chkCustomer.TabIndex = 0
        Me.chkCustomer.Text = "All Supplier"
        Me.chkCustomer.UseVisualStyleBackColor = True
        '
        'clbCustomer
        '
        Me.clbCustomer.CheckOnClick = True
        Me.clbCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbCustomer.FormattingEnabled = True
        Me.clbCustomer.Location = New System.Drawing.Point(3, 16)
        Me.clbCustomer.Name = "clbCustomer"
        Me.clbCustomer.Size = New System.Drawing.Size(402, 274)
        Me.clbCustomer.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkProduct)
        Me.GroupBox1.Controls.Add(Me.clbProduct)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(417, 300)
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
        Me.clbProduct.Size = New System.Drawing.Size(411, 274)
        Me.clbProduct.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 6)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Size = New System.Drawing.Size(832, 195)
        Me.SplitContainer1.SplitterDistance = 389
        Me.SplitContainer1.TabIndex = 18
        Me.SplitContainer1.TabStop = False
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(803, 487)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Report View"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 97)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(57, 13)
        Me.Label8.TabIndex = 24
        Me.Label8.Text = "Item Detail"
        '
        'CboItemDetail
        '
        Me.CboItemDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboItemDetail.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboItemDetail.FormattingEnabled = True
        Me.CboItemDetail.Items.AddRange(New Object() {"Code No", "Category", "Sub Brand", "Brand", "Product Group"})
        Me.CboItemDetail.Location = New System.Drawing.Point(81, 93)
        Me.CboItemDetail.Name = "CboItemDetail"
        Me.CboItemDetail.Size = New System.Drawing.Size(251, 21)
        Me.CboItemDetail.TabIndex = 23
        '
        'ChkProductDesc
        '
        Me.ChkProductDesc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkProductDesc.AutoSize = True
        Me.ChkProductDesc.Location = New System.Drawing.Point(26, 512)
        Me.ChkProductDesc.Name = "ChkProductDesc"
        Me.ChkProductDesc.Size = New System.Drawing.Size(149, 17)
        Me.ChkProductDesc.TabIndex = 22
        Me.ChkProductDesc.Text = "Show Product Description"
        Me.ChkProductDesc.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 228)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 13)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Data By"
        '
        'CboDataFrom
        '
        Me.CboDataFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboDataFrom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboDataFrom.FormattingEnabled = True
        Me.CboDataFrom.Items.AddRange(New Object() {"AOP", "PO"})
        Me.CboDataFrom.Location = New System.Drawing.Point(81, 224)
        Me.CboDataFrom.Name = "CboDataFrom"
        Me.CboDataFrom.Size = New System.Drawing.Size(251, 21)
        Me.CboDataFrom.TabIndex = 20
        '
        'cboSort1
        '
        Me.cboSort1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSort1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSort1.FormattingEnabled = True
        Me.cboSort1.Items.AddRange(New Object() {"Ascending", "Descending", "Qty - Ascending", "Qty - Descending", "Amount - Ascending", "Amount - Descending"})
        Me.cboSort1.Location = New System.Drawing.Point(81, 66)
        Me.cboSort1.Name = "cboSort1"
        Me.cboSort1.Size = New System.Drawing.Size(251, 21)
        Me.cboSort1.TabIndex = 18
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 43)
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
        Me.cboFormat.Location = New System.Drawing.Point(81, 12)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(251, 21)
        Me.cboFormat.TabIndex = 14
        '
        'cboGroupBy
        '
        Me.cboGroupBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboGroupBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupBy.FormattingEnabled = True
        Me.cboGroupBy.Items.AddRange(New Object() {"Rekap", "Detail"})
        Me.cboGroupBy.Location = New System.Drawing.Point(81, 39)
        Me.cboGroupBy.Name = "cboGroupBy"
        Me.cboGroupBy.Size = New System.Drawing.Size(251, 21)
        Me.cboGroupBy.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 124)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Period I"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 16)
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
        Me.dtpFrom.Location = New System.Drawing.Point(81, 120)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(251, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'btnShowReport
        '
        Me.btnShowReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShowReport.Location = New System.Drawing.Point(244, 511)
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
        Me.dtpTo2.Location = New System.Drawing.Point(81, 198)
        Me.dtpTo2.Name = "dtpTo2"
        Me.dtpTo2.Size = New System.Drawing.Size(251, 20)
        Me.dtpTo2.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 150)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "to"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 202)
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
        Me.dtpTo.Location = New System.Drawing.Point(81, 146)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(251, 20)
        Me.dtpTo.TabIndex = 9
        '
        'dtpFrom2
        '
        Me.dtpFrom2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom2.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom2.Location = New System.Drawing.Point(81, 172)
        Me.dtpFrom2.Name = "dtpFrom2"
        Me.dtpFrom2.Size = New System.Drawing.Size(251, 20)
        Me.dtpFrom2.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Period II"
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(20, 40)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(99, 30)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(20, 71)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(243, 17)
        Me.CheckBox1.TabIndex = 19
        Me.CheckBox1.TabStop = False
        Me.CheckBox1.Text = "Show supplier invoice instead of AOP Number"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 537)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(902, 22)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.ChkProductDesc)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TabControl2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.GroupBox2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnShowReport)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer4.Panel1.Controls.Add(Me.cboFormat)
        Me.SplitContainer4.Panel1.Controls.Add(Me.CboDataFrom)
        Me.SplitContainer4.Panel1.Controls.Add(Me.CboItemDetail)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer4.Panel1.Controls.Add(Me.dtpTo2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.dtpFrom2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer4.Panel1.Controls.Add(Me.dtpFrom)
        Me.SplitContainer4.Panel1.Controls.Add(Me.cboGroupBy)
        Me.SplitContainer4.Panel1.Controls.Add(Me.dtpTo)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer4.Panel1.Controls.Add(Me.cboSort1)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.crv1)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer4.Size = New System.Drawing.Size(902, 535)
        Me.SplitContainer4.SplitterDistance = 338
        Me.SplitContainer4.TabIndex = 6
        Me.SplitContainer4.TabStop = False
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Location = New System.Drawing.Point(6, 388)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(327, 119)
        Me.TabControl2.TabIndex = 27
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage3.Controls.Add(Me.DgvSupplier)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage3.Size = New System.Drawing.Size(319, 93)
        Me.TabPage3.TabIndex = 0
        Me.TabPage3.Text = "Supplier"
        '
        'DgvSupplier
        '
        Me.DgvSupplier.AllowUserToResizeColumns = False
        Me.DgvSupplier.AllowUserToResizeRows = False
        Me.DgvSupplier.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSupplier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSupplier.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSupp_No, Me.ColSupp_Code, Me.ColSupp_Name, Me.ColSupp_Remarks, Me.ColSupp_Primary})
        Me.DgvSupplier.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSupplier.Location = New System.Drawing.Point(3, 3)
        Me.DgvSupplier.Name = "DgvSupplier"
        Me.DgvSupplier.RowHeadersWidth = 24
        Me.DgvSupplier.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSupplier.RowTemplate.Height = 24
        Me.DgvSupplier.Size = New System.Drawing.Size(313, 87)
        Me.DgvSupplier.TabIndex = 5
        '
        'ColSupp_No
        '
        Me.ColSupp_No.HeaderText = "No"
        Me.ColSupp_No.Name = "ColSupp_No"
        Me.ColSupp_No.ReadOnly = True
        Me.ColSupp_No.Width = 46
        '
        'ColSupp_Code
        '
        Me.ColSupp_Code.HeaderText = "Supplier Code"
        Me.ColSupp_Code.Name = "ColSupp_Code"
        Me.ColSupp_Code.Width = 98
        '
        'ColSupp_Name
        '
        Me.ColSupp_Name.HeaderText = "Supplier Name"
        Me.ColSupp_Name.Name = "ColSupp_Name"
        Me.ColSupp_Name.Width = 101
        '
        'ColSupp_Remarks
        '
        Me.ColSupp_Remarks.HeaderText = "Remarks"
        Me.ColSupp_Remarks.Name = "ColSupp_Remarks"
        Me.ColSupp_Remarks.ReadOnly = True
        Me.ColSupp_Remarks.Visible = False
        Me.ColSupp_Remarks.Width = 89
        '
        'ColSupp_Primary
        '
        Me.ColSupp_Primary.HeaderText = "Primary"
        Me.ColSupp_Primary.Name = "ColSupp_Primary"
        Me.ColSupp_Primary.Visible = False
        Me.ColSupp_Primary.Width = 81
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage4.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage4.Size = New System.Drawing.Size(294, 74)
        Me.TabPage4.TabIndex = 1
        Me.TabPage4.Text = "Product"
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidth = 24
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(288, 68)
        Me.dgvItemProductGroup.TabIndex = 5
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.ReadOnly = True
        Me.ColProductGroup_No.Width = 46
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.HeaderText = "Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.Width = 57
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.ReadOnly = True
        Me.ColProductGroup_Remarks.Visible = False
        Me.ColProductGroup_Remarks.Width = 89
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvSortRekap2PeriodComp)
        Me.GroupBox2.Controls.Add(Me.dgvSortRekap)
        Me.GroupBox2.Controls.Add(Me.CheckBox1)
        Me.GroupBox2.Controls.Add(Me.btnExit)
        Me.GroupBox2.Controls.Add(Me.dgvSortDetail)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 251)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(327, 131)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Item Detail Sort"
        '
        'dgvSortRekap2PeriodComp
        '
        Me.dgvSortRekap2PeriodComp.AllowUserToAddRows = False
        Me.dgvSortRekap2PeriodComp.AllowUserToDeleteRows = False
        Me.dgvSortRekap2PeriodComp.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSortRekap2PeriodComp.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSortRekap2PeriodComp.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn2, Me.DataGridViewComboBoxColumn3, Me.DataGridViewComboBoxColumn4})
        Me.dgvSortRekap2PeriodComp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSortRekap2PeriodComp.Location = New System.Drawing.Point(3, 16)
        Me.dgvSortRekap2PeriodComp.Name = "dgvSortRekap2PeriodComp"
        Me.dgvSortRekap2PeriodComp.RowHeadersWidth = 24
        Me.dgvSortRekap2PeriodComp.RowTemplate.Height = 24
        Me.dgvSortRekap2PeriodComp.Size = New System.Drawing.Size(321, 112)
        Me.dgvSortRekap2PeriodComp.TabIndex = 2
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Column Name"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 98
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
        Me.dgvSortRekap.RowTemplate.Height = 24
        Me.dgvSortRekap.Size = New System.Drawing.Size(321, 112)
        Me.dgvSortRekap.TabIndex = 0
        '
        'ColName
        '
        Me.ColName.HeaderText = "Column Name"
        Me.ColName.Name = "ColName"
        Me.ColName.ReadOnly = True
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
        Me.dgvSortDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewComboBoxColumn1, Me.DataGridViewComboBoxColumn2})
        Me.dgvSortDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSortDetail.Location = New System.Drawing.Point(3, 16)
        Me.dgvSortDetail.Name = "dgvSortDetail"
        Me.dgvSortDetail.RowTemplate.Height = 24
        Me.dgvSortDetail.Size = New System.Drawing.Size(321, 112)
        Me.dgvSortDetail.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Column Name"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 98
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
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 70)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(26, 13)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "Sort"
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(0, 0)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(558, 536)
        Me.crv1.TabIndex = 5
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'BackgroundWorker3
        '
        '
        'BgwSupplier
        '
        '
        'FrmLaporanPembelian
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(902, 559)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.SplitContainer4)
        Me.MinimumSize = New System.Drawing.Size(918, 597)
        Me.Name = "FrmLaporanPembelian"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Laporan Pembelian"
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
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvSupplier, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvSortRekap2PeriodComp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSortRekap, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSortDetail, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cboSort1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents cboGroupBy As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents dtpTo2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents CboDataFrom As System.Windows.Forms.ComboBox
    Friend WithEvents ChkProductDesc As System.Windows.Forms.CheckBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CboItemDetail As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvSupplier As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvSortRekap2PeriodComp As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSortRekap As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSortDetail As System.Windows.Forms.DataGridView
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwSupplier As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColSupp_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn3 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn4 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSortNumber As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColSrtType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn1 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn2 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
