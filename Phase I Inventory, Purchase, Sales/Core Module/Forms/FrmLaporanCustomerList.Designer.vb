<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanCustomerList
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
        Me.Gbo = New System.Windows.Forms.GroupBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkJoinDate1 = New System.Windows.Forms.CheckBox
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvCustGroup = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCG_CustGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCG_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCG_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvSalesArea = New System.Windows.Forms.DataGridView
        Me.ColSA_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_SalesArea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvMarketSegmentation = New System.Windows.Forms.DataGridView
        Me.ColMS_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_MarketSeg = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.DgvCustomerType = New System.Windows.Forms.DataGridView
        Me.ColCT_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_CustType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.DgvCustRef = New System.Windows.Forms.DataGridView
        Me.ColCustomerReference_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerReference_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerReference_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BtnShow = New System.Windows.Forms.Button
        Me.ChkMultipleBank = New System.Windows.Forms.CheckBox
        Me.ChkTaxData = New System.Windows.Forms.CheckBox
        Me.ChkMultipleCP = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.ChkMultipleAddress = New System.Windows.Forms.CheckBox
        Me.CboGroup = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Bgw = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.Gbo.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvCustGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DgvCustRef, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Gbo
        '
        Me.Gbo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Gbo.Controls.Add(Me.GroupBox1)
        Me.Gbo.Controls.Add(Me.TabControl1)
        Me.Gbo.Controls.Add(Me.BtnClose)
        Me.Gbo.Controls.Add(Me.BtnShow)
        Me.Gbo.Controls.Add(Me.ChkMultipleBank)
        Me.Gbo.Controls.Add(Me.ChkTaxData)
        Me.Gbo.Controls.Add(Me.ChkMultipleCP)
        Me.Gbo.Controls.Add(Me.GroupBox2)
        Me.Gbo.Controls.Add(Me.ChkMultipleAddress)
        Me.Gbo.Controls.Add(Me.CboGroup)
        Me.Gbo.Controls.Add(Me.Label1)
        Me.Gbo.Location = New System.Drawing.Point(5, -2)
        Me.Gbo.Name = "Gbo"
        Me.Gbo.Size = New System.Drawing.Size(813, 505)
        Me.Gbo.TabIndex = 0
        Me.Gbo.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ChkJoinDate1)
        Me.GroupBox1.Controls.Add(Me.DtpFrom)
        Me.GroupBox1.Controls.Add(Me.DtpTo)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(604, 11)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(200, 89)
        Me.GroupBox1.TabIndex = 17
        Me.GroupBox1.TabStop = False
        '
        'ChkJoinDate1
        '
        Me.ChkJoinDate1.AutoSize = True
        Me.ChkJoinDate1.Location = New System.Drawing.Point(4, 0)
        Me.ChkJoinDate1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkJoinDate1.Name = "ChkJoinDate1"
        Me.ChkJoinDate1.Size = New System.Drawing.Size(71, 17)
        Me.ChkJoinDate1.TabIndex = 17
        Me.ChkJoinDate1.Text = "Join Date"
        Me.ChkJoinDate1.UseVisualStyleBackColor = True
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(68, 18)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(121, 20)
        Me.DtpFrom.TabIndex = 9
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(68, 43)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(121, 20)
        Me.DtpTo.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "From"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Location = New System.Drawing.Point(9, 11)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(394, 120)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvCustGroup)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(386, 94)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Customer Group"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvCustGroup
        '
        Me.DgvCustGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustGroup.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvCustGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.ColCG_CustGroup, Me.ColCG_Desc, Me.ColCG_Primary})
        Me.DgvCustGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustGroup.Location = New System.Drawing.Point(2, 2)
        Me.DgvCustGroup.Name = "DgvCustGroup"
        Me.DgvCustGroup.RowTemplate.Height = 24
        Me.DgvCustGroup.Size = New System.Drawing.Size(382, 90)
        Me.DgvCustGroup.TabIndex = 10
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "No"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 46
        '
        'ColCG_CustGroup
        '
        Me.ColCG_CustGroup.HeaderText = "Customer Group"
        Me.ColCG_CustGroup.Name = "ColCG_CustGroup"
        Me.ColCG_CustGroup.Width = 108
        '
        'ColCG_Desc
        '
        Me.ColCG_Desc.HeaderText = "Desc"
        Me.ColCG_Desc.Name = "ColCG_Desc"
        Me.ColCG_Desc.Width = 57
        '
        'ColCG_Primary
        '
        Me.ColCG_Primary.HeaderText = "Primary"
        Me.ColCG_Primary.Name = "ColCG_Primary"
        Me.ColCG_Primary.Visible = False
        Me.ColCG_Primary.Width = 81
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvSalesArea)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(386, 94)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Sales Area"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvSalesArea
        '
        Me.DgvSalesArea.AllowUserToResizeColumns = False
        Me.DgvSalesArea.AllowUserToResizeRows = False
        Me.DgvSalesArea.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalesArea.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSalesArea.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSA_No, Me.ColSA_SalesArea, Me.ColSA_Desc, Me.ColSA_Primary})
        Me.DgvSalesArea.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalesArea.Location = New System.Drawing.Point(2, 2)
        Me.DgvSalesArea.Name = "DgvSalesArea"
        Me.DgvSalesArea.RowHeadersWidth = 24
        Me.DgvSalesArea.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSalesArea.RowTemplate.Height = 24
        Me.DgvSalesArea.Size = New System.Drawing.Size(382, 90)
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
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvMarketSegmentation)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Size = New System.Drawing.Size(386, 94)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Customer Segmentation"
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
        Me.DgvMarketSegmentation.Location = New System.Drawing.Point(2, 2)
        Me.DgvMarketSegmentation.Name = "DgvMarketSegmentation"
        Me.DgvMarketSegmentation.RowHeadersWidth = 24
        Me.DgvMarketSegmentation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvMarketSegmentation.RowTemplate.Height = 24
        Me.DgvMarketSegmentation.Size = New System.Drawing.Size(382, 90)
        Me.DgvMarketSegmentation.TabIndex = 8
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
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DgvCustomerType)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage4.Size = New System.Drawing.Size(386, 94)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Customer Type"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DgvCustomerType
        '
        Me.DgvCustomerType.AllowUserToResizeColumns = False
        Me.DgvCustomerType.AllowUserToResizeRows = False
        Me.DgvCustomerType.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomerType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomerType.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCT_No, Me.ColCT_CustType, Me.ColCT_Desc, Me.ColCT_Primary})
        Me.DgvCustomerType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomerType.Location = New System.Drawing.Point(2, 2)
        Me.DgvCustomerType.Name = "DgvCustomerType"
        Me.DgvCustomerType.RowHeadersWidth = 24
        Me.DgvCustomerType.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomerType.RowTemplate.Height = 24
        Me.DgvCustomerType.Size = New System.Drawing.Size(382, 90)
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
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.DgvCustRef)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage5.Size = New System.Drawing.Size(386, 94)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Customer Reference"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DgvCustRef
        '
        Me.DgvCustRef.AllowUserToResizeColumns = False
        Me.DgvCustRef.AllowUserToResizeRows = False
        Me.DgvCustRef.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustRef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustRef.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCustomerReference_No, Me.ColCustomerReference, Me.ColCustomerReference_Desc, Me.ColCustomerReference_Primary})
        Me.DgvCustRef.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustRef.Location = New System.Drawing.Point(2, 2)
        Me.DgvCustRef.Name = "DgvCustRef"
        Me.DgvCustRef.RowHeadersWidth = 24
        Me.DgvCustRef.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustRef.RowTemplate.Height = 24
        Me.DgvCustRef.Size = New System.Drawing.Size(382, 90)
        Me.DgvCustRef.TabIndex = 2
        '
        'ColCustomerReference_No
        '
        Me.ColCustomerReference_No.HeaderText = "No"
        Me.ColCustomerReference_No.Name = "ColCustomerReference_No"
        Me.ColCustomerReference_No.ReadOnly = True
        Me.ColCustomerReference_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustomerReference_No.Width = 27
        '
        'ColCustomerReference
        '
        Me.ColCustomerReference.HeaderText = "Customer Reference"
        Me.ColCustomerReference.Name = "ColCustomerReference"
        Me.ColCustomerReference.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustomerReference.Width = 110
        '
        'ColCustomerReference_Desc
        '
        Me.ColCustomerReference_Desc.HeaderText = "Description"
        Me.ColCustomerReference_Desc.Name = "ColCustomerReference_Desc"
        Me.ColCustomerReference_Desc.ReadOnly = True
        Me.ColCustomerReference_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustomerReference_Desc.Width = 66
        '
        'ColCustomerReference_Primary
        '
        Me.ColCustomerReference_Primary.HeaderText = "Primary"
        Me.ColCustomerReference_Primary.Name = "ColCustomerReference_Primary"
        Me.ColCustomerReference_Primary.Visible = False
        Me.ColCustomerReference_Primary.Width = 81
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(724, 109)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 8
        Me.BtnClose.Text = "&Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(643, 109)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(75, 23)
        Me.BtnShow.TabIndex = 7
        Me.BtnShow.Text = "&Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'ChkMultipleBank
        '
        Me.ChkMultipleBank.AutoSize = True
        Me.ChkMultipleBank.Location = New System.Drawing.Point(419, 64)
        Me.ChkMultipleBank.Name = "ChkMultipleBank"
        Me.ChkMultipleBank.Size = New System.Drawing.Size(163, 17)
        Me.ChkMultipleBank.TabIndex = 5
        Me.ChkMultipleBank.Text = "Show Multiple Bank Account"
        Me.ChkMultipleBank.UseVisualStyleBackColor = True
        '
        'ChkTaxData
        '
        Me.ChkTaxData.AutoSize = True
        Me.ChkTaxData.Location = New System.Drawing.Point(419, 85)
        Me.ChkTaxData.Name = "ChkTaxData"
        Me.ChkTaxData.Size = New System.Drawing.Size(100, 17)
        Me.ChkTaxData.TabIndex = 2
        Me.ChkTaxData.Text = "Show Tax Data"
        Me.ChkTaxData.UseVisualStyleBackColor = True
        '
        'ChkMultipleCP
        '
        Me.ChkMultipleCP.AutoSize = True
        Me.ChkMultipleCP.Location = New System.Drawing.Point(419, 41)
        Me.ChkMultipleCP.Name = "ChkMultipleCP"
        Me.ChkMultipleCP.Size = New System.Drawing.Size(168, 17)
        Me.ChkMultipleCP.TabIndex = 4
        Me.ChkMultipleCP.Text = "Show Multiple Contact Person"
        Me.ChkMultipleCP.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Crv)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 138)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(801, 361)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'Crv
        '
        Me.Crv.ActiveViewIndex = -1
        Me.Crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crv.DisplayGroupTree = False
        Me.Crv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Crv.EnableDrillDown = False
        Me.Crv.Location = New System.Drawing.Point(3, 16)
        Me.Crv.Name = "Crv"
        Me.Crv.SelectionFormula = ""
        Me.Crv.ShowCloseButton = False
        Me.Crv.ShowGroupTreeButton = False
        Me.Crv.ShowRefreshButton = False
        Me.Crv.Size = New System.Drawing.Size(795, 342)
        Me.Crv.TabIndex = 0
        Me.Crv.ViewTimeSelectionFormula = ""
        '
        'ChkMultipleAddress
        '
        Me.ChkMultipleAddress.AutoSize = True
        Me.ChkMultipleAddress.Location = New System.Drawing.Point(419, 18)
        Me.ChkMultipleAddress.Name = "ChkMultipleAddress"
        Me.ChkMultipleAddress.Size = New System.Drawing.Size(133, 17)
        Me.ChkMultipleAddress.TabIndex = 3
        Me.ChkMultipleAddress.Text = "Show Multiple Address"
        Me.ChkMultipleAddress.UseVisualStyleBackColor = True
        '
        'CboGroup
        '
        Me.CboGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroup.FormattingEnabled = True
        Me.CboGroup.Items.AddRange(New Object() {"Customer Group", "Sales Area", "Customer Segmentation", "Customer Type", "Customer Reference"})
        Me.CboGroup.Location = New System.Drawing.Point(477, 112)
        Me.CboGroup.Name = "CboGroup"
        Me.CboGroup.Size = New System.Drawing.Size(123, 21)
        Me.CboGroup.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(417, 114)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Group by"
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
        'FrmLaporanCustomerList
        '
        Me.AcceptButton = Me.BtnShow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(830, 511)
        Me.Controls.Add(Me.Gbo)
        Me.MinimumSize = New System.Drawing.Size(802, 549)
        Me.Name = "FrmLaporanCustomerList"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer List Report"
        Me.Gbo.ResumeLayout(False)
        Me.Gbo.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvCustGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DgvCustRef, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Gbo As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents ChkMultipleAddress As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboGroup As System.Windows.Forms.ComboBox
    Friend WithEvents ChkTaxData As System.Windows.Forms.CheckBox
    Friend WithEvents ChkMultipleBank As System.Windows.Forms.CheckBox
    Friend WithEvents ChkMultipleCP As System.Windows.Forms.CheckBox
    Friend WithEvents Bgw As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents DgvCustGroup As System.Windows.Forms.DataGridView
    Friend WithEvents DgvSalesArea As System.Windows.Forms.DataGridView
    Friend WithEvents ColSA_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_SalesArea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvMarketSegmentation As System.Windows.Forms.DataGridView
    Friend WithEvents ColMS_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_MarketSeg As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvCustomerType As System.Windows.Forms.DataGridView
    Friend WithEvents ColCT_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_CustType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvCustRef As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCG_CustGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCG_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCG_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerReference_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerReference_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerReference_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkJoinDate1 As System.Windows.Forms.CheckBox
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
