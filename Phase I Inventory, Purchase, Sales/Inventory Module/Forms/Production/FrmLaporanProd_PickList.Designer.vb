<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_PickList
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboData = New System.Windows.Forms.ComboBox
        Me.btnTransfer = New System.Windows.Forms.Button
        Me.llbLocationPriority = New System.Windows.Forms.LinkLabel
        Me.llbColumnSetting = New System.Windows.Forms.LinkLabel
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.chkShowNotReadyOnly = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvProcess = New System.Windows.Forms.DataGridView
        Me.colProcess_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvWU = New System.Windows.Forms.DataGridView
        Me.ColWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWU_Group = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.dgvLocation = New System.Windows.Forms.DataGridView
        Me.ColLocation_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Group = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColLocation_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.dgvJO = New System.Windows.Forms.DataGridView
        Me.ColJo_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJo_JONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJo_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJo_Identififer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_CodeNoRM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_DescRM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Warehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_WU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_WONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_JONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyAllocated = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyUsed = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyNeed = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyOnSite = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyTransfer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyMinus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyPackage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_TransferDetail = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_Transfer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_PrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryLoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.BgwItem = New System.ComponentModel.BackgroundWorker
        Me.BgwProduct = New System.ComponentModel.BackgroundWorker
        Me.BgwProcess = New System.ComponentModel.BackgroundWorker
        Me.BgwWu = New System.ComponentModel.BackgroundWorker
        Me.BgwLocation = New System.ComponentModel.BackgroundWorker
        Me.BgwJO = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvProcess, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvWU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvJO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbUncheckAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbCheckAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvItem)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(926, 563)
        Me.SplitContainer1.SplitterDistance = 175
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboData)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnTransfer)
        Me.SplitContainer2.Panel1.Controls.Add(Me.llbLocationPriority)
        Me.SplitContainer2.Panel1.Controls.Add(Me.llbColumnSetting)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnPrint)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnShow)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkShowNotReadyOnly)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpTo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpFrom)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboFormat)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer2.Size = New System.Drawing.Size(926, 175)
        Me.SplitContainer2.SplitterDistance = 266
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Data"
        '
        'cboData
        '
        Me.cboData.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboData.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboData.FormattingEnabled = True
        Me.cboData.ItemHeight = 13
        Me.cboData.Items.AddRange(New Object() {"All Data", "Raw Material Only", "Work In Proccess Only"})
        Me.cboData.Location = New System.Drawing.Point(98, 80)
        Me.cboData.Name = "cboData"
        Me.cboData.Size = New System.Drawing.Size(165, 21)
        Me.cboData.TabIndex = 3
        '
        'btnTransfer
        '
        Me.btnTransfer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTransfer.Location = New System.Drawing.Point(106, 131)
        Me.btnTransfer.Name = "btnTransfer"
        Me.btnTransfer.Size = New System.Drawing.Size(75, 23)
        Me.btnTransfer.TabIndex = 6
        Me.btnTransfer.Text = "Transfer"
        Me.btnTransfer.UseVisualStyleBackColor = True
        '
        'llbLocationPriority
        '
        Me.llbLocationPriority.AutoSize = True
        Me.llbLocationPriority.Location = New System.Drawing.Point(96, 157)
        Me.llbLocationPriority.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbLocationPriority.Name = "llbLocationPriority"
        Me.llbLocationPriority.Size = New System.Drawing.Size(82, 13)
        Me.llbLocationPriority.TabIndex = 9
        Me.llbLocationPriority.TabStop = True
        Me.llbLocationPriority.Text = "Location Priority"
        '
        'llbColumnSetting
        '
        Me.llbColumnSetting.AutoSize = True
        Me.llbColumnSetting.Location = New System.Drawing.Point(9, 156)
        Me.llbColumnSetting.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbColumnSetting.Name = "llbColumnSetting"
        Me.llbColumnSetting.Size = New System.Drawing.Size(78, 13)
        Me.llbColumnSetting.TabIndex = 8
        Me.llbColumnSetting.TabStop = True
        Me.llbColumnSetting.Text = "Column Setting"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(187, 131)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 7
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(19, 131)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 23)
        Me.btnShow.TabIndex = 5
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'chkShowNotReadyOnly
        '
        Me.chkShowNotReadyOnly.AutoSize = True
        Me.chkShowNotReadyOnly.Location = New System.Drawing.Point(98, 61)
        Me.chkShowNotReadyOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowNotReadyOnly.Name = "chkShowNotReadyOnly"
        Me.chkShowNotReadyOnly.Size = New System.Drawing.Size(101, 17)
        Me.chkShowNotReadyOnly.TabIndex = 2
        Me.chkShowNotReadyOnly.Text = "Not Ready Only"
        Me.chkShowNotReadyOnly.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 39)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy HH.mm"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(98, 37)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(165, 20)
        Me.dtpTo.TabIndex = 1
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy HH.mm"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(98, 11)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(165, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Prod. Schedule"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.ItemHeight = 13
        Me.cboFormat.Items.AddRange(New Object() {"Recap", "Detail"})
        Me.cboFormat.Location = New System.Drawing.Point(98, 105)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(165, 21)
        Me.cboFormat.TabIndex = 4
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(657, 175)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvProduct)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(649, 149)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Raw Material"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvProduct
        '
        Me.dgvProduct.AllowUserToResizeColumns = False
        Me.dgvProduct.AllowUserToResizeRows = False
        Me.dgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvProduct.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvProduct.ColumnHeadersHeight = 25
        Me.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProduct_No, Me.colProduct_GroupLevel, Me.colProduct_Code, Me.colProduct_Desc, Me.colProduct_Primary})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvProduct.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvProduct.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProduct.Location = New System.Drawing.Point(2, 2)
        Me.dgvProduct.MultiSelect = False
        Me.dgvProduct.Name = "dgvProduct"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvProduct.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvProduct.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvProduct.RowTemplate.Height = 24
        Me.dgvProduct.Size = New System.Drawing.Size(645, 145)
        Me.dgvProduct.TabIndex = 0
        '
        'colProduct_No
        '
        Me.colProduct_No.HeaderText = "No"
        Me.colProduct_No.Name = "colProduct_No"
        Me.colProduct_No.Width = 46
        '
        'colProduct_GroupLevel
        '
        Me.colProduct_GroupLevel.HeaderText = "Group Level"
        Me.colProduct_GroupLevel.Items.AddRange(New Object() {"Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.colProduct_GroupLevel.Name = "colProduct_GroupLevel"
        Me.colProduct_GroupLevel.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colProduct_GroupLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colProduct_GroupLevel.Width = 90
        '
        'colProduct_Code
        '
        Me.colProduct_Code.HeaderText = "Code"
        Me.colProduct_Code.Name = "colProduct_Code"
        Me.colProduct_Code.Width = 57
        '
        'colProduct_Desc
        '
        Me.colProduct_Desc.HeaderText = "Desc"
        Me.colProduct_Desc.Name = "colProduct_Desc"
        Me.colProduct_Desc.ReadOnly = True
        Me.colProduct_Desc.Width = 57
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 66
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvProcess)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(649, 149)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Process"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvProcess
        '
        Me.dgvProcess.AllowUserToResizeColumns = False
        Me.dgvProcess.AllowUserToResizeRows = False
        Me.dgvProcess.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvProcess.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvProcess.ColumnHeadersHeight = 25
        Me.dgvProcess.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvProcess.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProcess_No, Me.colProcess_Code, Me.colProcess_Desc, Me.colProcess_Primary})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvProcess.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvProcess.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProcess.Location = New System.Drawing.Point(2, 2)
        Me.dgvProcess.MultiSelect = False
        Me.dgvProcess.Name = "dgvProcess"
        Me.dgvProcess.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvProcess.RowTemplate.Height = 24
        Me.dgvProcess.Size = New System.Drawing.Size(645, 145)
        Me.dgvProcess.TabIndex = 1
        '
        'colProcess_No
        '
        Me.colProcess_No.HeaderText = "No"
        Me.colProcess_No.Name = "colProcess_No"
        Me.colProcess_No.Width = 46
        '
        'colProcess_Code
        '
        Me.colProcess_Code.HeaderText = "Process Code"
        Me.colProcess_Code.Name = "colProcess_Code"
        Me.colProcess_Code.Width = 98
        '
        'colProcess_Desc
        '
        Me.colProcess_Desc.HeaderText = "Desc"
        Me.colProcess_Desc.Name = "colProcess_Desc"
        Me.colProcess_Desc.ReadOnly = True
        Me.colProcess_Desc.Width = 57
        '
        'colProcess_Primary
        '
        Me.colProcess_Primary.HeaderText = "Primary"
        Me.colProcess_Primary.Name = "colProcess_Primary"
        Me.colProcess_Primary.Visible = False
        Me.colProcess_Primary.Width = 66
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvWU)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(649, 149)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Work Unit"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvWU
        '
        Me.dgvWU.AllowUserToResizeColumns = False
        Me.dgvWU.AllowUserToResizeRows = False
        Me.dgvWU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWU.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvWU.ColumnHeadersHeight = 25
        Me.dgvWU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWU_No, Me.ColWU_Group, Me.ColWU_Code, Me.ColWU_Desc, Me.ColWU_Primary})
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvWU.DefaultCellStyle = DataGridViewCellStyle7
        Me.dgvWU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvWU.Location = New System.Drawing.Point(2, 2)
        Me.dgvWU.MultiSelect = False
        Me.dgvWU.Name = "dgvWU"
        Me.dgvWU.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvWU.RowTemplate.Height = 24
        Me.dgvWU.Size = New System.Drawing.Size(645, 145)
        Me.dgvWU.TabIndex = 2
        '
        'ColWU_No
        '
        Me.ColWU_No.HeaderText = "No"
        Me.ColWU_No.Name = "ColWU_No"
        Me.ColWU_No.Width = 46
        '
        'ColWU_Group
        '
        Me.ColWU_Group.HeaderText = "Group Level"
        Me.ColWU_Group.Items.AddRange(New Object() {"Work Unit Group", "Work Unit"})
        Me.ColWU_Group.Name = "ColWU_Group"
        Me.ColWU_Group.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWU_Group.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColWU_Group.Width = 90
        '
        'ColWU_Code
        '
        Me.ColWU_Code.HeaderText = "Code"
        Me.ColWU_Code.Name = "ColWU_Code"
        Me.ColWU_Code.Width = 57
        '
        'ColWU_Desc
        '
        Me.ColWU_Desc.HeaderText = "Desc"
        Me.ColWU_Desc.Name = "ColWU_Desc"
        Me.ColWU_Desc.ReadOnly = True
        Me.ColWU_Desc.Width = 57
        '
        'ColWU_Primary
        '
        Me.ColWU_Primary.HeaderText = "Primary"
        Me.ColWU_Primary.Name = "ColWU_Primary"
        Me.ColWU_Primary.Visible = False
        Me.ColWU_Primary.Width = 66
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.dgvLocation)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(649, 149)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Input Location"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'dgvLocation
        '
        Me.dgvLocation.AllowUserToResizeColumns = False
        Me.dgvLocation.AllowUserToResizeRows = False
        Me.dgvLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvLocation.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvLocation.ColumnHeadersHeight = 25
        Me.dgvLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColLocation_No, Me.ColLocation_Group, Me.ColLocation_Code, Me.ColLocation_Desc, Me.ColLocation_Primary})
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvLocation.DefaultCellStyle = DataGridViewCellStyle9
        Me.dgvLocation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvLocation.Location = New System.Drawing.Point(2, 2)
        Me.dgvLocation.MultiSelect = False
        Me.dgvLocation.Name = "dgvLocation"
        Me.dgvLocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvLocation.RowTemplate.Height = 24
        Me.dgvLocation.Size = New System.Drawing.Size(645, 145)
        Me.dgvLocation.TabIndex = 3
        '
        'ColLocation_No
        '
        Me.ColLocation_No.HeaderText = "No"
        Me.ColLocation_No.Name = "ColLocation_No"
        Me.ColLocation_No.Width = 46
        '
        'ColLocation_Group
        '
        Me.ColLocation_Group.HeaderText = "Group Level"
        Me.ColLocation_Group.Items.AddRange(New Object() {"Location", "Warehouse"})
        Me.ColLocation_Group.Name = "ColLocation_Group"
        Me.ColLocation_Group.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_Group.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColLocation_Group.Width = 90
        '
        'ColLocation_Code
        '
        Me.ColLocation_Code.HeaderText = "Code"
        Me.ColLocation_Code.Name = "ColLocation_Code"
        Me.ColLocation_Code.Width = 57
        '
        'ColLocation_Desc
        '
        Me.ColLocation_Desc.HeaderText = "Desc"
        Me.ColLocation_Desc.Name = "ColLocation_Desc"
        Me.ColLocation_Desc.ReadOnly = True
        Me.ColLocation_Desc.Width = 57
        '
        'ColLocation_Primary
        '
        Me.ColLocation_Primary.HeaderText = "Primary"
        Me.ColLocation_Primary.Name = "ColLocation_Primary"
        Me.ColLocation_Primary.Visible = False
        Me.ColLocation_Primary.Width = 66
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.dgvJO)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(649, 149)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Job Order"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'dgvJO
        '
        Me.dgvJO.AllowUserToResizeColumns = False
        Me.dgvJO.AllowUserToResizeRows = False
        Me.dgvJO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvJO.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        Me.dgvJO.ColumnHeadersHeight = 25
        Me.dgvJO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvJO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColJo_No, Me.ColJo_JONo, Me.ColJo_Date, Me.ColJo_Identififer})
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvJO.DefaultCellStyle = DataGridViewCellStyle11
        Me.dgvJO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvJO.Location = New System.Drawing.Point(2, 2)
        Me.dgvJO.MultiSelect = False
        Me.dgvJO.Name = "dgvJO"
        Me.dgvJO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvJO.RowTemplate.Height = 24
        Me.dgvJO.Size = New System.Drawing.Size(645, 145)
        Me.dgvJO.TabIndex = 4
        '
        'ColJo_No
        '
        Me.ColJo_No.HeaderText = "No"
        Me.ColJo_No.Name = "ColJo_No"
        Me.ColJo_No.Width = 46
        '
        'ColJo_JONo
        '
        Me.ColJo_JONo.HeaderText = "JO No"
        Me.ColJo_JONo.Name = "ColJo_JONo"
        Me.ColJo_JONo.Width = 62
        '
        'ColJo_Date
        '
        Me.ColJo_Date.HeaderText = "Date"
        Me.ColJo_Date.Name = "ColJo_Date"
        Me.ColJo_Date.ReadOnly = True
        Me.ColJo_Date.Width = 55
        '
        'ColJo_Identififer
        '
        Me.ColJo_Identififer.HeaderText = "Identifier"
        Me.ColJo_Identififer.Name = "ColJo_Identififer"
        Me.ColJo_Identififer.Visible = False
        Me.ColJo_Identififer.Width = 72
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(141, 363)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 20)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Update Location"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(63, 365)
        Me.llbUncheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 4
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(9, 365)
        Me.llbCheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 3
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvItem.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle12
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_CodeNoRM, Me.ColItem_DescRM, Me.ColItem_Location, Me.ColItem_Warehouse, Me.ColItem_WU, Me.ColItem_WONO, Me.ColItem_JONo, Me.ColItem_StartTime, Me.ColItem_Unit, Me.ColItem_QtyAllocated, Me.ColItem_QtyUsed, Me.ColItem_QtyNeed, Me.ColItem_QtyOnSite, Me.ColItem_QtyTransfer, Me.ColItem_QtyMinus, Me.ColItem_QtyPackage, Me.ColItem_TransferDetail, Me.ColItem_Transfer, Me.ColItem_PrimaryWU, Me.ColItem_Identifier, Me.ColItem_PrimaryProduct, Me.ColItem_PrimaryLoc})
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle20.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvItem.DefaultCellStyle = DataGridViewCellStyle20
        Me.dgvItem.Location = New System.Drawing.Point(0, 2)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle21.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle21.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle21.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle21.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle21.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvItem.RowHeadersDefaultCellStyle = DataGridViewCellStyle21
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(926, 359)
        Me.dgvItem.TabIndex = 0
        '
        'ColItem_No
        '
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 27
        '
        'ColItem_CodeNoRM
        '
        Me.ColItem_CodeNoRM.HeaderText = "Code No"
        Me.ColItem_CodeNoRM.Name = "ColItem_CodeNoRM"
        Me.ColItem_CodeNoRM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_CodeNoRM.Width = 55
        '
        'ColItem_DescRM
        '
        Me.ColItem_DescRM.HeaderText = "Description"
        Me.ColItem_DescRM.Name = "ColItem_DescRM"
        Me.ColItem_DescRM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_DescRM.Width = 66
        '
        'ColItem_Location
        '
        Me.ColItem_Location.HeaderText = "Location"
        Me.ColItem_Location.Name = "ColItem_Location"
        Me.ColItem_Location.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Location.Width = 54
        '
        'ColItem_Warehouse
        '
        Me.ColItem_Warehouse.HeaderText = "Warehouse"
        Me.ColItem_Warehouse.Name = "ColItem_Warehouse"
        Me.ColItem_Warehouse.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Warehouse.Width = 68
        '
        'ColItem_WU
        '
        Me.ColItem_WU.HeaderText = "Work Unit"
        Me.ColItem_WU.Name = "ColItem_WU"
        Me.ColItem_WU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_WU.Width = 61
        '
        'ColItem_WONO
        '
        Me.ColItem_WONO.HeaderText = "Wo No"
        Me.ColItem_WONO.Name = "ColItem_WONO"
        Me.ColItem_WONO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_WONO.Width = 47
        '
        'ColItem_JONo
        '
        Me.ColItem_JONo.HeaderText = "JO No"
        Me.ColItem_JONo.Name = "ColItem_JONo"
        Me.ColItem_JONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_JONo.Width = 43
        '
        'ColItem_StartTime
        '
        Me.ColItem_StartTime.HeaderText = "Start Time"
        Me.ColItem_StartTime.Name = "ColItem_StartTime"
        Me.ColItem_StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_StartTime.Width = 61
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Unit"
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Unit.Width = 32
        '
        'ColItem_QtyAllocated
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyAllocated.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColItem_QtyAllocated.HeaderText = "Qty Allocated"
        Me.ColItem_QtyAllocated.Name = "ColItem_QtyAllocated"
        Me.ColItem_QtyAllocated.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyAllocated.Visible = False
        Me.ColItem_QtyAllocated.Width = 76
        '
        'ColItem_QtyUsed
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyUsed.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColItem_QtyUsed.HeaderText = "Qty Used"
        Me.ColItem_QtyUsed.Name = "ColItem_QtyUsed"
        Me.ColItem_QtyUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyUsed.Visible = False
        Me.ColItem_QtyUsed.Width = 57
        '
        'ColItem_QtyNeed
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyNeed.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColItem_QtyNeed.HeaderText = "Qty Need"
        Me.ColItem_QtyNeed.Name = "ColItem_QtyNeed"
        Me.ColItem_QtyNeed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyNeed.Width = 58
        '
        'ColItem_QtyOnSite
        '
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyOnSite.DefaultCellStyle = DataGridViewCellStyle16
        Me.ColItem_QtyOnSite.HeaderText = "Qty On Site"
        Me.ColItem_QtyOnSite.Name = "ColItem_QtyOnSite"
        Me.ColItem_QtyOnSite.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyOnSite.Width = 67
        '
        'ColItem_QtyTransfer
        '
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyTransfer.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColItem_QtyTransfer.HeaderText = "Qty Transfer"
        Me.ColItem_QtyTransfer.Name = "ColItem_QtyTransfer"
        Me.ColItem_QtyTransfer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyTransfer.Width = 71
        '
        'ColItem_QtyMinus
        '
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyMinus.DefaultCellStyle = DataGridViewCellStyle18
        Me.ColItem_QtyMinus.HeaderText = "Qty Minus"
        Me.ColItem_QtyMinus.Name = "ColItem_QtyMinus"
        Me.ColItem_QtyMinus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyMinus.Width = 60
        '
        'ColItem_QtyPackage
        '
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyPackage.DefaultCellStyle = DataGridViewCellStyle19
        Me.ColItem_QtyPackage.HeaderText = "Qty Package"
        Me.ColItem_QtyPackage.Name = "ColItem_QtyPackage"
        Me.ColItem_QtyPackage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_QtyPackage.Width = 75
        '
        'ColItem_TransferDetail
        '
        Me.ColItem_TransferDetail.HeaderText = "Transfer Detail "
        Me.ColItem_TransferDetail.Name = "ColItem_TransferDetail"
        Me.ColItem_TransferDetail.Width = 85
        '
        'ColItem_Transfer
        '
        Me.ColItem_Transfer.HeaderText = "Transfer"
        Me.ColItem_Transfer.Name = "ColItem_Transfer"
        Me.ColItem_Transfer.Width = 52
        '
        'ColItem_PrimaryWU
        '
        Me.ColItem_PrimaryWU.HeaderText = "Primary WU"
        Me.ColItem_PrimaryWU.Name = "ColItem_PrimaryWU"
        Me.ColItem_PrimaryWU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_PrimaryWU.Visible = False
        Me.ColItem_PrimaryWU.Width = 69
        '
        'ColItem_Identifier
        '
        Me.ColItem_Identifier.HeaderText = "Identifier"
        Me.ColItem_Identifier.Name = "ColItem_Identifier"
        Me.ColItem_Identifier.ReadOnly = True
        Me.ColItem_Identifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Identifier.Width = 53
        '
        'ColItem_PrimaryProduct
        '
        Me.ColItem_PrimaryProduct.HeaderText = "Primary Code"
        Me.ColItem_PrimaryProduct.Name = "ColItem_PrimaryProduct"
        Me.ColItem_PrimaryProduct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_PrimaryProduct.Visible = False
        Me.ColItem_PrimaryProduct.Width = 75
        '
        'ColItem_PrimaryLoc
        '
        Me.ColItem_PrimaryLoc.HeaderText = "Primary Location"
        Me.ColItem_PrimaryLoc.Name = "ColItem_PrimaryLoc"
        Me.ColItem_PrimaryLoc.ReadOnly = True
        Me.ColItem_PrimaryLoc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_PrimaryLoc.Visible = False
        Me.ColItem_PrimaryLoc.Width = 91
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(400, 127)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 19)
        Me.btnClose.TabIndex = 2
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'BgwItem
        '
        '
        'BgwProduct
        '
        '
        'BgwProcess
        '
        '
        'BgwWu
        '
        '
        'BgwLocation
        '
        '
        'BgwJO
        '
        '
        'FrmLaporanProd_PickList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(926, 563)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(942, 601)
        Me.Name = "FrmLaporanProd_PickList"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Pick List"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvProcess, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvWU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.dgvJO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkShowNotReadyOnly As System.Windows.Forms.CheckBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents dgvProcess As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvWU As System.Windows.Forms.DataGridView
    Friend WithEvents dgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents dgvJO As System.Windows.Forms.DataGridView
    Friend WithEvents llbLocationPriority As System.Windows.Forms.LinkLabel
    Friend WithEvents llbColumnSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents btnTransfer As System.Windows.Forms.Button
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents BgwItem As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwProduct As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwProcess As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwWu As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwLocation As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwJO As System.ComponentModel.BackgroundWorker
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWU_Group As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Group As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJo_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJo_JONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJo_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJo_Identififer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboData As System.Windows.Forms.ComboBox
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_CodeNoRM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DescRM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Warehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_WU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_WONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_JONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyAllocated As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyUsed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyNeed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyOnSite As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyTransfer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyMinus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyPackage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_TransferDetail As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_Transfer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_PrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryLoc As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
