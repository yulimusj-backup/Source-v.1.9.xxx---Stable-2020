<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiAOP_CreateByStockIn
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
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvAOP_FilterSupplier = New System.Windows.Forms.DataGridView
        Me.ColAOP_SupplierCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAOP_SupplierName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAOP_PrimarySupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkDirectlySave = New System.Windows.Forms.CheckBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.cboAOP_ShowSupplierMode = New System.Windows.Forms.ComboBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.dtpAOP_From = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.dtpAOP_To = New System.Windows.Forms.DateTimePicker
        Me.btnAOP_Filter = New System.Windows.Forms.Button
        Me.btnCreateAOP = New System.Windows.Forms.Button
        Me.TxtAOP_Desc = New System.Windows.Forms.TextBox
        Me.TxtAOP_PONumber = New System.Windows.Forms.TextBox
        Me.TxtAOP_CodeNo = New System.Windows.Forms.TextBox
        Me.btnAOP_FindPO = New System.Windows.Forms.Button
        Me.TxtAOP_PrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtAOP_PrimaryPO = New System.Windows.Forms.TextBox
        Me.btnAOP_FindCodeNo = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.dgvAOP_TransStockIn = New System.Windows.Forms.DataGridView
        Me.ColItemStockIn_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_TransNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColItemStockIn_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_Supplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_SupplierDONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_CreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_CreateAOP = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItemStockIn_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_PrimarySupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemStockIn_PrimaryGRN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.lblProgressAOP = New System.Windows.Forms.Label
        Me.prgAOP = New System.Windows.Forms.ProgressBar
        Me.llbAOP_CheckAll = New System.Windows.Forms.LinkLabel
        Me.llbAOP_UncheckAll = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvAOP_FilterSupplier, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dgvAOP_TransStockIn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(12, 11)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.chkDirectlySave)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboAOP_ShowSupplierMode)
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnAOP_Filter)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnCreateAOP)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtAOP_Desc)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtAOP_PONumber)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtAOP_CodeNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnAOP_FindPO)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtAOP_PrimaryCodeNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtAOP_PrimaryPO)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnAOP_FindCodeNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Size = New System.Drawing.Size(973, 140)
        Me.SplitContainer2.SplitterDistance = 343
        Me.SplitContainer2.TabIndex = 2
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(343, 140)
        Me.TabControl1.TabIndex = 28
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvAOP_FilterSupplier)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(335, 111)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Supplier"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvAOP_FilterSupplier
        '
        Me.dgvAOP_FilterSupplier.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAOP_FilterSupplier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAOP_FilterSupplier.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAOP_SupplierCode, Me.ColAOP_SupplierName, Me.ColAOP_PrimarySupplier})
        Me.dgvAOP_FilterSupplier.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAOP_FilterSupplier.Location = New System.Drawing.Point(3, 3)
        Me.dgvAOP_FilterSupplier.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvAOP_FilterSupplier.Name = "dgvAOP_FilterSupplier"
        Me.dgvAOP_FilterSupplier.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvAOP_FilterSupplier.RowTemplate.Height = 24
        Me.dgvAOP_FilterSupplier.Size = New System.Drawing.Size(329, 105)
        Me.dgvAOP_FilterSupplier.TabIndex = 0
        '
        'ColAOP_SupplierCode
        '
        Me.ColAOP_SupplierCode.HeaderText = "Code"
        Me.ColAOP_SupplierCode.Name = "ColAOP_SupplierCode"
        Me.ColAOP_SupplierCode.Width = 66
        '
        'ColAOP_SupplierName
        '
        Me.ColAOP_SupplierName.HeaderText = "Name"
        Me.ColAOP_SupplierName.Name = "ColAOP_SupplierName"
        Me.ColAOP_SupplierName.Width = 70
        '
        'ColAOP_PrimarySupplier
        '
        Me.ColAOP_PrimarySupplier.HeaderText = "Primary Supplier"
        Me.ColAOP_PrimarySupplier.Name = "ColAOP_PrimarySupplier"
        Me.ColAOP_PrimarySupplier.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvUser)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(335, 111)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "PO Created By"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_Primary, Me.ColCreatedBy, Me.ColCreatedBy_Name})
        Me.DgvUser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvUser.Location = New System.Drawing.Point(3, 3)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(329, 105)
        Me.DgvUser.TabIndex = 1
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
        '
        'ColCreatedBy_Name
        '
        Me.ColCreatedBy_Name.HeaderText = "Name"
        Me.ColCreatedBy_Name.Name = "ColCreatedBy_Name"
        Me.ColCreatedBy_Name.ReadOnly = True
        Me.ColCreatedBy_Name.Width = 70
        '
        'chkDirectlySave
        '
        Me.chkDirectlySave.AutoSize = True
        Me.chkDirectlySave.Location = New System.Drawing.Point(493, 19)
        Me.chkDirectlySave.Name = "chkDirectlySave"
        Me.chkDirectlySave.Size = New System.Drawing.Size(113, 21)
        Me.chkDirectlySave.TabIndex = 16
        Me.chkDirectlySave.TabStop = False
        Me.chkDirectlySave.Text = "Directly Save"
        Me.chkDirectlySave.UseVisualStyleBackColor = True
        Me.chkDirectlySave.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(9, 114)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(63, 17)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "Code No"
        Me.Label11.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(9, 83)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(82, 17)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "PO Number"
        Me.Label10.Visible = False
        '
        'cboAOP_ShowSupplierMode
        '
        Me.cboAOP_ShowSupplierMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAOP_ShowSupplierMode.FormattingEnabled = True
        Me.cboAOP_ShowSupplierMode.Items.AddRange(New Object() {"Show Supplier Code", "Show Supplier Name", "Show Supplier Code & Name"})
        Me.cboAOP_ShowSupplierMode.Location = New System.Drawing.Point(275, 47)
        Me.cboAOP_ShowSupplierMode.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboAOP_ShowSupplierMode.Name = "cboAOP_ShowSupplierMode"
        Me.cboAOP_ShowSupplierMode.Size = New System.Drawing.Size(253, 24)
        Me.cboAOP_ShowSupplierMode.TabIndex = 7
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.dtpAOP_From)
        Me.GroupBox4.Controls.Add(Me.Label8)
        Me.GroupBox4.Controls.Add(Me.Label9)
        Me.GroupBox4.Controls.Add(Me.dtpAOP_To)
        Me.GroupBox4.Location = New System.Drawing.Point(3, 2)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox4.Size = New System.Drawing.Size(261, 75)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Stock In Period"
        '
        'dtpAOP_From
        '
        Me.dtpAOP_From.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpAOP_From.CustomFormat = "dd MMM yyyy"
        Me.dtpAOP_From.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAOP_From.Location = New System.Drawing.Point(52, 21)
        Me.dtpAOP_From.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpAOP_From.Name = "dtpAOP_From"
        Me.dtpAOP_From.Size = New System.Drawing.Size(200, 22)
        Me.dtpAOP_From.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(25, 17)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "To"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 22)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 17)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "From"
        '
        'dtpAOP_To
        '
        Me.dtpAOP_To.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpAOP_To.CustomFormat = "dd MMM yyyy"
        Me.dtpAOP_To.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAOP_To.Location = New System.Drawing.Point(52, 47)
        Me.dtpAOP_To.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpAOP_To.Name = "dtpAOP_To"
        Me.dtpAOP_To.Size = New System.Drawing.Size(200, 22)
        Me.dtpAOP_To.TabIndex = 1
        '
        'btnAOP_Filter
        '
        Me.btnAOP_Filter.Location = New System.Drawing.Point(275, 18)
        Me.btnAOP_Filter.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnAOP_Filter.Name = "btnAOP_Filter"
        Me.btnAOP_Filter.Size = New System.Drawing.Size(84, 23)
        Me.btnAOP_Filter.TabIndex = 0
        Me.btnAOP_Filter.Text = "Filter"
        Me.btnAOP_Filter.UseVisualStyleBackColor = True
        '
        'btnCreateAOP
        '
        Me.btnCreateAOP.Location = New System.Drawing.Point(365, 18)
        Me.btnCreateAOP.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCreateAOP.Name = "btnCreateAOP"
        Me.btnCreateAOP.Size = New System.Drawing.Size(122, 23)
        Me.btnCreateAOP.TabIndex = 1
        Me.btnCreateAOP.Text = "Create AOP"
        Me.btnCreateAOP.UseVisualStyleBackColor = True
        '
        'TxtAOP_Desc
        '
        Me.TxtAOP_Desc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAOP_Desc.Location = New System.Drawing.Point(341, 111)
        Me.TxtAOP_Desc.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtAOP_Desc.Name = "TxtAOP_Desc"
        Me.TxtAOP_Desc.ReadOnly = True
        Me.TxtAOP_Desc.Size = New System.Drawing.Size(278, 22)
        Me.TxtAOP_Desc.TabIndex = 14
        Me.TxtAOP_Desc.TabStop = False
        Me.TxtAOP_Desc.Visible = False
        '
        'TxtAOP_PONumber
        '
        Me.TxtAOP_PONumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAOP_PONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAOP_PONumber.Location = New System.Drawing.Point(97, 81)
        Me.TxtAOP_PONumber.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtAOP_PONumber.Name = "TxtAOP_PONumber"
        Me.TxtAOP_PONumber.Size = New System.Drawing.Size(143, 22)
        Me.TxtAOP_PONumber.TabIndex = 1
        Me.TxtAOP_PONumber.TabStop = False
        Me.TxtAOP_PONumber.Visible = False
        '
        'TxtAOP_CodeNo
        '
        Me.TxtAOP_CodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAOP_CodeNo.Location = New System.Drawing.Point(97, 111)
        Me.TxtAOP_CodeNo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtAOP_CodeNo.Name = "TxtAOP_CodeNo"
        Me.TxtAOP_CodeNo.Size = New System.Drawing.Size(168, 22)
        Me.TxtAOP_CodeNo.TabIndex = 5
        Me.TxtAOP_CodeNo.TabStop = False
        Me.TxtAOP_CodeNo.Visible = False
        '
        'btnAOP_FindPO
        '
        Me.btnAOP_FindPO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAOP_FindPO.Location = New System.Drawing.Point(246, 81)
        Me.btnAOP_FindPO.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnAOP_FindPO.Name = "btnAOP_FindPO"
        Me.btnAOP_FindPO.Size = New System.Drawing.Size(64, 23)
        Me.btnAOP_FindPO.TabIndex = 2
        Me.btnAOP_FindPO.TabStop = False
        Me.btnAOP_FindPO.Text = "Find"
        Me.btnAOP_FindPO.UseVisualStyleBackColor = True
        Me.btnAOP_FindPO.Visible = False
        '
        'TxtAOP_PrimaryCodeNo
        '
        Me.TxtAOP_PrimaryCodeNo.Location = New System.Drawing.Point(77, 112)
        Me.TxtAOP_PrimaryCodeNo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtAOP_PrimaryCodeNo.Name = "TxtAOP_PrimaryCodeNo"
        Me.TxtAOP_PrimaryCodeNo.Size = New System.Drawing.Size(33, 22)
        Me.TxtAOP_PrimaryCodeNo.TabIndex = 13
        Me.TxtAOP_PrimaryCodeNo.TabStop = False
        Me.TxtAOP_PrimaryCodeNo.Text = "-1"
        Me.TxtAOP_PrimaryCodeNo.Visible = False
        '
        'TxtAOP_PrimaryPO
        '
        Me.TxtAOP_PrimaryPO.Location = New System.Drawing.Point(78, 82)
        Me.TxtAOP_PrimaryPO.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtAOP_PrimaryPO.Name = "TxtAOP_PrimaryPO"
        Me.TxtAOP_PrimaryPO.Size = New System.Drawing.Size(33, 22)
        Me.TxtAOP_PrimaryPO.TabIndex = 4
        Me.TxtAOP_PrimaryPO.TabStop = False
        Me.TxtAOP_PrimaryPO.Text = "-1"
        Me.TxtAOP_PrimaryPO.Visible = False
        '
        'btnAOP_FindCodeNo
        '
        Me.btnAOP_FindCodeNo.Location = New System.Drawing.Point(271, 111)
        Me.btnAOP_FindCodeNo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnAOP_FindCodeNo.Name = "btnAOP_FindCodeNo"
        Me.btnAOP_FindCodeNo.Size = New System.Drawing.Size(64, 23)
        Me.btnAOP_FindCodeNo.TabIndex = 6
        Me.btnAOP_FindCodeNo.TabStop = False
        Me.btnAOP_FindCodeNo.Text = "Find"
        Me.btnAOP_FindCodeNo.UseVisualStyleBackColor = True
        Me.btnAOP_FindCodeNo.Visible = False
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(379, 18)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(95, 23)
        Me.btnClose.TabIndex = 15
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'dgvAOP_TransStockIn
        '
        Me.dgvAOP_TransStockIn.AllowUserToAddRows = False
        Me.dgvAOP_TransStockIn.AllowUserToDeleteRows = False
        Me.dgvAOP_TransStockIn.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAOP_TransStockIn.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAOP_TransStockIn.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItemStockIn_No, Me.ColItemStockIn_TransNo, Me.ColItemStockIn_Date, Me.ColItemStockIn_Supplier, Me.ColItemStockIn_Remarks, Me.ColItemStockIn_SupplierDONumber, Me.ColItemStockIn_CreatedBy, Me.ColItemStockIn_CreateAOP, Me.ColItemStockIn_Primary, Me.ColItemStockIn_PrimaryCurrency, Me.ColItemStockIn_PrimarySupplier, Me.ColItemStockIn_PrimaryGRN})
        Me.dgvAOP_TransStockIn.Location = New System.Drawing.Point(11, 155)
        Me.dgvAOP_TransStockIn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvAOP_TransStockIn.Name = "dgvAOP_TransStockIn"
        Me.dgvAOP_TransStockIn.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAOP_TransStockIn.RowTemplate.Height = 24
        Me.dgvAOP_TransStockIn.Size = New System.Drawing.Size(971, 440)
        Me.dgvAOP_TransStockIn.TabIndex = 23
        '
        'ColItemStockIn_No
        '
        Me.ColItemStockIn_No.HeaderText = "No"
        Me.ColItemStockIn_No.Name = "ColItemStockIn_No"
        Me.ColItemStockIn_No.Width = 51
        '
        'ColItemStockIn_TransNo
        '
        Me.ColItemStockIn_TransNo.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColItemStockIn_TransNo.HeaderText = "Document No"
        Me.ColItemStockIn_TransNo.Name = "ColItemStockIn_TransNo"
        Me.ColItemStockIn_TransNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItemStockIn_TransNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColItemStockIn_TransNo.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColItemStockIn_TransNo.Width = 119
        '
        'ColItemStockIn_Date
        '
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        Me.ColItemStockIn_Date.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItemStockIn_Date.HeaderText = "Date"
        Me.ColItemStockIn_Date.Name = "ColItemStockIn_Date"
        Me.ColItemStockIn_Date.Width = 63
        '
        'ColItemStockIn_Supplier
        '
        Me.ColItemStockIn_Supplier.HeaderText = "Supplier"
        Me.ColItemStockIn_Supplier.Name = "ColItemStockIn_Supplier"
        Me.ColItemStockIn_Supplier.Width = 85
        '
        'ColItemStockIn_Remarks
        '
        Me.ColItemStockIn_Remarks.HeaderText = "Remarks"
        Me.ColItemStockIn_Remarks.Name = "ColItemStockIn_Remarks"
        Me.ColItemStockIn_Remarks.Width = 89
        '
        'ColItemStockIn_SupplierDONumber
        '
        Me.ColItemStockIn_SupplierDONumber.HeaderText = "Supplier's DO Number"
        Me.ColItemStockIn_SupplierDONumber.Name = "ColItemStockIn_SupplierDONumber"
        Me.ColItemStockIn_SupplierDONumber.Width = 174
        '
        'ColItemStockIn_CreatedBy
        '
        Me.ColItemStockIn_CreatedBy.HeaderText = "Created By"
        Me.ColItemStockIn_CreatedBy.Name = "ColItemStockIn_CreatedBy"
        Me.ColItemStockIn_CreatedBy.Width = 103
        '
        'ColItemStockIn_CreateAOP
        '
        Me.ColItemStockIn_CreateAOP.HeaderText = "Create AOP"
        Me.ColItemStockIn_CreateAOP.Name = "ColItemStockIn_CreateAOP"
        Me.ColItemStockIn_CreateAOP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItemStockIn_CreateAOP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColItemStockIn_CreateAOP.Width = 108
        '
        'ColItemStockIn_Primary
        '
        Me.ColItemStockIn_Primary.HeaderText = "Primary"
        Me.ColItemStockIn_Primary.Name = "ColItemStockIn_Primary"
        Me.ColItemStockIn_Primary.Visible = False
        Me.ColItemStockIn_Primary.Width = 81
        '
        'ColItemStockIn_PrimaryCurrency
        '
        Me.ColItemStockIn_PrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColItemStockIn_PrimaryCurrency.Name = "ColItemStockIn_PrimaryCurrency"
        Me.ColItemStockIn_PrimaryCurrency.Visible = False
        Me.ColItemStockIn_PrimaryCurrency.Width = 142
        '
        'ColItemStockIn_PrimarySupplier
        '
        Me.ColItemStockIn_PrimarySupplier.HeaderText = "Primary Supplier"
        Me.ColItemStockIn_PrimarySupplier.Name = "ColItemStockIn_PrimarySupplier"
        Me.ColItemStockIn_PrimarySupplier.Visible = False
        Me.ColItemStockIn_PrimarySupplier.Width = 137
        '
        'ColItemStockIn_PrimaryGRN
        '
        Me.ColItemStockIn_PrimaryGRN.HeaderText = "Primary GRN"
        Me.ColItemStockIn_PrimaryGRN.Name = "ColItemStockIn_PrimaryGRN"
        Me.ColItemStockIn_PrimaryGRN.Visible = False
        Me.ColItemStockIn_PrimaryGRN.Width = 116
        '
        'lblProgressAOP
        '
        Me.lblProgressAOP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProgressAOP.AutoSize = True
        Me.lblProgressAOP.Location = New System.Drawing.Point(541, 602)
        Me.lblProgressAOP.Name = "lblProgressAOP"
        Me.lblProgressAOP.Size = New System.Drawing.Size(51, 17)
        Me.lblProgressAOP.TabIndex = 27
        Me.lblProgressAOP.Text = "Label6"
        Me.lblProgressAOP.Visible = False
        '
        'prgAOP
        '
        Me.prgAOP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.prgAOP.Location = New System.Drawing.Point(180, 599)
        Me.prgAOP.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prgAOP.Name = "prgAOP"
        Me.prgAOP.Size = New System.Drawing.Size(355, 23)
        Me.prgAOP.TabIndex = 26
        Me.prgAOP.Visible = False
        '
        'llbAOP_CheckAll
        '
        Me.llbAOP_CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbAOP_CheckAll.AutoSize = True
        Me.llbAOP_CheckAll.Location = New System.Drawing.Point(13, 602)
        Me.llbAOP_CheckAll.Name = "llbAOP_CheckAll"
        Me.llbAOP_CheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbAOP_CheckAll.TabIndex = 24
        Me.llbAOP_CheckAll.TabStop = True
        Me.llbAOP_CheckAll.Text = "Check All"
        '
        'llbAOP_UncheckAll
        '
        Me.llbAOP_UncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbAOP_UncheckAll.AutoSize = True
        Me.llbAOP_UncheckAll.Location = New System.Drawing.Point(92, 602)
        Me.llbAOP_UncheckAll.Name = "llbAOP_UncheckAll"
        Me.llbAOP_UncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbAOP_UncheckAll.TabIndex = 25
        Me.llbAOP_UncheckAll.TabStop = True
        Me.llbAOP_UncheckAll.Text = "Uncheck All"
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiAOP_CreateByStockIn
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.dgvAOP_TransStockIn)
        Me.Controls.Add(Me.lblProgressAOP)
        Me.Controls.Add(Me.prgAOP)
        Me.Controls.Add(Me.llbAOP_CheckAll)
        Me.Controls.Add(Me.llbAOP_UncheckAll)
        Me.Controls.Add(Me.SplitContainer2)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmTransaksiAOP_CreateByStockIn"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create AOP By Stock In"
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvAOP_FilterSupplier, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.dgvAOP_TransStockIn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvAOP_FilterSupplier As System.Windows.Forms.DataGridView
    Friend WithEvents ColAOP_SupplierCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAOP_SupplierName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAOP_PrimarySupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cboAOP_ShowSupplierMode As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpAOP_From As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dtpAOP_To As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnAOP_Filter As System.Windows.Forms.Button
    Friend WithEvents btnCreateAOP As System.Windows.Forms.Button
    Friend WithEvents TxtAOP_Desc As System.Windows.Forms.TextBox
    Friend WithEvents TxtAOP_PONumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtAOP_CodeNo As System.Windows.Forms.TextBox
    Friend WithEvents btnAOP_FindPO As System.Windows.Forms.Button
    Friend WithEvents TxtAOP_PrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtAOP_PrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents btnAOP_FindCodeNo As System.Windows.Forms.Button
    Friend WithEvents dgvAOP_TransStockIn As System.Windows.Forms.DataGridView
    Friend WithEvents lblProgressAOP As System.Windows.Forms.Label
    Friend WithEvents prgAOP As System.Windows.Forms.ProgressBar
    Friend WithEvents llbAOP_CheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbAOP_UncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents chkDirectlySave As System.Windows.Forms.CheckBox
    Friend WithEvents ColItemStockIn_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_TransNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColItemStockIn_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_Supplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_SupplierDONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_CreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_CreateAOP As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItemStockIn_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_PrimarySupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemStockIn_PrimaryGRN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
End Class
