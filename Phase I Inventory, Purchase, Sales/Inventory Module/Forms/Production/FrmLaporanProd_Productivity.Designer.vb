<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_Productivity
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
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.colMO_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colMO_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colMO_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colMO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DgvMO = New System.Windows.Forms.DataGridView
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.CboStatus = New System.Windows.Forms.ComboBox
        Me.BtnShow = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.BtnClose = New System.Windows.Forms.Button
        Me.Crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.ChkToleranceQtyResult = New System.Windows.Forms.CheckBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtToleranceQtyResult = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.CboQtyResult = New System.Windows.Forms.ComboBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.ChkToleranceProdRate = New System.Windows.Forms.CheckBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtToleranceProdRate = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.CboProdRate = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.ChkToleranceGP = New System.Windows.Forms.CheckBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtToleranceGP = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.CboGP = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ChkToleranceQtyProduce = New System.Windows.Forms.CheckBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtToleranceQtyProduce = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.CboQtyProduce = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.BgwProduct = New System.ComponentModel.BackgroundWorker
        Me.BgwMO = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvMO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(92, 49)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(218, 22)
        Me.dtpTo.TabIndex = 3
        '
        'colMO_DocNo
        '
        Me.colMO_DocNo.HeaderText = "MO No"
        Me.colMO_DocNo.Name = "colMO_DocNo"
        Me.colMO_DocNo.Width = 77
        '
        'colMO_Primary
        '
        Me.colMO_Primary.HeaderText = "Primary"
        Me.colMO_Primary.Name = "colMO_Primary"
        Me.colMO_Primary.Visible = False
        Me.colMO_Primary.Width = 66
        '
        'colMO_Date
        '
        Me.colMO_Date.HeaderText = "Date"
        Me.colMO_Date.Name = "colMO_Date"
        Me.colMO_Date.Width = 63
        '
        'dgvProduct
        '
        Me.dgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProduct.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProduct_No, Me.colProduct_GroupLevel, Me.colProduct_Code, Me.colProduct_Desc, Me.colProduct_Primary})
        Me.dgvProduct.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProduct.Location = New System.Drawing.Point(4, 19)
        Me.dgvProduct.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvProduct.Name = "dgvProduct"
        Me.dgvProduct.RowTemplate.Height = 24
        Me.dgvProduct.Size = New System.Drawing.Size(313, 190)
        Me.dgvProduct.TabIndex = 0
        '
        'colProduct_No
        '
        Me.colProduct_No.HeaderText = "No"
        Me.colProduct_No.Name = "colProduct_No"
        Me.colProduct_No.Width = 51
        '
        'colProduct_GroupLevel
        '
        Me.colProduct_GroupLevel.HeaderText = "Group Level"
        Me.colProduct_GroupLevel.Items.AddRange(New Object() {"Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.colProduct_GroupLevel.Name = "colProduct_GroupLevel"
        Me.colProduct_GroupLevel.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colProduct_GroupLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colProduct_GroupLevel.Width = 111
        '
        'colProduct_Code
        '
        Me.colProduct_Code.HeaderText = "Code"
        Me.colProduct_Code.Name = "colProduct_Code"
        Me.colProduct_Code.Width = 66
        '
        'colProduct_Desc
        '
        Me.colProduct_Desc.HeaderText = "Desc"
        Me.colProduct_Desc.Name = "colProduct_Desc"
        Me.colProduct_Desc.ReadOnly = True
        Me.colProduct_Desc.Width = 65
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 66
        '
        'colMO_No
        '
        Me.colMO_No.HeaderText = "No"
        Me.colMO_No.Name = "colMO_No"
        Me.colMO_No.Width = 51
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvProduct)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 201)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(321, 213)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Product"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 22)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'DgvMO
        '
        Me.DgvMO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvMO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvMO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvMO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colMO_No, Me.colMO_DocNo, Me.colMO_Date, Me.colMO_Primary})
        Me.DgvMO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvMO.Location = New System.Drawing.Point(4, 19)
        Me.DgvMO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DgvMO.Name = "DgvMO"
        Me.DgvMO.RowTemplate.Height = 24
        Me.DgvMO.Size = New System.Drawing.Size(313, 249)
        Me.DgvMO.TabIndex = 8
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.CboStatus)
        Me.GroupBox1.Controls.Add(Me.BtnShow)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.CboFormat)
        Me.GroupBox1.Controls.Add(Me.BtnClose)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 9)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(321, 185)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "MO Date"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(11, 116)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(48, 17)
        Me.Label16.TabIndex = 10
        Me.Label16.Text = "Status"
        '
        'CboStatus
        '
        Me.CboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatus.FormattingEnabled = True
        Me.CboStatus.Items.AddRange(New Object() {"All", "Finished Only", "On Progress"})
        Me.CboStatus.Location = New System.Drawing.Point(92, 112)
        Me.CboStatus.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboStatus.Name = "CboStatus"
        Me.CboStatus.Size = New System.Drawing.Size(218, 24)
        Me.CboStatus.TabIndex = 5
        '
        'BtnShow
        '
        Me.BtnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnShow.Location = New System.Drawing.Point(227, 149)
        Me.BtnShow.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(85, 28)
        Me.BtnShow.TabIndex = 6
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 54)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(92, 18)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(219, 22)
        Me.dtpFrom.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 85)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(52, 17)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Format"
        '
        'CboFormat
        '
        Me.CboFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Productivity by JO", "Productivity Recap", "Work Unit Recap", "Work Unit Detail", "Process Recap", "Process Detail"})
        Me.CboFormat.Location = New System.Drawing.Point(92, 81)
        Me.CboFormat.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(218, 24)
        Me.CboFormat.TabIndex = 4
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(226, 149)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(87, 28)
        Me.BtnClose.TabIndex = 12
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'Crv
        '
        Me.Crv.ActiveViewIndex = -1
        Me.Crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crv.DisplayGroupTree = False
        Me.Crv.Location = New System.Drawing.Point(4, 103)
        Me.Crv.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Crv.Name = "Crv"
        Me.Crv.SelectionFormula = ""
        Me.Crv.ShowCloseButton = False
        Me.Crv.ShowGroupTreeButton = False
        Me.Crv.ShowRefreshButton = False
        Me.Crv.Size = New System.Drawing.Size(986, 585)
        Me.Crv.TabIndex = 1
        Me.Crv.ViewTimeSelectionFormula = ""
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.DgvMO)
        Me.GroupBox4.Location = New System.Drawing.Point(3, 421)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(321, 272)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "MO"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Crv)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox3)
        Me.SplitContainer1.Size = New System.Drawing.Size(1333, 692)
        Me.SplitContainer1.SplitterDistance = 329
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 2
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.ChkToleranceQtyResult)
        Me.GroupBox8.Controls.Add(Me.Label17)
        Me.GroupBox8.Controls.Add(Me.TxtToleranceQtyResult)
        Me.GroupBox8.Controls.Add(Me.Label18)
        Me.GroupBox8.Controls.Add(Me.Label19)
        Me.GroupBox8.Controls.Add(Me.CboQtyResult)
        Me.GroupBox8.Controls.Add(Me.Label20)
        Me.GroupBox8.Location = New System.Drawing.Point(247, 9)
        Me.GroupBox8.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox8.Size = New System.Drawing.Size(235, 87)
        Me.GroupBox8.TabIndex = 13
        Me.GroupBox8.TabStop = False
        '
        'ChkToleranceQtyResult
        '
        Me.ChkToleranceQtyResult.AutoSize = True
        Me.ChkToleranceQtyResult.Location = New System.Drawing.Point(12, 0)
        Me.ChkToleranceQtyResult.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkToleranceQtyResult.Name = "ChkToleranceQtyResult"
        Me.ChkToleranceQtyResult.Size = New System.Drawing.Size(96, 21)
        Me.ChkToleranceQtyResult.TabIndex = 12
        Me.ChkToleranceQtyResult.Text = "Qty Result"
        Me.ChkToleranceQtyResult.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(176, 59)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(20, 17)
        Me.Label17.TabIndex = 7
        Me.Label17.Text = "%"
        '
        'TxtToleranceQtyResult
        '
        Me.TxtToleranceQtyResult.Location = New System.Drawing.Point(89, 55)
        Me.TxtToleranceQtyResult.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtToleranceQtyResult.Name = "TxtToleranceQtyResult"
        Me.TxtToleranceQtyResult.Size = New System.Drawing.Size(77, 22)
        Me.TxtToleranceQtyResult.TabIndex = 14
        Me.TxtToleranceQtyResult.Text = "0.00"
        Me.TxtToleranceQtyResult.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(8, 59)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 17)
        Me.Label18.TabIndex = 5
        Me.Label18.Text = "Tolerance"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(177, 30)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(50, 17)
        Me.Label19.TabIndex = 4
        Me.Label19.Text = "Target"
        '
        'CboQtyResult
        '
        Me.CboQtyResult.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboQtyResult.FormattingEnabled = True
        Me.CboQtyResult.Items.AddRange(New Object() {"<", "=", ">"})
        Me.CboQtyResult.Location = New System.Drawing.Point(89, 25)
        Me.CboQtyResult.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboQtyResult.Name = "CboQtyResult"
        Me.CboQtyResult.Size = New System.Drawing.Size(77, 24)
        Me.CboQtyResult.TabIndex = 13
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(8, 30)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(47, 17)
        Me.Label20.TabIndex = 0
        Me.Label20.Text = "Actual"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.ChkToleranceProdRate)
        Me.GroupBox7.Controls.Add(Me.Label15)
        Me.GroupBox7.Controls.Add(Me.TxtToleranceProdRate)
        Me.GroupBox7.Controls.Add(Me.Label10)
        Me.GroupBox7.Controls.Add(Me.Label11)
        Me.GroupBox7.Controls.Add(Me.CboProdRate)
        Me.GroupBox7.Controls.Add(Me.Label12)
        Me.GroupBox7.Location = New System.Drawing.Point(732, 9)
        Me.GroupBox7.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox7.Size = New System.Drawing.Size(235, 87)
        Me.GroupBox7.TabIndex = 8
        Me.GroupBox7.TabStop = False
        '
        'ChkToleranceProdRate
        '
        Me.ChkToleranceProdRate.AutoSize = True
        Me.ChkToleranceProdRate.Location = New System.Drawing.Point(12, 0)
        Me.ChkToleranceProdRate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkToleranceProdRate.Name = "ChkToleranceProdRate"
        Me.ChkToleranceProdRate.Size = New System.Drawing.Size(132, 21)
        Me.ChkToleranceProdRate.TabIndex = 18
        Me.ChkToleranceProdRate.Text = "Production Rate"
        Me.ChkToleranceProdRate.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(177, 59)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(20, 17)
        Me.Label15.TabIndex = 8
        Me.Label15.Text = "%"
        '
        'TxtToleranceProdRate
        '
        Me.TxtToleranceProdRate.Location = New System.Drawing.Point(89, 55)
        Me.TxtToleranceProdRate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtToleranceProdRate.Name = "TxtToleranceProdRate"
        Me.TxtToleranceProdRate.Size = New System.Drawing.Size(77, 22)
        Me.TxtToleranceProdRate.TabIndex = 20
        Me.TxtToleranceProdRate.Text = "0.00"
        Me.TxtToleranceProdRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 59)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 17)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Tolerance"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(177, 30)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(50, 17)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "Target"
        '
        'CboProdRate
        '
        Me.CboProdRate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboProdRate.FormattingEnabled = True
        Me.CboProdRate.Items.AddRange(New Object() {"<", "=", ">"})
        Me.CboProdRate.Location = New System.Drawing.Point(89, 25)
        Me.CboProdRate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboProdRate.Name = "CboProdRate"
        Me.CboProdRate.Size = New System.Drawing.Size(77, 24)
        Me.CboProdRate.TabIndex = 19
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(8, 30)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(47, 17)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Actual"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.ChkToleranceGP)
        Me.GroupBox6.Controls.Add(Me.Label14)
        Me.GroupBox6.Controls.Add(Me.TxtToleranceGP)
        Me.GroupBox6.Controls.Add(Me.Label7)
        Me.GroupBox6.Controls.Add(Me.Label8)
        Me.GroupBox6.Controls.Add(Me.CboGP)
        Me.GroupBox6.Controls.Add(Me.Label9)
        Me.GroupBox6.Location = New System.Drawing.Point(489, 9)
        Me.GroupBox6.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox6.Size = New System.Drawing.Size(235, 87)
        Me.GroupBox6.TabIndex = 7
        Me.GroupBox6.TabStop = False
        '
        'ChkToleranceGP
        '
        Me.ChkToleranceGP.AutoSize = True
        Me.ChkToleranceGP.Location = New System.Drawing.Point(12, 0)
        Me.ChkToleranceGP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkToleranceGP.Name = "ChkToleranceGP"
        Me.ChkToleranceGP.Size = New System.Drawing.Size(142, 21)
        Me.ChkToleranceGP.TabIndex = 15
        Me.ChkToleranceGP.Text = "Good Percentage"
        Me.ChkToleranceGP.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(176, 59)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(20, 17)
        Me.Label14.TabIndex = 8
        Me.Label14.Text = "%"
        '
        'TxtToleranceGP
        '
        Me.TxtToleranceGP.Location = New System.Drawing.Point(89, 55)
        Me.TxtToleranceGP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtToleranceGP.Name = "TxtToleranceGP"
        Me.TxtToleranceGP.Size = New System.Drawing.Size(77, 22)
        Me.TxtToleranceGP.TabIndex = 17
        Me.TxtToleranceGP.Text = "0.00"
        Me.TxtToleranceGP.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 59)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 17)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Tolerance"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(177, 30)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 17)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Target"
        '
        'CboGP
        '
        Me.CboGP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGP.FormattingEnabled = True
        Me.CboGP.Items.AddRange(New Object() {"<", "=", ">"})
        Me.CboGP.Location = New System.Drawing.Point(89, 25)
        Me.CboGP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboGP.Name = "CboGP"
        Me.CboGP.Size = New System.Drawing.Size(77, 24)
        Me.CboGP.TabIndex = 16
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 30)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(47, 17)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Actual"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ChkToleranceQtyProduce)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.TxtToleranceQtyProduce)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.CboQtyProduce)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Location = New System.Drawing.Point(4, 9)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Size = New System.Drawing.Size(235, 87)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'ChkToleranceQtyProduce
        '
        Me.ChkToleranceQtyProduce.AutoSize = True
        Me.ChkToleranceQtyProduce.Location = New System.Drawing.Point(12, 0)
        Me.ChkToleranceQtyProduce.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkToleranceQtyProduce.Name = "ChkToleranceQtyProduce"
        Me.ChkToleranceQtyProduce.Size = New System.Drawing.Size(109, 21)
        Me.ChkToleranceQtyProduce.TabIndex = 9
        Me.ChkToleranceQtyProduce.Text = "Qty Produce"
        Me.ChkToleranceQtyProduce.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(176, 59)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(20, 17)
        Me.Label13.TabIndex = 7
        Me.Label13.Text = "%"
        '
        'TxtToleranceQtyProduce
        '
        Me.TxtToleranceQtyProduce.Location = New System.Drawing.Point(89, 55)
        Me.TxtToleranceQtyProduce.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtToleranceQtyProduce.Name = "TxtToleranceQtyProduce"
        Me.TxtToleranceQtyProduce.Size = New System.Drawing.Size(77, 22)
        Me.TxtToleranceQtyProduce.TabIndex = 11
        Me.TxtToleranceQtyProduce.Text = "0.00"
        Me.TxtToleranceQtyProduce.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 59)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 17)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Tolerance"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(177, 30)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(50, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Target"
        '
        'CboQtyProduce
        '
        Me.CboQtyProduce.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboQtyProduce.FormattingEnabled = True
        Me.CboQtyProduce.Items.AddRange(New Object() {"<", "=", ">"})
        Me.CboQtyProduce.Location = New System.Drawing.Point(89, 25)
        Me.CboQtyProduce.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboQtyProduce.Name = "CboQtyProduce"
        Me.CboQtyProduce.Size = New System.Drawing.Size(77, 24)
        Me.CboQtyProduce.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 30)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(47, 17)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Actual"
        '
        'BgwProduct
        '
        '
        'BgwMO
        '
        '
        'FrmLaporanProd_Productivity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(1333, 692)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(1341, 729)
        Me.Name = "FrmLaporanProd_Productivity"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Productivity Report"
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgvMO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents colMO_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colMO_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colMO_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents colMO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DgvMO As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents CboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BgwProduct As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwMO As System.ComponentModel.BackgroundWorker
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkToleranceQtyResult As System.Windows.Forms.CheckBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtToleranceQtyResult As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents CboQtyResult As System.Windows.Forms.ComboBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkToleranceProdRate As System.Windows.Forms.CheckBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtToleranceProdRate As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents CboProdRate As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkToleranceGP As System.Windows.Forms.CheckBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtToleranceGP As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CboGP As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkToleranceQtyProduce As System.Windows.Forms.CheckBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtToleranceQtyProduce As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CboQtyProduce As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
