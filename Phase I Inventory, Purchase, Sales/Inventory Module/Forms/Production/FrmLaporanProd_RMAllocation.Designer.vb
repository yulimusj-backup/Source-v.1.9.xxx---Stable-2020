<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_RMAllocation
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboStatus = New System.Windows.Forms.ComboBox
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnShow = New System.Windows.Forms.Button
        Me.dtpWOStart = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.dgvProcess = New System.Windows.Forms.DataGridView
        Me.colProcess_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.chkWorkUnit = New System.Windows.Forms.CheckedListBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.rboWUG = New System.Windows.Forms.RadioButton
        Me.dgvWUG = New System.Windows.Forms.DataGridView
        Me.colWUG_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWUG_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWUG_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWUG_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.rboWU = New System.Windows.Forms.RadioButton
        Me.dgvWorkUnit = New System.Windows.Forms.DataGridView
        Me.colWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnClose = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dgvProcess, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.dgvWUG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dgvWorkUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnShow)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpWOStart)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1189, 693)
        Me.SplitContainer1.SplitterDistance = 285
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 74)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Stock Status"
        '
        'cboStatus
        '
        Me.cboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Items.AddRange(New Object() {"All", "Not Ready Only"})
        Me.cboStatus.Location = New System.Drawing.Point(107, 70)
        Me.cboStatus.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(167, 24)
        Me.cboStatus.TabIndex = 7
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy HH.mm"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(107, 38)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(167, 22)
        Me.dtpTo.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 43)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "WO To"
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(175, 103)
        Me.btnShow.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(100, 28)
        Me.btnShow.TabIndex = 1
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'dtpWOStart
        '
        Me.dtpWOStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpWOStart.CustomFormat = "dd MMM yyyy HH.mm"
        Me.dtpWOStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpWOStart.Location = New System.Drawing.Point(107, 6)
        Me.dtpWOStart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpWOStart.Name = "dtpWOStart"
        Me.dtpWOStart.Size = New System.Drawing.Size(167, 22)
        Me.dtpWOStart.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "WO From"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.dgvProcess)
        Me.GroupBox3.Location = New System.Drawing.Point(5, 420)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Size = New System.Drawing.Size(274, 240)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Process"
        '
        'dgvProcess
        '
        Me.dgvProcess.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProcess.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvProcess.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProcess.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProcess_No, Me.colProcess_Code, Me.colProcess_Desc, Me.colProcess_Primary})
        Me.dgvProcess.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProcess.Location = New System.Drawing.Point(4, 19)
        Me.dgvProcess.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvProcess.Name = "dgvProcess"
        Me.dgvProcess.RowTemplate.Height = 24
        Me.dgvProcess.Size = New System.Drawing.Size(266, 217)
        Me.dgvProcess.TabIndex = 0
        '
        'colProcess_No
        '
        Me.colProcess_No.HeaderText = "No"
        Me.colProcess_No.Name = "colProcess_No"
        Me.colProcess_No.Width = 51
        '
        'colProcess_Code
        '
        Me.colProcess_Code.HeaderText = "Process Code"
        Me.colProcess_Code.Name = "colProcess_Code"
        Me.colProcess_Code.Width = 121
        '
        'colProcess_Desc
        '
        Me.colProcess_Desc.HeaderText = "Desc"
        Me.colProcess_Desc.Name = "colProcess_Desc"
        Me.colProcess_Desc.Width = 65
        '
        'colProcess_Primary
        '
        Me.colProcess_Primary.HeaderText = "Primary"
        Me.colProcess_Primary.Name = "colProcess_Primary"
        Me.colProcess_Primary.Visible = False
        Me.colProcess_Primary.Width = 66
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvProduct)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 139)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(274, 274)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Raw Material"
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
        Me.dgvProduct.Size = New System.Drawing.Size(266, 251)
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
        Me.colProduct_Desc.Width = 65
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 66
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox6)
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox5)
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox4)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.crv)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Size = New System.Drawing.Size(899, 693)
        Me.SplitContainer2.SplitterDistance = 143
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 0
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox6.Controls.Add(Me.chkWorkUnit)
        Me.GroupBox6.Location = New System.Drawing.Point(641, 11)
        Me.GroupBox6.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox6.Size = New System.Drawing.Size(241, 129)
        Me.GroupBox6.TabIndex = 7
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Work Unit"
        '
        'chkWorkUnit
        '
        Me.chkWorkUnit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkWorkUnit.Enabled = False
        Me.chkWorkUnit.FormattingEnabled = True
        Me.chkWorkUnit.Location = New System.Drawing.Point(4, 19)
        Me.chkWorkUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkWorkUnit.Name = "chkWorkUnit"
        Me.chkWorkUnit.Size = New System.Drawing.Size(233, 106)
        Me.chkWorkUnit.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.rboWUG)
        Me.GroupBox5.Controls.Add(Me.dgvWUG)
        Me.GroupBox5.Location = New System.Drawing.Point(323, 11)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox5.Size = New System.Drawing.Size(311, 129)
        Me.GroupBox5.TabIndex = 6
        Me.GroupBox5.TabStop = False
        '
        'rboWUG
        '
        Me.rboWUG.AutoSize = True
        Me.rboWUG.Location = New System.Drawing.Point(8, -2)
        Me.rboWUG.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rboWUG.Name = "rboWUG"
        Me.rboWUG.Size = New System.Drawing.Size(135, 21)
        Me.rboWUG.TabIndex = 7
        Me.rboWUG.TabStop = True
        Me.rboWUG.Text = "Work Unit Group"
        Me.rboWUG.UseVisualStyleBackColor = True
        '
        'dgvWUG
        '
        Me.dgvWUG.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWUG.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWUG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWUG.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWUG_No, Me.colWUG_Code, Me.colWUG_Desc, Me.colWUG_Primary})
        Me.dgvWUG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvWUG.Location = New System.Drawing.Point(4, 19)
        Me.dgvWUG.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvWUG.Name = "dgvWUG"
        Me.dgvWUG.ReadOnly = True
        Me.dgvWUG.RowTemplate.Height = 24
        Me.dgvWUG.Size = New System.Drawing.Size(303, 106)
        Me.dgvWUG.TabIndex = 0
        '
        'colWUG_No
        '
        Me.colWUG_No.HeaderText = "No"
        Me.colWUG_No.Name = "colWUG_No"
        Me.colWUG_No.ReadOnly = True
        Me.colWUG_No.Width = 51
        '
        'colWUG_Code
        '
        Me.colWUG_Code.HeaderText = "Code"
        Me.colWUG_Code.Name = "colWUG_Code"
        Me.colWUG_Code.ReadOnly = True
        Me.colWUG_Code.Width = 66
        '
        'colWUG_Desc
        '
        Me.colWUG_Desc.HeaderText = "Desc"
        Me.colWUG_Desc.Name = "colWUG_Desc"
        Me.colWUG_Desc.ReadOnly = True
        Me.colWUG_Desc.Width = 65
        '
        'colWUG_Primary
        '
        Me.colWUG_Primary.HeaderText = "Primary"
        Me.colWUG_Primary.Name = "colWUG_Primary"
        Me.colWUG_Primary.ReadOnly = True
        Me.colWUG_Primary.Visible = False
        Me.colWUG_Primary.Width = 66
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.rboWU)
        Me.GroupBox4.Controls.Add(Me.dgvWorkUnit)
        Me.GroupBox4.Location = New System.Drawing.Point(4, 11)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(311, 129)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        '
        'rboWU
        '
        Me.rboWU.AutoSize = True
        Me.rboWU.Checked = True
        Me.rboWU.Location = New System.Drawing.Point(8, -2)
        Me.rboWU.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rboWU.Name = "rboWU"
        Me.rboWU.Size = New System.Drawing.Size(91, 21)
        Me.rboWU.TabIndex = 6
        Me.rboWU.TabStop = True
        Me.rboWU.Text = "Work Unit"
        Me.rboWU.UseVisualStyleBackColor = True
        '
        'dgvWorkUnit
        '
        Me.dgvWorkUnit.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWorkUnit.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWorkUnit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWorkUnit.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWU_No, Me.colWU_Code, Me.colWU_Desc, Me.colWU_Primary})
        Me.dgvWorkUnit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvWorkUnit.Location = New System.Drawing.Point(4, 19)
        Me.dgvWorkUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvWorkUnit.Name = "dgvWorkUnit"
        Me.dgvWorkUnit.RowTemplate.Height = 24
        Me.dgvWorkUnit.Size = New System.Drawing.Size(303, 106)
        Me.dgvWorkUnit.TabIndex = 0
        '
        'colWU_No
        '
        Me.colWU_No.HeaderText = "No"
        Me.colWU_No.Name = "colWU_No"
        Me.colWU_No.Width = 51
        '
        'colWU_Code
        '
        Me.colWU_Code.HeaderText = "Code"
        Me.colWU_Code.Name = "colWU_Code"
        Me.colWU_Code.Width = 66
        '
        'colWU_Desc
        '
        Me.colWU_Desc.HeaderText = "Desc"
        Me.colWU_Desc.Name = "colWU_Desc"
        Me.colWU_Desc.Width = 65
        '
        'colWU_Primary
        '
        Me.colWU_Primary.HeaderText = "Primary"
        Me.colWU_Primary.Name = "colWU_Primary"
        Me.colWU_Primary.Visible = False
        Me.colWU_Primary.Width = 66
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(0, 0)
        Me.crv.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(899, 545)
        Me.crv.TabIndex = 0
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(783, 483)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Show"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker4
        '
        '
        'BackgroundWorker5
        '
        '
        'FrmLaporanProd_RMAllocation
        '
        Me.AcceptButton = Me.btnShow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(1189, 693)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(1197, 730)
        Me.Name = "FrmLaporanProd_RMAllocation"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pick List Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.dgvProcess, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.dgvWUG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.dgvWorkUnit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvProcess As System.Windows.Forms.DataGridView
    Friend WithEvents colProcess_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dtpWOStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvWorkUnit As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvWUG As System.Windows.Forms.DataGridView
    Friend WithEvents colWUG_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWUG_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWUG_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWUG_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents chkWorkUnit As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents rboWUG As System.Windows.Forms.RadioButton
    Friend WithEvents rboWU As System.Windows.Forms.RadioButton
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents colWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
End Class
