<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransMO_WorkShift_Allocate
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.rboWUG = New System.Windows.Forms.RadioButton
        Me.gbo_cheklist = New System.Windows.Forms.GroupBox
        Me.llbUnAllocate = New System.Windows.Forms.LinkLabel
        Me.llbAllocate = New System.Windows.Forms.LinkLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbo_shift = New System.Windows.Forms.ComboBox
        Me.GB_WUG = New System.Windows.Forms.GroupBox
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.ClbWU = New System.Windows.Forms.CheckedListBox
        Me.dgvWUG = New System.Windows.Forms.DataGridView
        Me.colWUG_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWUG_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWUG_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWUG_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.rboWU = New System.Windows.Forms.RadioButton
        Me.GB_WU = New System.Windows.Forms.GroupBox
        Me.dgvWU = New System.Windows.Forms.DataGridView
        Me.colWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFilter = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWorkUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWorkUnitDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTglDari = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTglSampai = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button2 = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gbo_cheklist.SuspendLayout()
        Me.GB_WUG.SuspendLayout()
        CType(Me.dgvWUG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB_WU.SuspendLayout()
        CType(Me.dgvWU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.ProgressBar1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rboWUG)
        Me.SplitContainer1.Panel1.Controls.Add(Me.gbo_cheklist)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GB_WUG)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rboWU)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GB_WU)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Button1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvItem)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1283, 663)
        Me.SplitContainer1.SplitterDistance = 371
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 630)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(240, 23)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 0
        Me.ProgressBar1.Visible = False
        '
        'rboWUG
        '
        Me.rboWUG.AutoSize = True
        Me.rboWUG.Location = New System.Drawing.Point(139, 69)
        Me.rboWUG.Margin = New System.Windows.Forms.Padding(4)
        Me.rboWUG.Name = "rboWUG"
        Me.rboWUG.Size = New System.Drawing.Size(155, 21)
        Me.rboWUG.TabIndex = 3
        Me.rboWUG.Text = "By Work Unit Group"
        Me.rboWUG.UseVisualStyleBackColor = True
        '
        'gbo_cheklist
        '
        Me.gbo_cheklist.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbo_cheklist.Controls.Add(Me.llbUnAllocate)
        Me.gbo_cheklist.Controls.Add(Me.llbAllocate)
        Me.gbo_cheklist.Controls.Add(Me.Label3)
        Me.gbo_cheklist.Controls.Add(Me.cbo_shift)
        Me.gbo_cheklist.Location = New System.Drawing.Point(7, 546)
        Me.gbo_cheklist.Margin = New System.Windows.Forms.Padding(4)
        Me.gbo_cheklist.Name = "gbo_cheklist"
        Me.gbo_cheklist.Padding = New System.Windows.Forms.Padding(4)
        Me.gbo_cheklist.Size = New System.Drawing.Size(358, 78)
        Me.gbo_cheklist.TabIndex = 10
        Me.gbo_cheklist.TabStop = False
        '
        'llbUnAllocate
        '
        Me.llbUnAllocate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUnAllocate.AutoSize = True
        Me.llbUnAllocate.Location = New System.Drawing.Point(246, 52)
        Me.llbUnAllocate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbUnAllocate.Name = "llbUnAllocate"
        Me.llbUnAllocate.Size = New System.Drawing.Size(94, 17)
        Me.llbUnAllocate.TabIndex = 5
        Me.llbUnAllocate.TabStop = True
        Me.llbUnAllocate.Text = "Unallocate All"
        '
        'llbAllocate
        '
        Me.llbAllocate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAllocate.AutoSize = True
        Me.llbAllocate.Location = New System.Drawing.Point(159, 52)
        Me.llbAllocate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbAllocate.Name = "llbAllocate"
        Me.llbAllocate.Size = New System.Drawing.Size(77, 17)
        Me.llbAllocate.TabIndex = 4
        Me.llbAllocate.TabStop = True
        Me.llbAllocate.Text = "Allocate All"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 27)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Work shift"
        '
        'cbo_shift
        '
        Me.cbo_shift.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbo_shift.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo_shift.FormattingEnabled = True
        Me.cbo_shift.Location = New System.Drawing.Point(89, 23)
        Me.cbo_shift.Margin = New System.Windows.Forms.Padding(4)
        Me.cbo_shift.Name = "cbo_shift"
        Me.cbo_shift.Size = New System.Drawing.Size(251, 24)
        Me.cbo_shift.TabIndex = 2
        '
        'GB_WUG
        '
        Me.GB_WUG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GB_WUG.Controls.Add(Me.llbUncheckAll)
        Me.GB_WUG.Controls.Add(Me.llbCheckAll)
        Me.GB_WUG.Controls.Add(Me.ClbWU)
        Me.GB_WUG.Controls.Add(Me.dgvWUG)
        Me.GB_WUG.Location = New System.Drawing.Point(7, 271)
        Me.GB_WUG.Margin = New System.Windows.Forms.Padding(4)
        Me.GB_WUG.Name = "GB_WUG"
        Me.GB_WUG.Padding = New System.Windows.Forms.Padding(4)
        Me.GB_WUG.Size = New System.Drawing.Size(358, 268)
        Me.GB_WUG.TabIndex = 2
        Me.GB_WUG.TabStop = False
        Me.GB_WUG.Text = "Work Unit Group"
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(267, 176)
        Me.llbUncheckAll.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncheckAll.TabIndex = 7
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(187, 176)
        Me.llbCheckAll.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCheckAll.TabIndex = 6
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'ClbWU
        '
        Me.ClbWU.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClbWU.CheckOnClick = True
        Me.ClbWU.Enabled = False
        Me.ClbWU.FormattingEnabled = True
        Me.ClbWU.Location = New System.Drawing.Point(3, 193)
        Me.ClbWU.Margin = New System.Windows.Forms.Padding(4)
        Me.ClbWU.MultiColumn = True
        Me.ClbWU.Name = "ClbWU"
        Me.ClbWU.Size = New System.Drawing.Size(346, 55)
        Me.ClbWU.TabIndex = 8
        '
        'dgvWUG
        '
        Me.dgvWUG.AllowUserToResizeRows = False
        Me.dgvWUG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWUG.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWUG.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWUG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWUG.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWUG_No, Me.colWUG_Code, Me.colWUG_Desc, Me.colWUG_Primary})
        Me.dgvWUG.Location = New System.Drawing.Point(4, 23)
        Me.dgvWUG.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvWUG.MultiSelect = False
        Me.dgvWUG.Name = "dgvWUG"
        Me.dgvWUG.RowTemplate.Height = 24
        Me.dgvWUG.Size = New System.Drawing.Size(350, 149)
        Me.dgvWUG.TabIndex = 5
        '
        'colWUG_No
        '
        Me.colWUG_No.HeaderText = "No"
        Me.colWUG_No.Name = "colWUG_No"
        Me.colWUG_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colWUG_No.Width = 32
        '
        'colWUG_Code
        '
        Me.colWUG_Code.HeaderText = "Work Unit"
        Me.colWUG_Code.Name = "colWUG_Code"
        Me.colWUG_Code.Width = 95
        '
        'colWUG_Desc
        '
        Me.colWUG_Desc.HeaderText = "Desc"
        Me.colWUG_Desc.Name = "colWUG_Desc"
        Me.colWUG_Desc.Width = 65
        '
        'colWUG_Primary
        '
        Me.colWUG_Primary.HeaderText = "Primary WU"
        Me.colWUG_Primary.Name = "colWUG_Primary"
        Me.colWUG_Primary.Visible = False
        Me.colWUG_Primary.Width = 88
        '
        'rboWU
        '
        Me.rboWU.AutoSize = True
        Me.rboWU.Checked = True
        Me.rboWU.Location = New System.Drawing.Point(11, 69)
        Me.rboWU.Margin = New System.Windows.Forms.Padding(4)
        Me.rboWU.Name = "rboWU"
        Me.rboWU.Size = New System.Drawing.Size(111, 21)
        Me.rboWU.TabIndex = 2
        Me.rboWU.TabStop = True
        Me.rboWU.Text = "By Work Unit"
        Me.rboWU.UseVisualStyleBackColor = True
        '
        'GB_WU
        '
        Me.GB_WU.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GB_WU.Controls.Add(Me.dgvWU)
        Me.GB_WU.Location = New System.Drawing.Point(3, 97)
        Me.GB_WU.Margin = New System.Windows.Forms.Padding(4)
        Me.GB_WU.Name = "GB_WU"
        Me.GB_WU.Padding = New System.Windows.Forms.Padding(4)
        Me.GB_WU.Size = New System.Drawing.Size(358, 166)
        Me.GB_WU.TabIndex = 1
        Me.GB_WU.TabStop = False
        Me.GB_WU.Text = "Work Unit"
        '
        'dgvWU
        '
        Me.dgvWU.AllowUserToResizeRows = False
        Me.dgvWU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWU.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWU_No, Me.colWU_Code, Me.colWU_Desc, Me.colWU_Primary})
        Me.dgvWU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvWU.Location = New System.Drawing.Point(4, 19)
        Me.dgvWU.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvWU.MultiSelect = False
        Me.dgvWU.Name = "dgvWU"
        Me.dgvWU.RowTemplate.Height = 24
        Me.dgvWU.Size = New System.Drawing.Size(350, 143)
        Me.dgvWU.TabIndex = 4
        '
        'colWU_No
        '
        Me.colWU_No.HeaderText = "No"
        Me.colWU_No.Name = "colWU_No"
        Me.colWU_No.ReadOnly = True
        Me.colWU_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colWU_No.Width = 32
        '
        'colWU_Code
        '
        Me.colWU_Code.HeaderText = "Work Unit"
        Me.colWU_Code.Name = "colWU_Code"
        Me.colWU_Code.Width = 95
        '
        'colWU_Desc
        '
        Me.colWU_Desc.HeaderText = "Desc"
        Me.colWU_Desc.Name = "colWU_Desc"
        Me.colWU_Desc.ReadOnly = True
        Me.colWU_Desc.Width = 65
        '
        'colWU_Primary
        '
        Me.colWU_Primary.HeaderText = "Primary WU"
        Me.colWU_Primary.Name = "colWU_Primary"
        Me.colWU_Primary.ReadOnly = True
        Me.colWU_Primary.Visible = False
        Me.colWU_Primary.Width = 88
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer2)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 4)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(358, 57)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Date"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpFrom)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.dtpTo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer2.Size = New System.Drawing.Size(348, 33)
        Me.SplitContainer2.SplitterDistance = 165
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 3
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(39, 9)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(121, 22)
        Me.dtpFrom.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-4, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(40, 9)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(122, 22)
        Me.dtpTo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 12)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(255, 628)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(109, 26)
        Me.btnFilter.TabIndex = 9
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(263, 628)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(92, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colWorkUnitCode, Me.colWorkUnitDesc, Me.colDate, Me.colTglDari, Me.colTglSampai, Me.colPrimaryWU})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.Size = New System.Drawing.Size(907, 663)
        Me.dgvItem.TabIndex = 10
        '
        'colNo
        '
        Me.colNo.Frozen = True
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colNo.Width = 32
        '
        'colWorkUnitCode
        '
        Me.colWorkUnitCode.Frozen = True
        Me.colWorkUnitCode.HeaderText = "Work Unit"
        Me.colWorkUnitCode.Name = "colWorkUnitCode"
        Me.colWorkUnitCode.ReadOnly = True
        Me.colWorkUnitCode.Width = 95
        '
        'colWorkUnitDesc
        '
        Me.colWorkUnitDesc.HeaderText = "Desc"
        Me.colWorkUnitDesc.Name = "colWorkUnitDesc"
        Me.colWorkUnitDesc.ReadOnly = True
        Me.colWorkUnitDesc.Width = 65
        '
        'colDate
        '
        Me.colDate.HeaderText = "Date"
        Me.colDate.Name = "colDate"
        Me.colDate.ReadOnly = True
        Me.colDate.Width = 63
        '
        'colTglDari
        '
        Me.colTglDari.HeaderText = "Tanggal Dari"
        Me.colTglDari.Name = "colTglDari"
        Me.colTglDari.Visible = False
        Me.colTglDari.Width = 93
        '
        'colTglSampai
        '
        Me.colTglSampai.HeaderText = "Tanggal Sampai"
        Me.colTglSampai.Name = "colTglSampai"
        Me.colTglSampai.Visible = False
        Me.colTglSampai.Width = 109
        '
        'colPrimaryWU
        '
        Me.colPrimaryWU.HeaderText = "Primary WU"
        Me.colPrimaryWU.Name = "colPrimaryWU"
        Me.colPrimaryWU.ReadOnly = True
        Me.colPrimaryWU.Visible = False
        Me.colPrimaryWU.Width = 88
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(108, 631)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Filter"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'FrmTransMO_WorkShift_Allocate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(1283, 663)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(1291, 700)
        Me.Name = "FrmTransMO_WorkShift_Allocate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Unit Workshift Allocation"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.gbo_cheklist.ResumeLayout(False)
        Me.gbo_cheklist.PerformLayout()
        Me.GB_WUG.ResumeLayout(False)
        Me.GB_WUG.PerformLayout()
        CType(Me.dgvWUG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB_WU.ResumeLayout(False)
        CType(Me.dgvWU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GB_WUG As System.Windows.Forms.GroupBox
    Friend WithEvents rboWUG As System.Windows.Forms.RadioButton
    Friend WithEvents rboWU As System.Windows.Forms.RadioButton
    Friend WithEvents GB_WU As System.Windows.Forms.GroupBox
    Friend WithEvents dgvWU As System.Windows.Forms.DataGridView
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents dgvWUG As System.Windows.Forms.DataGridView
    Friend WithEvents ClbWU As System.Windows.Forms.CheckedListBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents colWUG_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWUG_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWUG_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWUG_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents gbo_cheklist As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbo_shift As System.Windows.Forms.ComboBox
    Friend WithEvents llbUnAllocate As System.Windows.Forms.LinkLabel
    Friend WithEvents llbAllocate As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWorkUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWorkUnitDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTglDari As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTglSampai As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
