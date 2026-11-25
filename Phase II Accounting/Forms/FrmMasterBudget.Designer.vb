<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterBudget
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
        Dim DataGridViewCellStyle22 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle23 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle24 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle25 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle26 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle27 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle28 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle29 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle30 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle31 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle32 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle33 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle34 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle35 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle36 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle37 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle38 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle39 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle40 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle41 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle42 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvBudget = New System.Windows.Forms.DataGridView
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFeb = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMar = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColApr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJune = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJul = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAgt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNov = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLevel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkYear1_Bud = New System.Windows.Forms.CheckBox
        Me.chkYear1_Rea = New System.Windows.Forms.CheckBox
        Me.chkYear2_Rea = New System.Windows.Forms.CheckBox
        Me.chkYear2_Bud = New System.Windows.Forms.CheckBox
        Me.chkYear3_Rea = New System.Windows.Forms.CheckBox
        Me.chkYear3_Bud = New System.Windows.Forms.CheckBox
        Me.chkNextYear = New System.Windows.Forms.CheckBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnShowDetail = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.cboFilter = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dgvYearly = New System.Windows.Forms.DataGridView
        Me.ColAccount_Yearly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYear3_Bud = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYear3_Rea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYear2_Budget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYear2_Rea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYear1_Bud = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYear1_Rea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurYear_Budget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrenctYear_Rea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNextYear_Bud = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkunYearly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        CType(Me.dgvBudget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvYearly, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvBudget
        '
        Me.dgvBudget.AllowUserToAddRows = False
        Me.dgvBudget.AllowUserToDeleteRows = False
        Me.dgvBudget.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvBudget.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvBudget.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColJan, Me.ColFeb, Me.ColMar, Me.ColApr, Me.ColMay, Me.ColJune, Me.ColJul, Me.ColAgt, Me.ColSept, Me.ColOct, Me.ColNov, Me.ColDec, Me.ColRemarks, Me.ColPrimaryAkun, Me.ColLevel, Me.ColHeader})
        Me.dgvBudget.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvBudget.Location = New System.Drawing.Point(0, 0)
        Me.dgvBudget.Name = "dgvBudget"
        Me.dgvBudget.Size = New System.Drawing.Size(697, 517)
        Me.dgvBudget.TabIndex = 0
        '
        'ColAccount
        '
        Me.ColAccount.Frozen = True
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount.Width = 53
        '
        'ColJan
        '
        DataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJan.DefaultCellStyle = DataGridViewCellStyle22
        Me.ColJan.HeaderText = "January"
        Me.ColJan.Name = "ColJan"
        Me.ColJan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJan.Width = 50
        '
        'ColFeb
        '
        DataGridViewCellStyle23.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColFeb.DefaultCellStyle = DataGridViewCellStyle23
        Me.ColFeb.HeaderText = "February"
        Me.ColFeb.Name = "ColFeb"
        Me.ColFeb.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFeb.Width = 54
        '
        'ColMar
        '
        DataGridViewCellStyle24.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMar.DefaultCellStyle = DataGridViewCellStyle24
        Me.ColMar.HeaderText = "March"
        Me.ColMar.Name = "ColMar"
        Me.ColMar.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMar.Width = 43
        '
        'ColApr
        '
        DataGridViewCellStyle25.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColApr.DefaultCellStyle = DataGridViewCellStyle25
        Me.ColApr.HeaderText = "April"
        Me.ColApr.Name = "ColApr"
        Me.ColApr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColApr.Width = 33
        '
        'ColMay
        '
        DataGridViewCellStyle26.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMay.DefaultCellStyle = DataGridViewCellStyle26
        Me.ColMay.HeaderText = "May"
        Me.ColMay.Name = "ColMay"
        Me.ColMay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMay.Width = 33
        '
        'ColJune
        '
        DataGridViewCellStyle27.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJune.DefaultCellStyle = DataGridViewCellStyle27
        Me.ColJune.HeaderText = "June"
        Me.ColJune.Name = "ColJune"
        Me.ColJune.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJune.Width = 36
        '
        'ColJul
        '
        DataGridViewCellStyle28.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJul.DefaultCellStyle = DataGridViewCellStyle28
        Me.ColJul.HeaderText = "July"
        Me.ColJul.Name = "ColJul"
        Me.ColJul.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJul.Width = 31
        '
        'ColAgt
        '
        DataGridViewCellStyle29.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAgt.DefaultCellStyle = DataGridViewCellStyle29
        Me.ColAgt.HeaderText = "August"
        Me.ColAgt.Name = "ColAgt"
        Me.ColAgt.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAgt.Width = 46
        '
        'ColSept
        '
        DataGridViewCellStyle30.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSept.DefaultCellStyle = DataGridViewCellStyle30
        Me.ColSept.HeaderText = "September"
        Me.ColSept.Name = "ColSept"
        Me.ColSept.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSept.Width = 64
        '
        'ColOct
        '
        DataGridViewCellStyle31.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOct.DefaultCellStyle = DataGridViewCellStyle31
        Me.ColOct.HeaderText = "October"
        Me.ColOct.Name = "ColOct"
        Me.ColOct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOct.Width = 51
        '
        'ColNov
        '
        DataGridViewCellStyle32.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNov.DefaultCellStyle = DataGridViewCellStyle32
        Me.ColNov.HeaderText = "November"
        Me.ColNov.Name = "ColNov"
        Me.ColNov.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNov.Width = 62
        '
        'ColDec
        '
        DataGridViewCellStyle33.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDec.DefaultCellStyle = DataGridViewCellStyle33
        Me.ColDec.HeaderText = "December"
        Me.ColDec.Name = "ColDec"
        Me.ColDec.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDec.Width = 62
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "Primary Acc"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 88
        '
        'ColLevel
        '
        Me.ColLevel.HeaderText = "Level"
        Me.ColLevel.Name = "ColLevel"
        Me.ColLevel.Visible = False
        Me.ColLevel.Width = 58
        '
        'ColHeader
        '
        Me.ColHeader.HeaderText = "Header"
        Me.ColHeader.Name = "ColHeader"
        Me.ColHeader.Visible = False
        Me.ColHeader.Width = 67
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(821, 237)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Financial Year"
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(96, 9)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.ShowUpDown = True
        Me.dtpPeriod.Size = New System.Drawing.Size(124, 20)
        Me.dtpPeriod.TabIndex = 3
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Monthly", "Yearly"})
        Me.cboFormat.Location = New System.Drawing.Point(96, 35)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(124, 21)
        Me.cboFormat.TabIndex = 4
        Me.cboFormat.TabStop = False
        Me.cboFormat.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Format"
        Me.Label2.Visible = False
        '
        'chkYear1_Bud
        '
        Me.chkYear1_Bud.AutoSize = True
        Me.chkYear1_Bud.Location = New System.Drawing.Point(19, 91)
        Me.chkYear1_Bud.Name = "chkYear1_Bud"
        Me.chkYear1_Bud.Size = New System.Drawing.Size(155, 17)
        Me.chkYear1_Bud.TabIndex = 6
        Me.chkYear1_Bud.Text = "Show 1 year before Budget"
        Me.chkYear1_Bud.UseVisualStyleBackColor = True
        Me.chkYear1_Bud.Visible = False
        '
        'chkYear1_Rea
        '
        Me.chkYear1_Rea.AutoSize = True
        Me.chkYear1_Rea.Location = New System.Drawing.Point(19, 113)
        Me.chkYear1_Rea.Name = "chkYear1_Rea"
        Me.chkYear1_Rea.Size = New System.Drawing.Size(173, 17)
        Me.chkYear1_Rea.TabIndex = 7
        Me.chkYear1_Rea.Text = "Show 1 year before Realization"
        Me.chkYear1_Rea.UseVisualStyleBackColor = True
        Me.chkYear1_Rea.Visible = False
        '
        'chkYear2_Rea
        '
        Me.chkYear2_Rea.AutoSize = True
        Me.chkYear2_Rea.Location = New System.Drawing.Point(19, 158)
        Me.chkYear2_Rea.Name = "chkYear2_Rea"
        Me.chkYear2_Rea.Size = New System.Drawing.Size(173, 17)
        Me.chkYear2_Rea.TabIndex = 9
        Me.chkYear2_Rea.Text = "Show 2 year before Realization"
        Me.chkYear2_Rea.UseVisualStyleBackColor = True
        Me.chkYear2_Rea.Visible = False
        '
        'chkYear2_Bud
        '
        Me.chkYear2_Bud.AutoSize = True
        Me.chkYear2_Bud.Location = New System.Drawing.Point(19, 136)
        Me.chkYear2_Bud.Name = "chkYear2_Bud"
        Me.chkYear2_Bud.Size = New System.Drawing.Size(155, 17)
        Me.chkYear2_Bud.TabIndex = 8
        Me.chkYear2_Bud.Text = "Show 2 year before Budget"
        Me.chkYear2_Bud.UseVisualStyleBackColor = True
        Me.chkYear2_Bud.Visible = False
        '
        'chkYear3_Rea
        '
        Me.chkYear3_Rea.AutoSize = True
        Me.chkYear3_Rea.Location = New System.Drawing.Point(19, 203)
        Me.chkYear3_Rea.Name = "chkYear3_Rea"
        Me.chkYear3_Rea.Size = New System.Drawing.Size(173, 17)
        Me.chkYear3_Rea.TabIndex = 11
        Me.chkYear3_Rea.Text = "Show 3 year before Realization"
        Me.chkYear3_Rea.UseVisualStyleBackColor = True
        Me.chkYear3_Rea.Visible = False
        '
        'chkYear3_Bud
        '
        Me.chkYear3_Bud.AutoSize = True
        Me.chkYear3_Bud.Location = New System.Drawing.Point(19, 181)
        Me.chkYear3_Bud.Name = "chkYear3_Bud"
        Me.chkYear3_Bud.Size = New System.Drawing.Size(155, 17)
        Me.chkYear3_Bud.TabIndex = 10
        Me.chkYear3_Bud.Text = "Show 3 year before Budget"
        Me.chkYear3_Bud.UseVisualStyleBackColor = True
        Me.chkYear3_Bud.Visible = False
        '
        'chkNextYear
        '
        Me.chkNextYear.AutoSize = True
        Me.chkNextYear.Location = New System.Drawing.Point(19, 226)
        Me.chkNextYear.Name = "chkNextYear"
        Me.chkNextYear.Size = New System.Drawing.Size(146, 17)
        Me.chkNextYear.TabIndex = 12
        Me.chkNextYear.Text = "Show 1 year after Budget"
        Me.chkNextYear.UseVisualStyleBackColor = True
        Me.chkNextYear.Visible = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnShowDetail)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkNextYear)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpPeriod)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkYear3_Rea)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboFormat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkYear3_Bud)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkYear2_Rea)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkYear1_Bud)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkYear2_Bud)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkYear1_Rea)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvBudget)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvYearly)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Size = New System.Drawing.Size(931, 519)
        Me.SplitContainer1.SplitterDistance = 228
        Me.SplitContainer1.TabIndex = 13
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(19, 91)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(71, 23)
        Me.btnDelete.TabIndex = 17
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnShowDetail
        '
        Me.btnShowDetail.Location = New System.Drawing.Point(96, 62)
        Me.btnShowDetail.Name = "btnShowDetail"
        Me.btnShowDetail.Size = New System.Drawing.Size(78, 23)
        Me.btnShowDetail.TabIndex = 16
        Me.btnShowDetail.Text = "Show Detail"
        Me.btnShowDetail.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(19, 62)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(71, 23)
        Me.btnSave.TabIndex = 15
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'cboFilter
        '
        Me.cboFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFilter.FormattingEnabled = True
        Me.cboFilter.Items.AddRange(New Object() {"Balance Sheet", "Profit And Loss"})
        Me.cboFilter.Location = New System.Drawing.Point(96, 35)
        Me.cboFilter.Name = "cboFilter"
        Me.cboFilter.Size = New System.Drawing.Size(124, 21)
        Me.cboFilter.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Filter"
        '
        'dgvYearly
        '
        Me.dgvYearly.AllowUserToAddRows = False
        Me.dgvYearly.AllowUserToDeleteRows = False
        Me.dgvYearly.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvYearly.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvYearly.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount_Yearly, Me.ColYear3_Bud, Me.ColYear3_Rea, Me.ColYear2_Budget, Me.ColYear2_Rea, Me.ColYear1_Bud, Me.ColYear1_Rea, Me.ColCurYear_Budget, Me.ColCurrenctYear_Rea, Me.ColNextYear_Bud, Me.ColPrimaryAkunYearly})
        Me.dgvYearly.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvYearly.Location = New System.Drawing.Point(0, 0)
        Me.dgvYearly.Name = "dgvYearly"
        Me.dgvYearly.Size = New System.Drawing.Size(697, 517)
        Me.dgvYearly.TabIndex = 1
        '
        'ColAccount_Yearly
        '
        Me.ColAccount_Yearly.Frozen = True
        Me.ColAccount_Yearly.HeaderText = "Account"
        Me.ColAccount_Yearly.Name = "ColAccount_Yearly"
        Me.ColAccount_Yearly.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount_Yearly.Width = 53
        '
        'ColYear3_Bud
        '
        DataGridViewCellStyle34.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColYear3_Bud.DefaultCellStyle = DataGridViewCellStyle34
        Me.ColYear3_Bud.HeaderText = "Year3_Budget"
        Me.ColYear3_Bud.Name = "ColYear3_Bud"
        Me.ColYear3_Bud.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColYear3_Bud.Width = 81
        '
        'ColYear3_Rea
        '
        DataGridViewCellStyle35.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColYear3_Rea.DefaultCellStyle = DataGridViewCellStyle35
        Me.ColYear3_Rea.HeaderText = "Year3 - Rea"
        Me.ColYear3_Rea.Name = "ColYear3_Rea"
        Me.ColYear3_Rea.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColYear3_Rea.Width = 70
        '
        'ColYear2_Budget
        '
        DataGridViewCellStyle36.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColYear2_Budget.DefaultCellStyle = DataGridViewCellStyle36
        Me.ColYear2_Budget.HeaderText = "Year2-Budget"
        Me.ColYear2_Budget.Name = "ColYear2_Budget"
        Me.ColYear2_Budget.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColYear2_Budget.Width = 78
        '
        'ColYear2_Rea
        '
        DataGridViewCellStyle37.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColYear2_Rea.DefaultCellStyle = DataGridViewCellStyle37
        Me.ColYear2_Rea.HeaderText = "Year 2 - Rea"
        Me.ColYear2_Rea.Name = "ColYear2_Rea"
        Me.ColYear2_Rea.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColYear2_Rea.Width = 73
        '
        'ColYear1_Bud
        '
        DataGridViewCellStyle38.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColYear1_Bud.DefaultCellStyle = DataGridViewCellStyle38
        Me.ColYear1_Bud.HeaderText = "Year1-Budget"
        Me.ColYear1_Bud.Name = "ColYear1_Bud"
        Me.ColYear1_Bud.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColYear1_Bud.Width = 78
        '
        'ColYear1_Rea
        '
        DataGridViewCellStyle39.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColYear1_Rea.DefaultCellStyle = DataGridViewCellStyle39
        Me.ColYear1_Rea.HeaderText = "Year-1 Rea"
        Me.ColYear1_Rea.Name = "ColYear1_Rea"
        Me.ColYear1_Rea.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColYear1_Rea.Width = 67
        '
        'ColCurYear_Budget
        '
        DataGridViewCellStyle40.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCurYear_Budget.DefaultCellStyle = DataGridViewCellStyle40
        Me.ColCurYear_Budget.HeaderText = "CurrentYear - Budget"
        Me.ColCurYear_Budget.Name = "ColCurYear_Budget"
        Me.ColCurYear_Budget.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCurYear_Budget.Width = 112
        '
        'ColCurrenctYear_Rea
        '
        DataGridViewCellStyle41.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCurrenctYear_Rea.DefaultCellStyle = DataGridViewCellStyle41
        Me.ColCurrenctYear_Rea.HeaderText = "Current Year - Rea"
        Me.ColCurrenctYear_Rea.Name = "ColCurrenctYear_Rea"
        Me.ColCurrenctYear_Rea.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCurrenctYear_Rea.Width = 101
        '
        'ColNextYear_Bud
        '
        DataGridViewCellStyle42.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNextYear_Bud.DefaultCellStyle = DataGridViewCellStyle42
        Me.ColNextYear_Bud.HeaderText = "NextYearBudget"
        Me.ColNextYear_Bud.Name = "ColNextYear_Bud"
        Me.ColNextYear_Bud.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNextYear_Bud.Width = 91
        '
        'ColPrimaryAkunYearly
        '
        Me.ColPrimaryAkunYearly.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkunYearly.Name = "ColPrimaryAkunYearly"
        Me.ColPrimaryAkunYearly.Visible = False
        Me.ColPrimaryAkunYearly.Width = 91
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(194, 150)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 522)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(931, 22)
        Me.StatusStrip1.TabIndex = 32
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(229, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(229, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(229, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(229, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'FrmMasterBudget
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(931, 544)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(947, 580)
        Me.Name = "FrmMasterBudget"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Budget"
        CType(Me.dgvBudget, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvYearly, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvBudget As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkYear1_Bud As System.Windows.Forms.CheckBox
    Friend WithEvents chkYear1_Rea As System.Windows.Forms.CheckBox
    Friend WithEvents chkYear2_Rea As System.Windows.Forms.CheckBox
    Friend WithEvents chkYear2_Bud As System.Windows.Forms.CheckBox
    Friend WithEvents chkYear3_Rea As System.Windows.Forms.CheckBox
    Friend WithEvents chkYear3_Bud As System.Windows.Forms.CheckBox
    Friend WithEvents chkNextYear As System.Windows.Forms.CheckBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvYearly As System.Windows.Forms.DataGridView
    Friend WithEvents cboFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnShowDetail As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFeb As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMar As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColApr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJune As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJul As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAgt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNov As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents ColAccount_Yearly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYear3_Bud As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYear3_Rea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYear2_Budget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYear2_Rea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYear1_Bud As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYear1_Rea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurYear_Budget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrenctYear_Rea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNextYear_Bud As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkunYearly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
