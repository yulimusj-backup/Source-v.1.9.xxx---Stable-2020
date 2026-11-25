<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCommission
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
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.cboProgresiveCalculation = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboLimitType = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.cboBaseCalculate = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_ConditionMin = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_ConditionMax = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCommision = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tspBar1 = New System.Windows.Forms.ToolStripProgressBar
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.chkSalesPeriod = New System.Windows.Forms.CheckBox
        Me.btnSet = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvFilterComm = New System.Windows.Forms.DataGridView
        Me.ColFilter_CommisionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_CommissionDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_PrimaryCommission = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkCommission = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkProduct = New System.Windows.Forms.CheckBox
        Me.dgvAllocate = New System.Windows.Forms.DataGridView
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.bgwCom = New System.ComponentModel.BackgroundWorker
        Me.bgwAllocate = New System.ComponentModel.BackgroundWorker
        Me.bgwProduct = New System.ComponentModel.BackgroundWorker
        Me.ColAllocate_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_CommisionnCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_CommissionDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_PrimaryCommision = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_Set = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvFilterComm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAllocate, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(792, 539)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.gboDetail)
        Me.TabPage1.Controls.Add(Me.StatusStrip2)
        Me.TabPage1.Controls.Add(Me.StatusStrip1)
        Me.TabPage1.Controls.Add(Me.btnEdit)
        Me.TabPage1.Controls.Add(Me.btnNew)
        Me.TabPage1.Controls.Add(Me.btnFind)
        Me.TabPage1.Controls.Add(Me.btnExit)
        Me.TabPage1.Controls.Add(Me.btnDelete)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Size = New System.Drawing.Size(784, 510)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Commission Type"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.cboProgresiveCalculation)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.cboLimitType)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.txtReadOnly)
        Me.gboDetail.Controls.Add(Me.cboBaseCalculate)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtRemarks)
        Me.gboDetail.Controls.Add(Me.txtPrimary)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.txtDescription)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.txtCode)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Location = New System.Drawing.Point(16, 9)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(761, 394)
        Me.gboDetail.TabIndex = 22
        Me.gboDetail.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(248, 354)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(117, 32)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'cboProgresiveCalculation
        '
        Me.cboProgresiveCalculation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboProgresiveCalculation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProgresiveCalculation.FormattingEnabled = True
        Me.cboProgresiveCalculation.Items.AddRange(New Object() {"No Progressive", "Progressive"})
        Me.cboProgresiveCalculation.Location = New System.Drawing.Point(189, 148)
        Me.cboProgresiveCalculation.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboProgresiveCalculation.Name = "cboProgresiveCalculation"
        Me.cboProgresiveCalculation.Size = New System.Drawing.Size(564, 24)
        Me.cboProgresiveCalculation.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 153)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(156, 17)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "Progressive Calculation"
        '
        'cboLimitType
        '
        Me.cboLimitType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLimitType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLimitType.FormattingEnabled = True
        Me.cboLimitType.Items.AddRange(New Object() {"Qty", "Amount"})
        Me.cboLimitType.Location = New System.Drawing.Point(189, 117)
        Me.cboLimitType.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboLimitType.Name = "cboLimitType"
        Me.cboLimitType.Size = New System.Drawing.Size(564, 24)
        Me.cboLimitType.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 122)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 49
        Me.Label3.Text = "Limit Type"
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReadOnly.Location = New System.Drawing.Point(7, 359)
        Me.txtReadOnly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(55, 22)
        Me.txtReadOnly.TabIndex = 48
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'cboBaseCalculate
        '
        Me.cboBaseCalculate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBaseCalculate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBaseCalculate.FormattingEnabled = True
        Me.cboBaseCalculate.Items.AddRange(New Object() {"Fixed Amount For Each Sold Item", "% of Sales Amount"})
        Me.cboBaseCalculate.Location = New System.Drawing.Point(189, 86)
        Me.cboBaseCalculate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboBaseCalculate.Name = "cboBaseCalculate"
        Me.cboBaseCalculate.Size = New System.Drawing.Size(564, 24)
        Me.cboBaseCalculate.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 91)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 17)
        Me.Label4.TabIndex = 46
        Me.Label4.Text = "Base Calculation"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.Col_ConditionMin, Me.ColMin, Me.Col_ConditionMax, Me.ColMax, Me.ColCommision})
        Me.dgvItem.Location = New System.Drawing.Point(7, 210)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(747, 137)
        Me.dgvItem.TabIndex = 6
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'Col_ConditionMin
        '
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Col_ConditionMin.DefaultCellStyle = DataGridViewCellStyle16
        Me.Col_ConditionMin.HeaderText = ""
        Me.Col_ConditionMin.Items.AddRange(New Object() {">", ">="})
        Me.Col_ConditionMin.MinimumWidth = 30
        Me.Col_ConditionMin.Name = "Col_ConditionMin"
        Me.Col_ConditionMin.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col_ConditionMin.Width = 30
        '
        'ColMin
        '
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMin.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColMin.HeaderText = "Min"
        Me.ColMin.MaxInputLength = 15
        Me.ColMin.Name = "ColMin"
        Me.ColMin.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMin.Width = 36
        '
        'Col_ConditionMax
        '
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Col_ConditionMax.DefaultCellStyle = DataGridViewCellStyle18
        Me.Col_ConditionMax.HeaderText = ""
        Me.Col_ConditionMax.Items.AddRange(New Object() {"<", "<="})
        Me.Col_ConditionMax.MinimumWidth = 30
        Me.Col_ConditionMax.Name = "Col_ConditionMax"
        Me.Col_ConditionMax.Width = 30
        '
        'ColMax
        '
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMax.DefaultCellStyle = DataGridViewCellStyle19
        Me.ColMax.HeaderText = "Max"
        Me.ColMax.MaxInputLength = 15
        Me.ColMax.Name = "ColMax"
        Me.ColMax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMax.Width = 39
        '
        'ColCommision
        '
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCommision.DefaultCellStyle = DataGridViewCellStyle20
        Me.ColCommision.HeaderText = "Comm. Value"
        Me.ColCommision.MaxInputLength = 15
        Me.ColCommision.Name = "ColCommision"
        Me.ColCommision.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCommision.Width = 97
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 185)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "Remarks"
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(189, 180)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtRemarks.MaxLength = 75
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(564, 22)
        Me.txtRemarks.TabIndex = 5
        '
        'txtPrimary
        '
        Me.txtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimary.Location = New System.Drawing.Point(125, 23)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(55, 22)
        Me.txtPrimary.TabIndex = 36
        Me.txtPrimary.TabStop = False
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 60)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(79, 17)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Description"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(189, 55)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDescription.MaxLength = 75
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(564, 22)
        Me.txtDescription.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 28)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Code"
        '
        'txtCode
        '
        Me.txtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCode.Location = New System.Drawing.Point(189, 23)
        Me.txtCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtCode.MaxLength = 45
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(564, 22)
        Me.txtCode.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(389, 354)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(117, 32)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(3, 458)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(778, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 24
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(379, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(379, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate, Me.tspBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(3, 483)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(778, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 23
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(379, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(379, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'tspBar1
        '
        Me.tspBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tspBar1.Name = "tspBar1"
        Me.tspBar1.Size = New System.Drawing.Size(133, 23)
        Me.tspBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.tspBar1.Visible = False
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(333, 411)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(117, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(205, 411)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(117, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(77, 411)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(117, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(589, 411)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(117, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(461, 411)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(117, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dtpTo)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.dtpFrom)
        Me.TabPage2.Controls.Add(Me.chkSalesPeriod)
        Me.TabPage2.Controls.Add(Me.btnSet)
        Me.TabPage2.Controls.Add(Me.btnFilter)
        Me.TabPage2.Controls.Add(Me.llbUncheckAll)
        Me.TabPage2.Controls.Add(Me.llbCheckAll)
        Me.TabPage2.Controls.Add(Me.SplitContainer1)
        Me.TabPage2.Controls.Add(Me.dgvAllocate)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Size = New System.Drawing.Size(784, 510)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Commision Allocation"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(334, 7)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(151, 22)
        Me.dtpTo.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(306, 10)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(20, 17)
        Me.Label6.TabIndex = 54
        Me.Label6.Text = "to"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(152, 7)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(151, 22)
        Me.dtpFrom.TabIndex = 0
        '
        'chkSalesPeriod
        '
        Me.chkSalesPeriod.AutoSize = True
        Me.chkSalesPeriod.Location = New System.Drawing.Point(11, 7)
        Me.chkSalesPeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkSalesPeriod.Name = "chkSalesPeriod"
        Me.chkSalesPeriod.Size = New System.Drawing.Size(135, 21)
        Me.chkSalesPeriod.TabIndex = 52
        Me.chkSalesPeriod.Text = "Purchase Period"
        Me.chkSalesPeriod.UseVisualStyleBackColor = True
        '
        'btnSet
        '
        Me.btnSet.Location = New System.Drawing.Point(591, 7)
        Me.btnSet.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnSet.Name = "btnSet"
        Me.btnSet.Size = New System.Drawing.Size(80, 26)
        Me.btnSet.TabIndex = 3
        Me.btnSet.Text = "Set"
        Me.btnSet.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(502, 7)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(80, 26)
        Me.btnFilter.TabIndex = 2
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(77, 490)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncheckAll.TabIndex = 6
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(5, 490)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCheckAll.TabIndex = 5
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(5, 36)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(772, 186)
        Me.SplitContainer1.SplitterDistance = 382
        Me.SplitContainer1.TabIndex = 47
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvFilterComm)
        Me.GroupBox1.Controls.Add(Me.chkCommission)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(382, 186)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'dgvFilterComm
        '
        Me.dgvFilterComm.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFilterComm.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilter_CommisionCode, Me.ColFilter_CommissionDesc, Me.ColFilter_PrimaryCommission})
        Me.dgvFilterComm.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFilterComm.Location = New System.Drawing.Point(3, 17)
        Me.dgvFilterComm.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvFilterComm.MultiSelect = False
        Me.dgvFilterComm.Name = "dgvFilterComm"
        Me.dgvFilterComm.ReadOnly = True
        Me.dgvFilterComm.RowTemplate.Height = 24
        Me.dgvFilterComm.Size = New System.Drawing.Size(376, 167)
        Me.dgvFilterComm.TabIndex = 0
        '
        'ColFilter_CommisionCode
        '
        Me.ColFilter_CommisionCode.HeaderText = "Code"
        Me.ColFilter_CommisionCode.Name = "ColFilter_CommisionCode"
        Me.ColFilter_CommisionCode.ReadOnly = True
        Me.ColFilter_CommisionCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColFilter_CommissionDesc
        '
        Me.ColFilter_CommissionDesc.HeaderText = "Description"
        Me.ColFilter_CommissionDesc.Name = "ColFilter_CommissionDesc"
        Me.ColFilter_CommissionDesc.ReadOnly = True
        Me.ColFilter_CommissionDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColFilter_PrimaryCommission
        '
        Me.ColFilter_PrimaryCommission.HeaderText = "Primary"
        Me.ColFilter_PrimaryCommission.Name = "ColFilter_PrimaryCommission"
        Me.ColFilter_PrimaryCommission.ReadOnly = True
        Me.ColFilter_PrimaryCommission.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFilter_PrimaryCommission.Visible = False
        '
        'chkCommission
        '
        Me.chkCommission.AutoSize = True
        Me.chkCommission.BackColor = System.Drawing.Color.White
        Me.chkCommission.Location = New System.Drawing.Point(5, -1)
        Me.chkCommission.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkCommission.Name = "chkCommission"
        Me.chkCommission.Size = New System.Drawing.Size(105, 21)
        Me.chkCommission.TabIndex = 0
        Me.chkCommission.Text = "Commission"
        Me.chkCommission.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvItemProductGroup)
        Me.GroupBox2.Controls.Add(Me.chkProduct)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(386, 186)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(3, 17)
        Me.dgvItemProductGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.ReadOnly = True
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(380, 167)
        Me.dgvItemProductGroup.TabIndex = 0
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.ReadOnly = True
        Me.ColProductGroup_No.Visible = False
        Me.ColProductGroup_No.Width = 51
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.ReadOnly = True
        Me.ColProductGroup_Level.Width = 48
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.ReadOnly = True
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 47
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
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
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 106
        '
        'chkProduct
        '
        Me.chkProduct.AutoSize = True
        Me.chkProduct.BackColor = System.Drawing.Color.White
        Me.chkProduct.Location = New System.Drawing.Point(5, -1)
        Me.chkProduct.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkProduct.Name = "chkProduct"
        Me.chkProduct.Size = New System.Drawing.Size(79, 21)
        Me.chkProduct.TabIndex = 0
        Me.chkProduct.Text = "Product"
        Me.chkProduct.UseVisualStyleBackColor = False
        '
        'dgvAllocate
        '
        Me.dgvAllocate.AllowUserToAddRows = False
        Me.dgvAllocate.AllowUserToDeleteRows = False
        Me.dgvAllocate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAllocate.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAllocate.ColumnHeadersHeight = 25
        Me.dgvAllocate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAllocate.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAllocate_No, Me.ColAllocate_CodeNo, Me.ColAllocate_Desc, Me.ColAllocate_CommisionnCode, Me.ColAllocate_CommissionDesc, Me.ColAllocate_PrimaryCodeNo, Me.ColAllocate_PrimaryCommision, Me.ColAllocate_Set})
        Me.dgvAllocate.Location = New System.Drawing.Point(8, 226)
        Me.dgvAllocate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvAllocate.MultiSelect = False
        Me.dgvAllocate.Name = "dgvAllocate"
        Me.dgvAllocate.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAllocate.RowTemplate.Height = 24
        Me.dgvAllocate.Size = New System.Drawing.Size(767, 261)
        Me.dgvAllocate.TabIndex = 4
        '
        'BackgroundWorker1
        '
        '
        'ColAllocate_No
        '
        Me.ColAllocate_No.HeaderText = "No"
        Me.ColAllocate_No.Name = "ColAllocate_No"
        Me.ColAllocate_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAllocate_No.Width = 32
        '
        'ColAllocate_CodeNo
        '
        Me.ColAllocate_CodeNo.HeaderText = "Code No"
        Me.ColAllocate_CodeNo.Name = "ColAllocate_CodeNo"
        Me.ColAllocate_CodeNo.Width = 88
        '
        'ColAllocate_Desc
        '
        Me.ColAllocate_Desc.HeaderText = "Description"
        Me.ColAllocate_Desc.Name = "ColAllocate_Desc"
        Me.ColAllocate_Desc.Width = 104
        '
        'ColAllocate_CommisionnCode
        '
        Me.ColAllocate_CommisionnCode.HeaderText = "Commission Code"
        Me.ColAllocate_CommisionnCode.Name = "ColAllocate_CommisionnCode"
        Me.ColAllocate_CommisionnCode.Width = 145
        '
        'ColAllocate_CommissionDesc
        '
        Me.ColAllocate_CommissionDesc.HeaderText = "Commission Description"
        Me.ColAllocate_CommissionDesc.Name = "ColAllocate_CommissionDesc"
        Me.ColAllocate_CommissionDesc.Width = 183
        '
        'ColAllocate_PrimaryCodeNo
        '
        Me.ColAllocate_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColAllocate_PrimaryCodeNo.Name = "ColAllocate_PrimaryCodeNo"
        Me.ColAllocate_PrimaryCodeNo.Visible = False
        Me.ColAllocate_PrimaryCodeNo.Width = 140
        '
        'ColAllocate_PrimaryCommision
        '
        Me.ColAllocate_PrimaryCommision.HeaderText = "Primary Commission"
        Me.ColAllocate_PrimaryCommision.Name = "ColAllocate_PrimaryCommision"
        Me.ColAllocate_PrimaryCommision.Visible = False
        Me.ColAllocate_PrimaryCommision.Width = 160
        '
        'ColAllocate_Set
        '
        Me.ColAllocate_Set.HeaderText = "Set"
        Me.ColAllocate_Set.Name = "ColAllocate_Set"
        Me.ColAllocate_Set.Width = 35
        '
        'FrmMasterCommission
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(811, 564)
        Me.Controls.Add(Me.TabControl1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(803, 601)
        Me.Name = "FrmMasterCommission"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Agent Commision Setting"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvFilterComm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAllocate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tspBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cboBaseCalculate As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvAllocate As System.Windows.Forms.DataGridView
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvFilterComm As System.Windows.Forms.DataGridView
    Friend WithEvents chkCommission As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkProduct As System.Windows.Forms.CheckBox
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents btnSet As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkSalesPeriod As System.Windows.Forms.CheckBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCom As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwAllocate As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwProduct As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColFilter_CommisionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_CommissionDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_PrimaryCommission As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cboProgresiveCalculation As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboLimitType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_ConditionMin As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_ConditionMax As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCommision As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColAllocate_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_CommisionnCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_CommissionDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_PrimaryCommision As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_Set As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
