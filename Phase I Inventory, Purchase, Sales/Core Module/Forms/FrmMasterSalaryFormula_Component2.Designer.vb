<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalaryFormula_Component2
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMasterSalaryFormula_Component2))
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DGVConstaint = New System.Windows.Forms.DataGridView
        Me.ColConts_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColConst_ParameterCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColConst_ParameterDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColConst_Type = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColConst_Compare = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColConst_RefValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColConst_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColConst_PrimaryParameter = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryRefValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtFormula = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DGVVariable = New System.Windows.Forms.DataGridView
        Me.Label28 = New System.Windows.Forms.Label
        Me.CboRoundingType = New System.Windows.Forms.ComboBox
        Me.TxtRoundingUnit = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BtnClose = New System.Windows.Forms.Button
        Me.TxtYearlyMaxValue = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtYearlyMinValue = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.DtpYear = New System.Windows.Forms.DateTimePicker
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtMonthlyMinValue = New System.Windows.Forms.TextBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TxtMonthlyMaxValue = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.CboRoundingUnit = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.ColVariable_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVariable_Parameter_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVariable_Parameter_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVariable_Type = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColVariable_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVariable_Variable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRoundingUnit_Variable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRoundingTipe_Variable = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColVariable_PrimaryParameter = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        CType(Me.DGVConstaint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGVVariable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DGVConstaint)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(797, 211)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Constraint"
        '
        'DGVConstaint
        '
        Me.DGVConstaint.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVConstaint.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVConstaint.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DGVConstaint.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVConstaint.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColConts_No, Me.ColConst_ParameterCode, Me.ColConst_ParameterDesc, Me.ColConst_Type, Me.ColConst_Compare, Me.ColConst_RefValue, Me.ColConst_Unit, Me.ColConst_PrimaryParameter, Me.ColPrimaryRefValue})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGVConstaint.DefaultCellStyle = DataGridViewCellStyle2
        Me.DGVConstaint.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVConstaint.Location = New System.Drawing.Point(3, 17)
        Me.DGVConstaint.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DGVConstaint.Name = "DGVConstaint"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVConstaint.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DGVConstaint.RowTemplate.Height = 24
        Me.DGVConstaint.Size = New System.Drawing.Size(791, 192)
        Me.DGVConstaint.TabIndex = 0
        '
        'ColConts_No
        '
        Me.ColConts_No.HeaderText = "No"
        Me.ColConts_No.Name = "ColConts_No"
        Me.ColConts_No.Width = 51
        '
        'ColConst_ParameterCode
        '
        Me.ColConst_ParameterCode.HeaderText = "Parameter"
        Me.ColConst_ParameterCode.Name = "ColConst_ParameterCode"
        Me.ColConst_ParameterCode.Width = 99
        '
        'ColConst_ParameterDesc
        '
        Me.ColConst_ParameterDesc.HeaderText = "Description"
        Me.ColConst_ParameterDesc.Name = "ColConst_ParameterDesc"
        Me.ColConst_ParameterDesc.Width = 104
        '
        'ColConst_Type
        '
        Me.ColConst_Type.HeaderText = "Type"
        Me.ColConst_Type.Items.AddRange(New Object() {"Daily", "Monthly", "Yearly", "Person Base"})
        Me.ColConst_Type.Name = "ColConst_Type"
        Me.ColConst_Type.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColConst_Type.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColConst_Type.Visible = False
        Me.ColConst_Type.Width = 65
        '
        'ColConst_Compare
        '
        Me.ColConst_Compare.HeaderText = "Compare"
        Me.ColConst_Compare.Items.AddRange(New Object() {"=", ">", "<", ">=", "<=", "<>"})
        Me.ColConst_Compare.Name = "ColConst_Compare"
        Me.ColConst_Compare.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColConst_Compare.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColConst_Compare.Width = 90
        '
        'ColConst_RefValue
        '
        Me.ColConst_RefValue.HeaderText = "Ref Value"
        Me.ColConst_RefValue.Name = "ColConst_RefValue"
        Me.ColConst_RefValue.Width = 95
        '
        'ColConst_Unit
        '
        Me.ColConst_Unit.HeaderText = "Unit"
        Me.ColConst_Unit.Name = "ColConst_Unit"
        Me.ColConst_Unit.Width = 58
        '
        'ColConst_PrimaryParameter
        '
        Me.ColConst_PrimaryParameter.HeaderText = "Primary Parameter"
        Me.ColConst_PrimaryParameter.Name = "ColConst_PrimaryParameter"
        Me.ColConst_PrimaryParameter.Visible = False
        Me.ColConst_PrimaryParameter.Width = 151
        '
        'ColPrimaryRefValue
        '
        Me.ColPrimaryRefValue.HeaderText = "Primary Ref Value"
        Me.ColPrimaryRefValue.Name = "ColPrimaryRefValue"
        Me.ColPrimaryRefValue.Visible = False
        Me.ColPrimaryRefValue.Width = 147
        '
        'TxtFormula
        '
        Me.TxtFormula.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFormula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFormula.Location = New System.Drawing.Point(183, 7)
        Me.TxtFormula.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtFormula.MaxLength = 250
        Me.TxtFormula.Name = "TxtFormula"
        Me.TxtFormula.Size = New System.Drawing.Size(628, 22)
        Me.TxtFormula.TabIndex = 0
        Me.TxtFormula.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 17)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Formula"
        '
        'DGVVariable
        '
        Me.DGVVariable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVVariable.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVVariable.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DGVVariable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVVariable.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColVariable_No, Me.ColVariable_Parameter_Code, Me.ColVariable_Parameter_Desc, Me.ColVariable_Type, Me.ColVariable_Unit, Me.ColVariable_Variable, Me.ColRoundingUnit_Variable, Me.ColRoundingTipe_Variable, Me.ColVariable_PrimaryParameter})
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGVVariable.DefaultCellStyle = DataGridViewCellStyle6
        Me.DGVVariable.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVVariable.Location = New System.Drawing.Point(3, 17)
        Me.DGVVariable.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DGVVariable.Name = "DGVVariable"
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVVariable.RowHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.DGVVariable.RowTemplate.Height = 24
        Me.DGVVariable.Size = New System.Drawing.Size(791, 194)
        Me.DGVVariable.TabIndex = 0
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(8, 7)
        Me.Label28.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(105, 17)
        Me.Label28.TabIndex = 17
        Me.Label28.Text = "Rounding Type"
        '
        'CboRoundingType
        '
        Me.CboRoundingType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboRoundingType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboRoundingType.FormattingEnabled = True
        Me.CboRoundingType.Items.AddRange(New Object() {"No Rounding", "Round Down", "Round Middlepoint", "Round Up"})
        Me.CboRoundingType.Location = New System.Drawing.Point(165, 4)
        Me.CboRoundingType.Margin = New System.Windows.Forms.Padding(4)
        Me.CboRoundingType.Name = "CboRoundingType"
        Me.CboRoundingType.Size = New System.Drawing.Size(227, 24)
        Me.CboRoundingType.TabIndex = 1
        '
        'TxtRoundingUnit
        '
        Me.TxtRoundingUnit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRoundingUnit.Location = New System.Drawing.Point(161, 5)
        Me.TxtRoundingUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRoundingUnit.MaxLength = 250
        Me.TxtRoundingUnit.Name = "TxtRoundingUnit"
        Me.TxtRoundingUnit.Size = New System.Drawing.Size(229, 22)
        Me.TxtRoundingUnit.TabIndex = 2
        Me.TxtRoundingUnit.TabStop = False
        Me.TxtRoundingUnit.Text = "1"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(8, 9)
        Me.Label29.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(98, 17)
        Me.Label29.TabIndex = 14
        Me.Label29.Text = "Rounding Unit"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.DGVVariable)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 2)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(797, 213)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Variable"
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(711, 604)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 28)
        Me.BtnClose.TabIndex = 6
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'TxtYearlyMaxValue
        '
        Me.TxtYearlyMaxValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtYearlyMaxValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtYearlyMaxValue.Location = New System.Drawing.Point(157, 4)
        Me.TxtYearlyMaxValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtYearlyMaxValue.MaxLength = 250
        Me.TxtYearlyMaxValue.Name = "TxtYearlyMaxValue"
        Me.TxtYearlyMaxValue.Size = New System.Drawing.Size(120, 22)
        Me.TxtYearlyMaxValue.TabIndex = 4
        Me.TxtYearlyMaxValue.Text = "0"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 7)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(117, 17)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Yearly Max Value"
        '
        'TxtYearlyMinValue
        '
        Me.TxtYearlyMinValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtYearlyMinValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtYearlyMinValue.Location = New System.Drawing.Point(167, 4)
        Me.TxtYearlyMinValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtYearlyMinValue.MaxLength = 250
        Me.TxtYearlyMinValue.Name = "TxtYearlyMinValue"
        Me.TxtYearlyMinValue.Size = New System.Drawing.Size(227, 22)
        Me.TxtYearlyMinValue.TabIndex = 3
        Me.TxtYearlyMinValue.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 7)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(114, 17)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Yearly Min Value"
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(12, 604)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 8
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        Me.btnUP.Visible = False
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(57, 604)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 9
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        Me.btnDown.Visible = False
        '
        'DtpYear
        '
        Me.DtpYear.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpYear.CustomFormat = "MMMM"
        Me.DtpYear.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpYear.Location = New System.Drawing.Point(4, 4)
        Me.DtpYear.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpYear.Name = "DtpYear"
        Me.DtpYear.ShowUpDown = True
        Me.DtpYear.Size = New System.Drawing.Size(97, 22)
        Me.DtpYear.TabIndex = 5
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(16, 79)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtMonthlyMinValue)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtYearlyMinValue)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtMonthlyMaxValue)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(803, 69)
        Me.SplitContainer1.SplitterDistance = 397
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 2
        Me.SplitContainer1.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 39)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(123, 17)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Monthly Min Value"
        '
        'TxtMonthlyMinValue
        '
        Me.TxtMonthlyMinValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMonthlyMinValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMonthlyMinValue.Location = New System.Drawing.Point(167, 36)
        Me.TxtMonthlyMinValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMonthlyMinValue.MaxLength = 250
        Me.TxtMonthlyMinValue.Name = "TxtMonthlyMinValue"
        Me.TxtMonthlyMinValue.Size = New System.Drawing.Size(227, 22)
        Me.TxtMonthlyMinValue.TabIndex = 6
        Me.TxtMonthlyMinValue.Text = "0"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(5, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtYearlyMaxValue)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.DtpYear)
        Me.SplitContainer3.Size = New System.Drawing.Size(393, 31)
        Me.SplitContainer3.SplitterDistance = 282
        Me.SplitContainer3.SplitterWidth = 5
        Me.SplitContainer3.TabIndex = 0
        Me.SplitContainer3.TabStop = False
        '
        'TxtMonthlyMaxValue
        '
        Me.TxtMonthlyMaxValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMonthlyMaxValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMonthlyMaxValue.Location = New System.Drawing.Point(163, 36)
        Me.TxtMonthlyMaxValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMonthlyMaxValue.MaxLength = 250
        Me.TxtMonthlyMaxValue.Name = "TxtMonthlyMaxValue"
        Me.TxtMonthlyMaxValue.Size = New System.Drawing.Size(230, 22)
        Me.TxtMonthlyMaxValue.TabIndex = 7
        Me.TxtMonthlyMaxValue.Text = "0"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 39)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(126, 17)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Monthly Max Value"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(16, 38)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.CboRoundingType)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label28)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.CboRoundingUnit)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRoundingUnit)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label29)
        Me.SplitContainer2.Size = New System.Drawing.Size(803, 33)
        Me.SplitContainer2.SplitterDistance = 398
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 1
        Me.SplitContainer2.TabStop = False
        '
        'CboRoundingUnit
        '
        Me.CboRoundingUnit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboRoundingUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboRoundingUnit.FormattingEnabled = True
        Me.CboRoundingUnit.Items.AddRange(New Object() {"0.000001", "0.00001", "0.0001", "0.001", "0.01", "0.1", "1", "10", "100", "1000", "10000", "100000", "1000000"})
        Me.CboRoundingUnit.Location = New System.Drawing.Point(161, 4)
        Me.CboRoundingUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.CboRoundingUnit.Name = "CboRoundingUnit"
        Me.CboRoundingUnit.Size = New System.Drawing.Size(229, 24)
        Me.CboRoundingUnit.TabIndex = 2
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(607, 604)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(100, 28)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
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
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(16, 155)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer4.Size = New System.Drawing.Size(803, 439)
        Me.SplitContainer4.SplitterDistance = 218
        Me.SplitContainer4.SplitterWidth = 5
        Me.SplitContainer4.TabIndex = 12
        '
        'ColVariable_No
        '
        Me.ColVariable_No.HeaderText = "No"
        Me.ColVariable_No.Name = "ColVariable_No"
        Me.ColVariable_No.Width = 51
        '
        'ColVariable_Parameter_Code
        '
        Me.ColVariable_Parameter_Code.HeaderText = "Parameter"
        Me.ColVariable_Parameter_Code.Name = "ColVariable_Parameter_Code"
        Me.ColVariable_Parameter_Code.Width = 99
        '
        'ColVariable_Parameter_Desc
        '
        Me.ColVariable_Parameter_Desc.HeaderText = "Description"
        Me.ColVariable_Parameter_Desc.Name = "ColVariable_Parameter_Desc"
        Me.ColVariable_Parameter_Desc.Width = 104
        '
        'ColVariable_Type
        '
        Me.ColVariable_Type.HeaderText = "Type"
        Me.ColVariable_Type.Items.AddRange(New Object() {"Daily", "Monthly", "Yearly", "Person Base"})
        Me.ColVariable_Type.Name = "ColVariable_Type"
        Me.ColVariable_Type.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColVariable_Type.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColVariable_Type.Visible = False
        Me.ColVariable_Type.Width = 65
        '
        'ColVariable_Unit
        '
        Me.ColVariable_Unit.HeaderText = "Unit"
        Me.ColVariable_Unit.Name = "ColVariable_Unit"
        Me.ColVariable_Unit.Width = 58
        '
        'ColVariable_Variable
        '
        Me.ColVariable_Variable.HeaderText = "Variabel"
        Me.ColVariable_Variable.MaxInputLength = 5
        Me.ColVariable_Variable.Name = "ColVariable_Variable"
        Me.ColVariable_Variable.Width = 85
        '
        'ColRoundingUnit_Variable
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRoundingUnit_Variable.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRoundingUnit_Variable.HeaderText = "Rounding Unit"
        Me.ColRoundingUnit_Variable.Name = "ColRoundingUnit_Variable"
        Me.ColRoundingUnit_Variable.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRoundingUnit_Variable.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRoundingUnit_Variable.Width = 104
        '
        'ColRoundingTipe_Variable
        '
        Me.ColRoundingTipe_Variable.HeaderText = "Rounding Tipe"
        Me.ColRoundingTipe_Variable.Items.AddRange(New Object() {"No Rounding", "Round Down", "Round Middlepoint", "Round Up"})
        Me.ColRoundingTipe_Variable.Name = "ColRoundingTipe_Variable"
        Me.ColRoundingTipe_Variable.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRoundingTipe_Variable.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRoundingTipe_Variable.Width = 126
        '
        'ColVariable_PrimaryParameter
        '
        Me.ColVariable_PrimaryParameter.HeaderText = "Primary Parameter"
        Me.ColVariable_PrimaryParameter.Name = "ColVariable_PrimaryParameter"
        Me.ColVariable_PrimaryParameter.Visible = False
        Me.ColVariable_PrimaryParameter.Width = 151
        '
        'FrmMasterSalaryFormula_Component2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(837, 651)
        Me.Controls.Add(Me.SplitContainer4)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.TxtFormula)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnUP)
        Me.Controls.Add(Me.btnDown)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(845, 688)
        Me.Name = "FrmMasterSalaryFormula_Component2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Salary Component Formula"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DGVConstaint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGVVariable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents TxtFormula As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DGVVariable As System.Windows.Forms.DataGridView
    Friend WithEvents DGVConstaint As System.Windows.Forms.DataGridView
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents CboRoundingType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtRoundingUnit As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtYearlyMaxValue As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtYearlyMinValue As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents DtpYear As System.Windows.Forms.DateTimePicker
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtMonthlyMaxValue As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtMonthlyMinValue As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents CboRoundingUnit As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents ColConts_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColConst_ParameterCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColConst_ParameterDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColConst_Type As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColConst_Compare As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColConst_RefValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColConst_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColConst_PrimaryParameter As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryRefValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVariable_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVariable_Parameter_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVariable_Parameter_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVariable_Type As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColVariable_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVariable_Variable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRoundingUnit_Variable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRoundingTipe_Variable As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColVariable_PrimaryParameter As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
