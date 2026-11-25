<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Export_eSPT_Check
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
        Me.btnStart = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboCheckType = New System.Windows.Forms.ComboBox
        Me.TxtCheckDesc = New System.Windows.Forms.TextBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGrossIncomeMonthly = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColGrossIncome_PrevIncomeTax = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColGrossIncomeYearly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncomeTaxMonthly = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColIncomeTaxYearly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColA1FormNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.ChkShowWrongDataOnly = New System.Windows.Forms.CheckBox
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(353, 13)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 1
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Check Type"
        '
        'cboCheckType
        '
        Me.cboCheckType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCheckType.FormattingEnabled = True
        Me.cboCheckType.Items.AddRange(New Object() {"Gross Income Data Check", "Prev. Income Data Check"})
        Me.cboCheckType.Location = New System.Drawing.Point(101, 12)
        Me.cboCheckType.Name = "cboCheckType"
        Me.cboCheckType.Size = New System.Drawing.Size(246, 24)
        Me.cboCheckType.TabIndex = 0
        '
        'TxtCheckDesc
        '
        Me.TxtCheckDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCheckDesc.Location = New System.Drawing.Point(101, 42)
        Me.TxtCheckDesc.Multiline = True
        Me.TxtCheckDesc.Name = "TxtCheckDesc"
        Me.TxtCheckDesc.ReadOnly = True
        Me.TxtCheckDesc.Size = New System.Drawing.Size(600, 75)
        Me.TxtCheckDesc.TabIndex = 4
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
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColGrossIncomeMonthly, Me.ColGrossIncome_PrevIncomeTax, Me.ColGrossIncomeYearly, Me.ColIncomeTaxMonthly, Me.ColIncomeTaxYearly, Me.ColA1FormNo, Me.ColEmployer, Me.ColPrimaryEmployee})
        Me.dgvItem.Location = New System.Drawing.Point(15, 123)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(686, 474)
        Me.dgvItem.TabIndex = 7
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.ReadOnly = True
        Me.ColEmployeeCode.Width = 66
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.Width = 70
        '
        'ColGrossIncomeMonthly
        '
        Me.ColGrossIncomeMonthly.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColGrossIncomeMonthly.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColGrossIncomeMonthly.HeaderText = "Sum Of Monthly Gross Income"
        Me.ColGrossIncomeMonthly.Name = "ColGrossIncomeMonthly"
        Me.ColGrossIncomeMonthly.ReadOnly = True
        Me.ColGrossIncomeMonthly.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColGrossIncomeMonthly.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColGrossIncomeMonthly.Width = 205
        '
        'ColGrossIncome_PrevIncomeTax
        '
        Me.ColGrossIncome_PrevIncomeTax.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColGrossIncome_PrevIncomeTax.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColGrossIncome_PrevIncomeTax.HeaderText = "Prev. Period Income Tax"
        Me.ColGrossIncome_PrevIncomeTax.Name = "ColGrossIncome_PrevIncomeTax"
        Me.ColGrossIncome_PrevIncomeTax.ReadOnly = True
        Me.ColGrossIncome_PrevIncomeTax.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColGrossIncome_PrevIncomeTax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColGrossIncome_PrevIncomeTax.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColGrossIncome_PrevIncomeTax.Width = 187
        '
        'ColGrossIncomeYearly
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColGrossIncomeYearly.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColGrossIncomeYearly.HeaderText = "End Of Year Gross Income"
        Me.ColGrossIncomeYearly.Name = "ColGrossIncomeYearly"
        Me.ColGrossIncomeYearly.ReadOnly = True
        Me.ColGrossIncomeYearly.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColGrossIncomeYearly.Width = 183
        '
        'ColIncomeTaxMonthly
        '
        Me.ColIncomeTaxMonthly.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColIncomeTaxMonthly.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColIncomeTaxMonthly.HeaderText = "Sum Of Monthly Income Tax"
        Me.ColIncomeTaxMonthly.Name = "ColIncomeTaxMonthly"
        Me.ColIncomeTaxMonthly.ReadOnly = True
        Me.ColIncomeTaxMonthly.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIncomeTaxMonthly.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColIncomeTaxMonthly.Width = 190
        '
        'ColIncomeTaxYearly
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColIncomeTaxYearly.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColIncomeTaxYearly.HeaderText = "End Of Year Income Tax"
        Me.ColIncomeTaxYearly.Name = "ColIncomeTaxYearly"
        Me.ColIncomeTaxYearly.ReadOnly = True
        Me.ColIncomeTaxYearly.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIncomeTaxYearly.Width = 168
        '
        'ColA1FormNo
        '
        Me.ColA1FormNo.HeaderText = "A1 Form No"
        Me.ColA1FormNo.Name = "ColA1FormNo"
        Me.ColA1FormNo.ReadOnly = True
        Me.ColA1FormNo.Width = 108
        '
        'ColEmployer
        '
        Me.ColEmployer.HeaderText = "Employer"
        Me.ColEmployer.Name = "ColEmployer"
        Me.ColEmployer.ReadOnly = True
        Me.ColEmployer.Width = 92
        '
        'ColPrimaryEmployee
        '
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.ReadOnly = True
        Me.ColPrimaryEmployee.Visible = False
        Me.ColPrimaryEmployee.Width = 147
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(434, 12)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ChkShowWrongDataOnly
        '
        Me.ChkShowWrongDataOnly.AutoSize = True
        Me.ChkShowWrongDataOnly.Location = New System.Drawing.Point(515, 14)
        Me.ChkShowWrongDataOnly.Name = "ChkShowWrongDataOnly"
        Me.ChkShowWrongDataOnly.Size = New System.Drawing.Size(158, 21)
        Me.ChkShowWrongDataOnly.TabIndex = 3
        Me.ChkShowWrongDataOnly.Text = "Show exception only"
        Me.ChkShowWrongDataOnly.UseVisualStyleBackColor = True
        '
        'FrmTrans_Export_eSPT_Check
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(713, 609)
        Me.Controls.Add(Me.ChkShowWrongDataOnly)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.TxtCheckDesc)
        Me.Controls.Add(Me.cboCheckType)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnStart)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(731, 656)
        Me.Name = "FrmTrans_Export_eSPT_Check"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Yearly Data Check"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboCheckType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtCheckDesc As System.Windows.Forms.TextBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ChkShowWrongDataOnly As System.Windows.Forms.CheckBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGrossIncomeMonthly As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColGrossIncome_PrevIncomeTax As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColGrossIncomeYearly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncomeTaxMonthly As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColIncomeTaxYearly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColA1FormNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
