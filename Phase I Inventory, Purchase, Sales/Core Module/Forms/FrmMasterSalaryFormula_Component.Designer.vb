<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalaryFormula_Component
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMasterSalaryFormula_Component))
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvComponent = New System.Windows.Forms.DataGridView
        Me.ColIncome_ComponentTipe = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColIncome_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_Routine = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColIncome_SlipDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_TaxComponent = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColIncome_ExpenseAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_PaymentAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFormula = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_CalculatePeriod = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColIncome_CalculateOn = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColEditFormula = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColIncome_PrimaryExpenseAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_Urutan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_PrimarySlipDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRoundingUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRoundingType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYearlyMaxValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYearlyMinValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColYearlyDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMonthlyMaxValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMonthlyMinValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Colincome_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_PrimaryPaymentAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_Duplicate = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColIncome_CreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_CreatedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_ModifiedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncome_ModifiedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColEmp_EmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_EmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_Basic = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_TaxCalculation = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColEmp_SalarySlipDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_TaxComponent = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColEmp_ExpenseAcc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_ExpenseAccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_PaymentAcc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_PaymentAccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_OtherDeduction = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColEmp_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_PrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_PrimarySlipDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_PrimaryExpenseAcc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_PrimaryPayment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_CreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_CreatedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_ModifedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_ModifiedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmp_IsModify = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker_CloseForm = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.btnSave = New System.Windows.Forms.Button
        Me.llbImport = New System.Windows.Forms.LinkLabel
        Me.llbExport = New System.Windows.Forms.LinkLabel
        Me.SaveDia = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvComponent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(651, 442)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(570, 442)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(714, 424)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvComponent)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(706, 398)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Salary Component"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvComponent
        '
        Me.DgvComponent.AllowUserToResizeColumns = False
        Me.DgvComponent.AllowUserToResizeRows = False
        Me.DgvComponent.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvComponent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvComponent.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColIncome_ComponentTipe, Me.ColIncome_Desc, Me.ColIncome_Routine, Me.ColIncome_SlipDescription, Me.ColIncome_TaxComponent, Me.ColIncome_ExpenseAccount, Me.ColIncome_PaymentAccount, Me.ColIncome_Identifier, Me.ColFormula, Me.ColIncome_CalculatePeriod, Me.ColIncome_CalculateOn, Me.ColEditFormula, Me.ColIncome_PrimaryExpenseAccount, Me.ColIncome_Urutan, Me.ColIncome_PrimarySlipDesc, Me.ColRoundingUnit, Me.ColRoundingType, Me.ColYearlyMaxValue, Me.ColYearlyMinValue, Me.ColYearlyDate, Me.ColMonthlyMaxValue, Me.ColMonthlyMinValue, Me.Colincome_Primary, Me.ColIncome_PrimaryPaymentAccount, Me.ColIncome_Duplicate, Me.ColIncome_CreatedBy, Me.ColIncome_CreatedDate, Me.ColIncome_ModifiedBy, Me.ColIncome_ModifiedDate, Me.ColIsModify})
        Me.DgvComponent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvComponent.Location = New System.Drawing.Point(3, 3)
        Me.DgvComponent.MultiSelect = False
        Me.DgvComponent.Name = "DgvComponent"
        Me.DgvComponent.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvComponent.RowTemplate.Height = 24
        Me.DgvComponent.Size = New System.Drawing.Size(700, 392)
        Me.DgvComponent.TabIndex = 0
        '
        'ColIncome_ComponentTipe
        '
        Me.ColIncome_ComponentTipe.Frozen = True
        Me.ColIncome_ComponentTipe.HeaderText = "Type"
        Me.ColIncome_ComponentTipe.Items.AddRange(New Object() {"Income", "Deduction", "Company Expense", "Employee Expense"})
        Me.ColIncome_ComponentTipe.Name = "ColIncome_ComponentTipe"
        Me.ColIncome_ComponentTipe.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIncome_ComponentTipe.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColIncome_ComponentTipe.Width = 56
        '
        'ColIncome_Desc
        '
        Me.ColIncome_Desc.Frozen = True
        Me.ColIncome_Desc.HeaderText = "Component Name"
        Me.ColIncome_Desc.Name = "ColIncome_Desc"
        Me.ColIncome_Desc.Width = 117
        '
        'ColIncome_Routine
        '
        Me.ColIncome_Routine.HeaderText = "Routine/Annualized"
        Me.ColIncome_Routine.Name = "ColIncome_Routine"
        Me.ColIncome_Routine.Width = 107
        '
        'ColIncome_SlipDescription
        '
        Me.ColIncome_SlipDescription.HeaderText = "Salary Slip Description"
        Me.ColIncome_SlipDescription.Name = "ColIncome_SlipDescription"
        Me.ColIncome_SlipDescription.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIncome_SlipDescription.Width = 137
        '
        'ColIncome_TaxComponent
        '
        Me.ColIncome_TaxComponent.HeaderText = "Tax Component"
        Me.ColIncome_TaxComponent.Items.AddRange(New Object() {"GAJI/PENSIUN ATAU THT/JHT", "TUNJANGAN LAINNYA, UANG LEMBUR DAN SEBAGAINYA", "HONORARIUM DAN IMBALAN LAIN SEJENISNYA", "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA", "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN P" & _
                        "Ph PASAL 21", "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR", "IURAN PENSIUN ATAU IURAN THT/JHT", "BUKAN OBJEK PAJAK", "PESANGON"})
        Me.ColIncome_TaxComponent.Name = "ColIncome_TaxComponent"
        Me.ColIncome_TaxComponent.Width = 88
        '
        'ColIncome_ExpenseAccount
        '
        Me.ColIncome_ExpenseAccount.HeaderText = "Expense Account"
        Me.ColIncome_ExpenseAccount.Name = "ColIncome_ExpenseAccount"
        Me.ColIncome_ExpenseAccount.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIncome_ExpenseAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIncome_ExpenseAccount.Width = 97
        '
        'ColIncome_PaymentAccount
        '
        Me.ColIncome_PaymentAccount.HeaderText = "Payment Account"
        Me.ColIncome_PaymentAccount.Name = "ColIncome_PaymentAccount"
        Me.ColIncome_PaymentAccount.Width = 116
        '
        'ColIncome_Identifier
        '
        Me.ColIncome_Identifier.HeaderText = "Identifier"
        Me.ColIncome_Identifier.Name = "ColIncome_Identifier"
        Me.ColIncome_Identifier.Width = 72
        '
        'ColFormula
        '
        Me.ColFormula.HeaderText = "Formula"
        Me.ColFormula.Name = "ColFormula"
        Me.ColFormula.ReadOnly = True
        Me.ColFormula.Width = 69
        '
        'ColIncome_CalculatePeriod
        '
        Me.ColIncome_CalculatePeriod.HeaderText = "Calculation Period"
        Me.ColIncome_CalculatePeriod.Items.AddRange(New Object() {"Daily", "Monthly", "Yearly"})
        Me.ColIncome_CalculatePeriod.Name = "ColIncome_CalculatePeriod"
        Me.ColIncome_CalculatePeriod.Width = 98
        '
        'ColIncome_CalculateOn
        '
        Me.ColIncome_CalculateOn.HeaderText = "Calculate On"
        Me.ColIncome_CalculateOn.Items.AddRange(New Object() {"Jan", "Feb", "Mar", "Apr", "May", "June", "Jul", "Agt", "Sept", "Oct", "Nov", "Dec"})
        Me.ColIncome_CalculateOn.Name = "ColIncome_CalculateOn"
        Me.ColIncome_CalculateOn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIncome_CalculateOn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColIncome_CalculateOn.Width = 93
        '
        'ColEditFormula
        '
        Me.ColEditFormula.HeaderText = "Edit Formula"
        Me.ColEditFormula.Name = "ColEditFormula"
        Me.ColEditFormula.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColEditFormula.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColEditFormula.Width = 90
        '
        'ColIncome_PrimaryExpenseAccount
        '
        Me.ColIncome_PrimaryExpenseAccount.HeaderText = "Primary Expense Acc"
        Me.ColIncome_PrimaryExpenseAccount.Name = "ColIncome_PrimaryExpenseAccount"
        Me.ColIncome_PrimaryExpenseAccount.Visible = False
        Me.ColIncome_PrimaryExpenseAccount.Width = 166
        '
        'ColIncome_Urutan
        '
        Me.ColIncome_Urutan.HeaderText = "Urutan"
        Me.ColIncome_Urutan.Name = "ColIncome_Urutan"
        Me.ColIncome_Urutan.Visible = False
        Me.ColIncome_Urutan.Width = 76
        '
        'ColIncome_PrimarySlipDesc
        '
        Me.ColIncome_PrimarySlipDesc.HeaderText = "Primary Slip Desc"
        Me.ColIncome_PrimarySlipDesc.Name = "ColIncome_PrimarySlipDesc"
        Me.ColIncome_PrimarySlipDesc.Visible = False
        Me.ColIncome_PrimarySlipDesc.Width = 144
        '
        'ColRoundingUnit
        '
        Me.ColRoundingUnit.HeaderText = "RoundingUnit"
        Me.ColRoundingUnit.Name = "ColRoundingUnit"
        Me.ColRoundingUnit.Visible = False
        Me.ColRoundingUnit.Width = 119
        '
        'ColRoundingType
        '
        Me.ColRoundingType.HeaderText = "RoundingType"
        Me.ColRoundingType.Name = "ColRoundingType"
        Me.ColRoundingType.Visible = False
        Me.ColRoundingType.Width = 126
        '
        'ColYearlyMaxValue
        '
        Me.ColYearlyMaxValue.HeaderText = "YearlyMaxValue"
        Me.ColYearlyMaxValue.Name = "ColYearlyMaxValue"
        Me.ColYearlyMaxValue.Visible = False
        Me.ColYearlyMaxValue.Width = 134
        '
        'ColYearlyMinValue
        '
        Me.ColYearlyMinValue.HeaderText = "YearlyMinValue"
        Me.ColYearlyMinValue.Name = "ColYearlyMinValue"
        Me.ColYearlyMinValue.Visible = False
        Me.ColYearlyMinValue.Width = 131
        '
        'ColYearlyDate
        '
        Me.ColYearlyDate.HeaderText = "YearlyDate"
        Me.ColYearlyDate.Name = "ColYearlyDate"
        Me.ColYearlyDate.Visible = False
        Me.ColYearlyDate.Width = 103
        '
        'ColMonthlyMaxValue
        '
        Me.ColMonthlyMaxValue.HeaderText = "MonthlyMaxValue"
        Me.ColMonthlyMaxValue.Name = "ColMonthlyMaxValue"
        Me.ColMonthlyMaxValue.Visible = False
        Me.ColMonthlyMaxValue.Width = 143
        '
        'ColMonthlyMinValue
        '
        Me.ColMonthlyMinValue.HeaderText = "MonthlyMinValue"
        Me.ColMonthlyMinValue.Name = "ColMonthlyMinValue"
        Me.ColMonthlyMinValue.Visible = False
        Me.ColMonthlyMinValue.Width = 140
        '
        'Colincome_Primary
        '
        Me.Colincome_Primary.HeaderText = "Primary"
        Me.Colincome_Primary.Name = "Colincome_Primary"
        Me.Colincome_Primary.Visible = False
        Me.Colincome_Primary.Width = 81
        '
        'ColIncome_PrimaryPaymentAccount
        '
        Me.ColIncome_PrimaryPaymentAccount.HeaderText = "Primary Payment Account"
        Me.ColIncome_PrimaryPaymentAccount.Name = "ColIncome_PrimaryPaymentAccount"
        Me.ColIncome_PrimaryPaymentAccount.Visible = False
        Me.ColIncome_PrimaryPaymentAccount.Width = 195
        '
        'ColIncome_Duplicate
        '
        Me.ColIncome_Duplicate.HeaderText = "Duplicate"
        Me.ColIncome_Duplicate.Name = "ColIncome_Duplicate"
        Me.ColIncome_Duplicate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIncome_Duplicate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColIncome_Duplicate.Width = 77
        '
        'ColIncome_CreatedBy
        '
        Me.ColIncome_CreatedBy.HeaderText = "Created By"
        Me.ColIncome_CreatedBy.Name = "ColIncome_CreatedBy"
        Me.ColIncome_CreatedBy.Width = 84
        '
        'ColIncome_CreatedDate
        '
        Me.ColIncome_CreatedDate.HeaderText = "Created Date"
        Me.ColIncome_CreatedDate.Name = "ColIncome_CreatedDate"
        Me.ColIncome_CreatedDate.Width = 95
        '
        'ColIncome_ModifiedBy
        '
        Me.ColIncome_ModifiedBy.HeaderText = "Modified By"
        Me.ColIncome_ModifiedBy.Name = "ColIncome_ModifiedBy"
        Me.ColIncome_ModifiedBy.Width = 87
        '
        'ColIncome_ModifiedDate
        '
        Me.ColIncome_ModifiedDate.HeaderText = "Modified Date"
        Me.ColIncome_ModifiedDate.Name = "ColIncome_ModifiedDate"
        Me.ColIncome_ModifiedDate.Width = 98
        '
        'ColIsModify
        '
        Me.ColIsModify.HeaderText = "IsModify"
        Me.ColIsModify.Name = "ColIsModify"
        Me.ColIsModify.Visible = False
        Me.ColIsModify.Width = 65
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvEmployee)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(712, 399)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Detail Employee"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AllowUserToResizeColumns = False
        Me.dgvEmployee.AllowUserToResizeRows = False
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColEmp_EmployeeCode, Me.ColEmp_EmployeeName, Me.ColEmp_Basic, Me.ColEmp_TaxCalculation, Me.ColEmp_SalarySlipDesc, Me.ColEmp_TaxComponent, Me.ColEmp_ExpenseAcc, Me.ColEmp_ExpenseAccDesc, Me.ColEmp_PaymentAcc, Me.ColEmp_PaymentAccDesc, Me.ColEmp_OtherDeduction, Me.ColEmp_Remarks, Me.ColEmp_PrimaryEmployee, Me.ColEmp_PrimarySlipDesc, Me.ColEmp_PrimaryExpenseAcc, Me.ColEmp_PrimaryPayment, Me.ColEmp_Primary, Me.ColEmp_CreatedBy, Me.ColEmp_CreatedDate, Me.ColEmp_ModifedBy, Me.ColEmp_ModifiedDate, Me.ColEmp_IsModify})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(3, 3)
        Me.dgvEmployee.MultiSelect = False
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(706, 393)
        Me.dgvEmployee.TabIndex = 0
        '
        'ColEmp_EmployeeCode
        '
        Me.ColEmp_EmployeeCode.HeaderText = "Code"
        Me.ColEmp_EmployeeCode.Name = "ColEmp_EmployeeCode"
        Me.ColEmp_EmployeeCode.Width = 57
        '
        'ColEmp_EmployeeName
        '
        Me.ColEmp_EmployeeName.HeaderText = "Name"
        Me.ColEmp_EmployeeName.Name = "ColEmp_EmployeeName"
        Me.ColEmp_EmployeeName.Width = 60
        '
        'ColEmp_Basic
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColEmp_Basic.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColEmp_Basic.HeaderText = "Basic Salary"
        Me.ColEmp_Basic.Name = "ColEmp_Basic"
        Me.ColEmp_Basic.Width = 90
        '
        'ColEmp_TaxCalculation
        '
        Me.ColEmp_TaxCalculation.HeaderText = "Tax Calculation"
        Me.ColEmp_TaxCalculation.Items.AddRange(New Object() {"Gross Up", "Netto", "Gross"})
        Me.ColEmp_TaxCalculation.Name = "ColEmp_TaxCalculation"
        Me.ColEmp_TaxCalculation.Width = 86
        '
        'ColEmp_SalarySlipDesc
        '
        Me.ColEmp_SalarySlipDesc.HeaderText = "Salary Slip Description"
        Me.ColEmp_SalarySlipDesc.Name = "ColEmp_SalarySlipDesc"
        Me.ColEmp_SalarySlipDesc.Width = 137
        '
        'ColEmp_TaxComponent
        '
        Me.ColEmp_TaxComponent.HeaderText = "Tax Component"
        Me.ColEmp_TaxComponent.Items.AddRange(New Object() {"GAJI/PENSIUN ATAU THT/JHT"})
        Me.ColEmp_TaxComponent.Name = "ColEmp_TaxComponent"
        Me.ColEmp_TaxComponent.ReadOnly = True
        Me.ColEmp_TaxComponent.Width = 88
        '
        'ColEmp_ExpenseAcc
        '
        Me.ColEmp_ExpenseAcc.HeaderText = "Expense Acc."
        Me.ColEmp_ExpenseAcc.Name = "ColEmp_ExpenseAcc"
        Me.ColEmp_ExpenseAcc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColEmp_ExpenseAcc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmp_ExpenseAcc.Width = 79
        '
        'ColEmp_ExpenseAccDesc
        '
        Me.ColEmp_ExpenseAccDesc.HeaderText = "Expense Acc. Description"
        Me.ColEmp_ExpenseAccDesc.Name = "ColEmp_ExpenseAccDesc"
        Me.ColEmp_ExpenseAccDesc.ReadOnly = True
        Me.ColEmp_ExpenseAccDesc.Width = 154
        '
        'ColEmp_PaymentAcc
        '
        Me.ColEmp_PaymentAcc.HeaderText = "Default Payment Acc."
        Me.ColEmp_PaymentAcc.Name = "ColEmp_PaymentAcc"
        Me.ColEmp_PaymentAcc.Width = 135
        '
        'ColEmp_PaymentAccDesc
        '
        Me.ColEmp_PaymentAccDesc.HeaderText = "Default Payment Acc. Description"
        Me.ColEmp_PaymentAccDesc.Name = "ColEmp_PaymentAccDesc"
        Me.ColEmp_PaymentAccDesc.ReadOnly = True
        Me.ColEmp_PaymentAccDesc.Width = 191
        '
        'ColEmp_OtherDeduction
        '
        Me.ColEmp_OtherDeduction.HeaderText = "Other Deduction"
        Me.ColEmp_OtherDeduction.Name = "ColEmp_OtherDeduction"
        Me.ColEmp_OtherDeduction.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColEmp_OtherDeduction.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColEmp_OtherDeduction.Text = "Other Deduction"
        Me.ColEmp_OtherDeduction.ToolTipText = "Other Deduction"
        Me.ColEmp_OtherDeduction.Width = 110
        '
        'ColEmp_Remarks
        '
        Me.ColEmp_Remarks.HeaderText = "Remarks"
        Me.ColEmp_Remarks.Name = "ColEmp_Remarks"
        Me.ColEmp_Remarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColEmp_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmp_Remarks.Width = 55
        '
        'ColEmp_PrimaryEmployee
        '
        Me.ColEmp_PrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColEmp_PrimaryEmployee.Name = "ColEmp_PrimaryEmployee"
        Me.ColEmp_PrimaryEmployee.Visible = False
        Me.ColEmp_PrimaryEmployee.Width = 147
        '
        'ColEmp_PrimarySlipDesc
        '
        Me.ColEmp_PrimarySlipDesc.HeaderText = "PrimarySlipDesc"
        Me.ColEmp_PrimarySlipDesc.Name = "ColEmp_PrimarySlipDesc"
        Me.ColEmp_PrimarySlipDesc.Visible = False
        Me.ColEmp_PrimarySlipDesc.Width = 136
        '
        'ColEmp_PrimaryExpenseAcc
        '
        Me.ColEmp_PrimaryExpenseAcc.HeaderText = "PrimaryExpenseAcc"
        Me.ColEmp_PrimaryExpenseAcc.Name = "ColEmp_PrimaryExpenseAcc"
        Me.ColEmp_PrimaryExpenseAcc.Visible = False
        Me.ColEmp_PrimaryExpenseAcc.Width = 158
        '
        'ColEmp_PrimaryPayment
        '
        Me.ColEmp_PrimaryPayment.HeaderText = "Primary Payment"
        Me.ColEmp_PrimaryPayment.Name = "ColEmp_PrimaryPayment"
        Me.ColEmp_PrimaryPayment.Visible = False
        Me.ColEmp_PrimaryPayment.Width = 140
        '
        'ColEmp_Primary
        '
        Me.ColEmp_Primary.HeaderText = "ColEmp_Primary"
        Me.ColEmp_Primary.Name = "ColEmp_Primary"
        Me.ColEmp_Primary.Visible = False
        Me.ColEmp_Primary.Width = 137
        '
        'ColEmp_CreatedBy
        '
        Me.ColEmp_CreatedBy.HeaderText = "Created By"
        Me.ColEmp_CreatedBy.Name = "ColEmp_CreatedBy"
        Me.ColEmp_CreatedBy.Width = 84
        '
        'ColEmp_CreatedDate
        '
        Me.ColEmp_CreatedDate.HeaderText = "Created Date"
        Me.ColEmp_CreatedDate.Name = "ColEmp_CreatedDate"
        Me.ColEmp_CreatedDate.Width = 95
        '
        'ColEmp_ModifedBy
        '
        Me.ColEmp_ModifedBy.HeaderText = "Modified By"
        Me.ColEmp_ModifedBy.Name = "ColEmp_ModifedBy"
        Me.ColEmp_ModifedBy.Width = 87
        '
        'ColEmp_ModifiedDate
        '
        Me.ColEmp_ModifiedDate.HeaderText = "Modified Date"
        Me.ColEmp_ModifiedDate.Name = "ColEmp_ModifiedDate"
        Me.ColEmp_ModifiedDate.Width = 98
        '
        'ColEmp_IsModify
        '
        Me.ColEmp_IsModify.HeaderText = "IsModify"
        Me.ColEmp_IsModify.Name = "ColEmp_IsModify"
        Me.ColEmp_IsModify.Visible = False
        Me.ColEmp_IsModify.Width = 84
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(80, 449)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(112, 13)
        Me.LinkLabel1.TabIndex = 3
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Salary Slip Description"
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(12, 442)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 1
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(46, 442)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 2
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(198, 449)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(146, 13)
        Me.LinkLabel2.TabIndex = 4
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Copy Data Salary Component"
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
        'BackgroundWorker_CloseForm
        '
        '
        'BackgroundWorker4
        '
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(489, 442)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'llbImport
        '
        Me.llbImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbImport.AutoSize = True
        Me.llbImport.Location = New System.Drawing.Point(434, 449)
        Me.llbImport.Name = "llbImport"
        Me.llbImport.Size = New System.Drawing.Size(36, 13)
        Me.llbImport.TabIndex = 6
        Me.llbImport.TabStop = True
        Me.llbImport.Text = "Import"
        '
        'llbExport
        '
        Me.llbExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbExport.AutoSize = True
        Me.llbExport.Location = New System.Drawing.Point(392, 449)
        Me.llbExport.Name = "llbExport"
        Me.llbExport.Size = New System.Drawing.Size(37, 13)
        Me.llbExport.TabIndex = 5
        Me.llbExport.TabStop = True
        Me.llbExport.Text = "Export"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'FrmMasterSalaryFormula_Component
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(740, 477)
        Me.Controls.Add(Me.llbExport)
        Me.Controls.Add(Me.llbImport)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.btnUP)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(756, 515)
        Me.Name = "FrmMasterSalaryFormula_Component"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Salary Detail"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvComponent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents DgvComponent As System.Windows.Forms.DataGridView
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker_CloseForm As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents ColEmp_EmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_EmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_Basic As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_TaxCalculation As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColEmp_SalarySlipDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_TaxComponent As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColEmp_ExpenseAcc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_ExpenseAccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_PaymentAcc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_PaymentAccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_OtherDeduction As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColEmp_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_PrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_PrimarySlipDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_PrimaryExpenseAcc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_PrimaryPayment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_CreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_CreatedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_ModifedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_ModifiedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmp_IsModify As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_ComponentTipe As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColIncome_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_Routine As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColIncome_SlipDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_TaxComponent As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColIncome_ExpenseAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_PaymentAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFormula As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_CalculatePeriod As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColIncome_CalculateOn As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColEditFormula As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColIncome_PrimaryExpenseAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_Urutan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_PrimarySlipDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRoundingUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRoundingType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYearlyMaxValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYearlyMinValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColYearlyDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMonthlyMaxValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMonthlyMinValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Colincome_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_PrimaryPaymentAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_Duplicate As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColIncome_CreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_CreatedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_ModifiedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncome_ModifiedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents llbImport As System.Windows.Forms.LinkLabel
    Friend WithEvents llbExport As System.Windows.Forms.LinkLabel
    Friend WithEvents SaveDia As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
End Class
