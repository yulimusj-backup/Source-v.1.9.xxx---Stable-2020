<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_Override
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.LblAdvance = New System.Windows.Forms.LinkLabel
        Me.lblImport = New System.Windows.Forms.LinkLabel
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.LblQuickInput = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartmentCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartmentDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOverrideAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOverrideReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.lblExport = New System.Windows.Forms.LinkLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtDocNumber = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtPrimarySalaryFormula = New System.Windows.Forms.TextBox
        Me.CboFormulaType = New System.Windows.Forms.ComboBox
        Me.TxtIdentifierSalaryFormula = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtComponentName = New System.Windows.Forms.TextBox
        Me.BtnFindComponentName = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtFormulaCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtFormulaDesc = New System.Windows.Forms.TextBox
        Me.TxtSubTotal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllEmployeeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyByDepartementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyToAJobPositionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CopyFromOtherDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CopyFromSalaryCalculationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BackgroundWorker_DetailEmployee = New System.ComponentModel.BackgroundWorker
        Me.FrmDia = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.LblAdvance)
        Me.gboDetail.Controls.Add(Me.lblImport)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.lblExport)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(14, -1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(828, 446)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'LblAdvance
        '
        Me.LblAdvance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAdvance.AutoSize = True
        Me.LblAdvance.Location = New System.Drawing.Point(769, 98)
        Me.LblAdvance.Name = "LblAdvance"
        Me.LblAdvance.Size = New System.Drawing.Size(50, 13)
        Me.LblAdvance.TabIndex = 2
        Me.LblAdvance.TabStop = True
        Me.LblAdvance.Text = "Advance"
        '
        'lblImport
        '
        Me.lblImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblImport.AutoSize = True
        Me.lblImport.Location = New System.Drawing.Point(727, 98)
        Me.lblImport.Name = "lblImport"
        Me.lblImport.Size = New System.Drawing.Size(36, 13)
        Me.lblImport.TabIndex = 1
        Me.lblImport.TabStop = True
        Me.lblImport.Text = "Import"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(6, 116)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblQuickInput)
        Me.SplitContainer2.Panel1Collapsed = True
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.dgvItem)
        Me.SplitContainer2.Size = New System.Drawing.Size(816, 292)
        Me.SplitContainer2.SplitterDistance = 39
        Me.SplitContainer2.TabIndex = 36
        '
        'LblQuickInput
        '
        Me.LblQuickInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblQuickInput.AutoSize = True
        Me.LblQuickInput.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblQuickInput.ForeColor = System.Drawing.Color.Blue
        Me.LblQuickInput.Location = New System.Drawing.Point(741, 0)
        Me.LblQuickInput.Name = "LblQuickInput"
        Me.LblQuickInput.Size = New System.Drawing.Size(62, 13)
        Me.LblQuickInput.TabIndex = 56
        Me.LblQuickInput.Text = "Quick Input"
        '
        'dgvItem
        '
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColDepartmentCode, Me.ColDepartmentDesc, Me.ColJobPositionCode, Me.ColJobPositionDescription, Me.ColOverrideAmount, Me.ColOverrideReason, Me.ColPrimaryEmployee})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(816, 292)
        Me.dgvItem.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.Width = 106
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.Width = 109
        '
        'ColDepartmentCode
        '
        Me.ColDepartmentCode.HeaderText = "Department Code"
        Me.ColDepartmentCode.Name = "ColDepartmentCode"
        Me.ColDepartmentCode.Visible = False
        Me.ColDepartmentCode.Width = 115
        '
        'ColDepartmentDesc
        '
        Me.ColDepartmentDesc.HeaderText = "Department Description"
        Me.ColDepartmentDesc.Name = "ColDepartmentDesc"
        Me.ColDepartmentDesc.Width = 143
        '
        'ColJobPositionCode
        '
        Me.ColJobPositionCode.HeaderText = "Job Position Code"
        Me.ColJobPositionCode.Name = "ColJobPositionCode"
        Me.ColJobPositionCode.Visible = False
        Me.ColJobPositionCode.Width = 117
        '
        'ColJobPositionDescription
        '
        Me.ColJobPositionDescription.HeaderText = "Job Position Description"
        Me.ColJobPositionDescription.Name = "ColJobPositionDescription"
        Me.ColJobPositionDescription.Width = 145
        '
        'ColOverrideAmount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOverrideAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColOverrideAmount.HeaderText = "Override Amount"
        Me.ColOverrideAmount.MaxInputLength = 12
        Me.ColOverrideAmount.Name = "ColOverrideAmount"
        Me.ColOverrideAmount.Width = 111
        '
        'ColOverrideReason
        '
        Me.ColOverrideReason.HeaderText = "Override Reason"
        Me.ColOverrideReason.MaxInputLength = 250
        Me.ColOverrideReason.Name = "ColOverrideReason"
        Me.ColOverrideReason.Width = 112
        '
        'ColPrimaryEmployee
        '
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.Visible = False
        Me.ColPrimaryEmployee.Width = 115
        '
        'lblExport
        '
        Me.lblExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblExport.AutoSize = True
        Me.lblExport.Location = New System.Drawing.Point(685, 98)
        Me.lblExport.Name = "lblExport"
        Me.lblExport.Size = New System.Drawing.Size(37, 13)
        Me.lblExport.TabIndex = 0
        Me.lblExport.TabStop = True
        Me.lblExport.Text = "Export"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 11)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDocNumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimarySalaryFormula)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboFormulaType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtIdentifierSalaryFormula)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtComponentName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindComponentName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtFormulaCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtFormulaDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtSubTotal)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Size = New System.Drawing.Size(816, 83)
        Me.SplitContainer1.SplitterDistance = 398
        Me.SplitContainer1.TabIndex = 0
        '
        'TxtDocNumber
        '
        Me.TxtDocNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocNumber.Location = New System.Drawing.Point(106, 11)
        Me.TxtDocNumber.MaxLength = 50
        Me.TxtDocNumber.Name = "TxtDocNumber"
        Me.TxtDocNumber.ReadOnly = True
        Me.TxtDocNumber.Size = New System.Drawing.Size(144, 20)
        Me.TxtDocNumber.TabIndex = 0
        Me.TxtDocNumber.TabStop = False
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(62, 11)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(289, 11)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(100, 20)
        Me.DtpDate.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Document No"
        '
        'Label11
        '
        Me.Label11.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(256, 13)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(30, 13)
        Me.Label11.TabIndex = 50
        Me.Label11.Text = "Date"
        '
        'TxtPrimarySalaryFormula
        '
        Me.TxtPrimarySalaryFormula.Location = New System.Drawing.Point(9, 80)
        Me.TxtPrimarySalaryFormula.Name = "TxtPrimarySalaryFormula"
        Me.TxtPrimarySalaryFormula.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimarySalaryFormula.TabIndex = 47
        Me.TxtPrimarySalaryFormula.TabStop = False
        Me.TxtPrimarySalaryFormula.Text = "-1"
        Me.TxtPrimarySalaryFormula.Visible = False
        '
        'CboFormulaType
        '
        Me.CboFormulaType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboFormulaType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormulaType.FormattingEnabled = True
        Me.CboFormulaType.Items.AddRange(New Object() {"Salary Formula"})
        Me.CboFormulaType.Location = New System.Drawing.Point(106, 35)
        Me.CboFormulaType.Name = "CboFormulaType"
        Me.CboFormulaType.Size = New System.Drawing.Size(284, 21)
        Me.CboFormulaType.TabIndex = 2
        '
        'TxtIdentifierSalaryFormula
        '
        Me.TxtIdentifierSalaryFormula.Location = New System.Drawing.Point(43, 80)
        Me.TxtIdentifierSalaryFormula.Name = "TxtIdentifierSalaryFormula"
        Me.TxtIdentifierSalaryFormula.Size = New System.Drawing.Size(28, 20)
        Me.TxtIdentifierSalaryFormula.TabIndex = 35
        Me.TxtIdentifierSalaryFormula.TabStop = False
        Me.TxtIdentifierSalaryFormula.Text = "-1-1"
        Me.TxtIdentifierSalaryFormula.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 38)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 13)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Formula Type"
        '
        'TxtComponentName
        '
        Me.TxtComponentName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtComponentName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtComponentName.Location = New System.Drawing.Point(106, 61)
        Me.TxtComponentName.MaxLength = 100
        Me.TxtComponentName.Name = "TxtComponentName"
        Me.TxtComponentName.Size = New System.Drawing.Size(224, 20)
        Me.TxtComponentName.TabIndex = 3
        '
        'BtnFindComponentName
        '
        Me.BtnFindComponentName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindComponentName.Location = New System.Drawing.Point(336, 61)
        Me.BtnFindComponentName.Name = "BtnFindComponentName"
        Me.BtnFindComponentName.Size = New System.Drawing.Size(55, 20)
        Me.BtnFindComponentName.TabIndex = 4
        Me.BtnFindComponentName.TabStop = False
        Me.BtnFindComponentName.Text = "Find"
        Me.BtnFindComponentName.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 13)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Component Name"
        '
        'TxtFormulaCode
        '
        Me.TxtFormulaCode.Location = New System.Drawing.Point(87, 11)
        Me.TxtFormulaCode.MaxLength = 50
        Me.TxtFormulaCode.Name = "TxtFormulaCode"
        Me.TxtFormulaCode.ReadOnly = True
        Me.TxtFormulaCode.Size = New System.Drawing.Size(75, 20)
        Me.TxtFormulaCode.TabIndex = 0
        Me.TxtFormulaCode.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 13)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Salary Formula"
        '
        'TxtFormulaDesc
        '
        Me.TxtFormulaDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFormulaDesc.Location = New System.Drawing.Point(171, 11)
        Me.TxtFormulaDesc.MaxLength = 50
        Me.TxtFormulaDesc.Name = "TxtFormulaDesc"
        Me.TxtFormulaDesc.ReadOnly = True
        Me.TxtFormulaDesc.Size = New System.Drawing.Size(238, 20)
        Me.TxtFormulaDesc.TabIndex = 1
        Me.TxtFormulaDesc.TabStop = False
        '
        'TxtSubTotal
        '
        Me.TxtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSubTotal.Location = New System.Drawing.Point(87, 36)
        Me.TxtSubTotal.MaxLength = 50
        Me.TxtSubTotal.Name = "TxtSubTotal"
        Me.TxtSubTotal.ReadOnly = True
        Me.TxtSubTotal.Size = New System.Drawing.Size(322, 20)
        Me.TxtSubTotal.TabIndex = 2
        Me.TxtSubTotal.TabStop = False
        Me.TxtSubTotal.Text = "0"
        Me.TxtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 48
        Me.Label4.Text = "Sub Total"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(12, 57)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(19, 20)
        Me.TxtReadOnly.TabIndex = 35
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "Remarks"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(87, 61)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(322, 20)
        Me.TxtRemarks.TabIndex = 3
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(421, 414)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(335, 414)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 478)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(863, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 6
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(212, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(212, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(212, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(212, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'BackgroundWorker1
        '
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(435, 447)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(263, 447)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(607, 447)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(349, 447)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(177, 447)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'BtnPrint
        '
        Me.BtnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnPrint.Location = New System.Drawing.Point(521, 447)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(80, 26)
        Me.BtnPrint.TabIndex = 4
        Me.BtnPrint.Text = "&Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllEmployeeToolStripMenuItem, Me.ApplyByDepartementToolStripMenuItem, Me.ApplyToAJobPositionToolStripMenuItem, Me.CopyFromOtherDataToolStripMenuItem, Me.CopyFromSalaryCalculationToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(231, 114)
        '
        'ApplyToAllEmployeeToolStripMenuItem
        '
        Me.ApplyToAllEmployeeToolStripMenuItem.Name = "ApplyToAllEmployeeToolStripMenuItem"
        Me.ApplyToAllEmployeeToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ApplyToAllEmployeeToolStripMenuItem.Text = "Apply To All Employee"
        '
        'ApplyByDepartementToolStripMenuItem
        '
        Me.ApplyByDepartementToolStripMenuItem.Name = "ApplyByDepartementToolStripMenuItem"
        Me.ApplyByDepartementToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ApplyByDepartementToolStripMenuItem.Text = "Apply To A Department"
        '
        'ApplyToAJobPositionToolStripMenuItem
        '
        Me.ApplyToAJobPositionToolStripMenuItem.Name = "ApplyToAJobPositionToolStripMenuItem"
        Me.ApplyToAJobPositionToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ApplyToAJobPositionToolStripMenuItem.Text = "Apply To A Job Position"
        '
        'CopyFromOtherDataToolStripMenuItem
        '
        Me.CopyFromOtherDataToolStripMenuItem.Name = "CopyFromOtherDataToolStripMenuItem"
        Me.CopyFromOtherDataToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.CopyFromOtherDataToolStripMenuItem.Text = "Copy From Other Data"
        '
        'CopyFromSalaryCalculationToolStripMenuItem
        '
        Me.CopyFromSalaryCalculationToolStripMenuItem.Name = "CopyFromSalaryCalculationToolStripMenuItem"
        Me.CopyFromSalaryCalculationToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.CopyFromSalaryCalculationToolStripMenuItem.Text = "Copy From Salary Calculation"
        '
        'BackgroundWorker_DetailEmployee
        '
        '
        'FrmDia
        '
        Me.FrmDia.DefaultExt = "csv"
        Me.FrmDia.Filter = "Csv Files|*.csv"
        Me.FrmDia.Title = "Save e-SPT data"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'FrmTrans_Gaji_Override
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(863, 500)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimumSize = New System.Drawing.Size(714, 526)
        Me.Name = "FrmTrans_Gaji_Override"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transaction Override Salary Amount"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents BtnFindComponentName As System.Windows.Forms.Button
    Friend WithEvents TxtComponentName As System.Windows.Forms.TextBox
    Friend WithEvents TxtIdentifierSalaryFormula As System.Windows.Forms.TextBox
    Friend WithEvents TxtDocNumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents CboFormulaType As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents LblQuickInput As System.Windows.Forms.Label
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySalaryFormula As System.Windows.Forms.TextBox
    Friend WithEvents LblAdvance As System.Windows.Forms.LinkLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyToAllEmployeeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyByDepartementToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToAJobPositionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyFromOtherDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackgroundWorker_DetailEmployee As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartmentCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartmentDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOverrideAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOverrideReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblImport As System.Windows.Forms.LinkLabel
    Friend WithEvents FrmDia As System.Windows.Forms.SaveFileDialog
    Friend WithEvents lblExport As System.Windows.Forms.LinkLabel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TxtFormulaCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtFormulaDesc As System.Windows.Forms.TextBox
    Friend WithEvents CopyFromSalaryCalculationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
