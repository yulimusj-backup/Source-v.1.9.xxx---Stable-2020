<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkUnit
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
        Me.btnEdit = New System.Windows.Forms.Button
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.BtnFindProductionLine = New System.Windows.Forms.Button
        Me.TxtProductionLineDesc = New System.Windows.Forms.TextBox
        Me.TxtProductionLineCode = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.BtnFindLocationOut = New System.Windows.Forms.Button
        Me.BtnFindLocationIn = New System.Windows.Forms.Button
        Me.txtDescLocationOut = New System.Windows.Forms.TextBox
        Me.txtCodeLocationOut = New System.Windows.Forms.TextBox
        Me.txtDescLocationIn = New System.Windows.Forms.TextBox
        Me.txtCodeLocationIn = New System.Windows.Forms.TextBox
        Me.txtGroupDesc = New System.Windows.Forms.TextBox
        Me.btnFindGroup = New System.Windows.Forms.Button
        Me.txtGroupCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvMachine = New System.Windows.Forms.DataGridView
        Me.ColMachine_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMachine_MachineCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMachine_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMachine_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMachine_PrimaryMachine = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColEmployee_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_ShiftCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_EmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_PrimaryShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvSetuptime = New System.Windows.Forms.DataGridView
        Me.ColSetupTime_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_ProcessCode1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_CodeNo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Desc1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_ProcessCode2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_CodeNo2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Desc2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_SetupTimeDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryProcess1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryCodeNo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryProses2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryCodeNo2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimarySetupTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkActive = New System.Windows.Forms.CheckBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.TxtUnitDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtUnitCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryWorkUnitGroup = New System.Windows.Forms.TextBox
        Me.TxtPrimaryUnit = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocationIn = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocationOut = New System.Windows.Forms.TextBox
        Me.TxtPrimaryProductionLine = New System.Windows.Forms.TextBox
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvMachine, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvSetuptime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(299, 398)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(332, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(127, 398)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(227, 359)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 17
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.BtnFindProductionLine)
        Me.gboDetail.Controls.Add(Me.TxtProductionLineDesc)
        Me.gboDetail.Controls.Add(Me.TxtProductionLineCode)
        Me.gboDetail.Controls.Add(Me.TextBox3)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.BtnFindLocationOut)
        Me.gboDetail.Controls.Add(Me.BtnFindLocationIn)
        Me.gboDetail.Controls.Add(Me.txtDescLocationOut)
        Me.gboDetail.Controls.Add(Me.txtCodeLocationOut)
        Me.gboDetail.Controls.Add(Me.txtDescLocationIn)
        Me.gboDetail.Controls.Add(Me.txtCodeLocationIn)
        Me.gboDetail.Controls.Add(Me.txtGroupDesc)
        Me.gboDetail.Controls.Add(Me.btnFindGroup)
        Me.gboDetail.Controls.Add(Me.txtGroupCode)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.TabControl2)
        Me.gboDetail.Controls.Add(Me.ChkActive)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtUnitDesc)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtUnitCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryWorkUnitGroup)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryUnit)
        Me.gboDetail.Controls.Add(Me.txtPrimaryLocationIn)
        Me.gboDetail.Controls.Add(Me.txtPrimaryLocationOut)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryProductionLine)
        Me.gboDetail.Location = New System.Drawing.Point(14, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(652, 391)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 171)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 13)
        Me.Label7.TabIndex = 78
        Me.Label7.Text = "Prod. Line"
        '
        'BtnFindProductionLine
        '
        Me.BtnFindProductionLine.Location = New System.Drawing.Point(175, 167)
        Me.BtnFindProductionLine.Name = "BtnFindProductionLine"
        Me.BtnFindProductionLine.Size = New System.Drawing.Size(57, 21)
        Me.BtnFindProductionLine.TabIndex = 13
        Me.BtnFindProductionLine.TabStop = False
        Me.BtnFindProductionLine.Text = "...."
        Me.BtnFindProductionLine.UseVisualStyleBackColor = True
        '
        'TxtProductionLineDesc
        '
        Me.TxtProductionLineDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductionLineDesc.Location = New System.Drawing.Point(238, 169)
        Me.TxtProductionLineDesc.MaxLength = 100
        Me.TxtProductionLineDesc.Name = "TxtProductionLineDesc"
        Me.TxtProductionLineDesc.ReadOnly = True
        Me.TxtProductionLineDesc.Size = New System.Drawing.Size(398, 20)
        Me.TxtProductionLineDesc.TabIndex = 14
        Me.TxtProductionLineDesc.TabStop = False
        '
        'TxtProductionLineCode
        '
        Me.TxtProductionLineCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductionLineCode.Location = New System.Drawing.Point(78, 169)
        Me.TxtProductionLineCode.MaxLength = 100
        Me.TxtProductionLineCode.Name = "TxtProductionLineCode"
        Me.TxtProductionLineCode.Size = New System.Drawing.Size(91, 20)
        Me.TxtProductionLineCode.TabIndex = 12
        '
        'TextBox3
        '
        Me.TextBox3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox3.Location = New System.Drawing.Point(83, 169)
        Me.TextBox3.MaxLength = 100
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(28, 20)
        Me.TextBox3.TabIndex = 79
        Me.TextBox3.Text = "-1"
        Me.TextBox3.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 145)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(68, 13)
        Me.Label6.TabIndex = 72
        Me.Label6.Text = "Location Out"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 121)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 71
        Me.Label5.Text = "Location In"
        '
        'BtnFindLocationOut
        '
        Me.BtnFindLocationOut.Location = New System.Drawing.Point(175, 141)
        Me.BtnFindLocationOut.Name = "BtnFindLocationOut"
        Me.BtnFindLocationOut.Size = New System.Drawing.Size(57, 21)
        Me.BtnFindLocationOut.TabIndex = 10
        Me.BtnFindLocationOut.TabStop = False
        Me.BtnFindLocationOut.Text = "...."
        Me.BtnFindLocationOut.UseVisualStyleBackColor = True
        '
        'BtnFindLocationIn
        '
        Me.BtnFindLocationIn.Location = New System.Drawing.Point(175, 117)
        Me.BtnFindLocationIn.Name = "BtnFindLocationIn"
        Me.BtnFindLocationIn.Size = New System.Drawing.Size(57, 21)
        Me.BtnFindLocationIn.TabIndex = 7
        Me.BtnFindLocationIn.TabStop = False
        Me.BtnFindLocationIn.Text = "...."
        Me.BtnFindLocationIn.UseVisualStyleBackColor = True
        '
        'txtDescLocationOut
        '
        Me.txtDescLocationOut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescLocationOut.Location = New System.Drawing.Point(238, 143)
        Me.txtDescLocationOut.MaxLength = 100
        Me.txtDescLocationOut.Name = "txtDescLocationOut"
        Me.txtDescLocationOut.ReadOnly = True
        Me.txtDescLocationOut.Size = New System.Drawing.Size(398, 20)
        Me.txtDescLocationOut.TabIndex = 11
        Me.txtDescLocationOut.TabStop = False
        '
        'txtCodeLocationOut
        '
        Me.txtCodeLocationOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodeLocationOut.Location = New System.Drawing.Point(78, 143)
        Me.txtCodeLocationOut.MaxLength = 100
        Me.txtCodeLocationOut.Name = "txtCodeLocationOut"
        Me.txtCodeLocationOut.Size = New System.Drawing.Size(91, 20)
        Me.txtCodeLocationOut.TabIndex = 9
        '
        'txtDescLocationIn
        '
        Me.txtDescLocationIn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescLocationIn.Location = New System.Drawing.Point(238, 119)
        Me.txtDescLocationIn.MaxLength = 100
        Me.txtDescLocationIn.Name = "txtDescLocationIn"
        Me.txtDescLocationIn.ReadOnly = True
        Me.txtDescLocationIn.Size = New System.Drawing.Size(398, 20)
        Me.txtDescLocationIn.TabIndex = 8
        Me.txtDescLocationIn.TabStop = False
        '
        'txtCodeLocationIn
        '
        Me.txtCodeLocationIn.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodeLocationIn.Location = New System.Drawing.Point(78, 119)
        Me.txtCodeLocationIn.MaxLength = 100
        Me.txtCodeLocationIn.Name = "txtCodeLocationIn"
        Me.txtCodeLocationIn.Size = New System.Drawing.Size(91, 20)
        Me.txtCodeLocationIn.TabIndex = 6
        '
        'txtGroupDesc
        '
        Me.txtGroupDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGroupDesc.Location = New System.Drawing.Point(238, 94)
        Me.txtGroupDesc.MaxLength = 100
        Me.txtGroupDesc.Name = "txtGroupDesc"
        Me.txtGroupDesc.ReadOnly = True
        Me.txtGroupDesc.Size = New System.Drawing.Size(398, 20)
        Me.txtGroupDesc.TabIndex = 5
        Me.txtGroupDesc.TabStop = False
        '
        'btnFindGroup
        '
        Me.btnFindGroup.Location = New System.Drawing.Point(175, 93)
        Me.btnFindGroup.Name = "btnFindGroup"
        Me.btnFindGroup.Size = New System.Drawing.Size(57, 21)
        Me.btnFindGroup.TabIndex = 4
        Me.btnFindGroup.TabStop = False
        Me.btnFindGroup.Text = "...."
        Me.btnFindGroup.UseVisualStyleBackColor = True
        '
        'txtGroupCode
        '
        Me.txtGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtGroupCode.Location = New System.Drawing.Point(78, 93)
        Me.txtGroupCode.MaxLength = 100
        Me.txtGroupCode.Name = "txtGroupCode"
        Me.txtGroupCode.Size = New System.Drawing.Size(91, 20)
        Me.txtGroupCode.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 97)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(36, 13)
        Me.Label4.TabIndex = 61
        Me.Label4.Text = "Group"
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage1)
        Me.TabControl2.Location = New System.Drawing.Point(5, 215)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(640, 140)
        Me.TabControl2.TabIndex = 16
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvMachine)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(632, 114)
        Me.TabPage2.TabIndex = 0
        Me.TabPage2.Text = "Machine"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvMachine
        '
        Me.DgvMachine.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvMachine.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvMachine.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvMachine.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMachine_No, Me.ColMachine_MachineCode, Me.ColMachine_Desc, Me.ColMachine_Remarks, Me.ColMachine_PrimaryMachine})
        Me.DgvMachine.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvMachine.Location = New System.Drawing.Point(2, 2)
        Me.DgvMachine.MultiSelect = False
        Me.DgvMachine.Name = "DgvMachine"
        Me.DgvMachine.RowTemplate.Height = 24
        Me.DgvMachine.Size = New System.Drawing.Size(628, 110)
        Me.DgvMachine.TabIndex = 2
        '
        'ColMachine_No
        '
        Me.ColMachine_No.Frozen = True
        Me.ColMachine_No.HeaderText = "No"
        Me.ColMachine_No.Name = "ColMachine_No"
        Me.ColMachine_No.ReadOnly = True
        Me.ColMachine_No.Width = 46
        '
        'ColMachine_MachineCode
        '
        Me.ColMachine_MachineCode.Frozen = True
        Me.ColMachine_MachineCode.HeaderText = "Machine Code"
        Me.ColMachine_MachineCode.MaxInputLength = 50
        Me.ColMachine_MachineCode.Name = "ColMachine_MachineCode"
        Me.ColMachine_MachineCode.Width = 101
        '
        'ColMachine_Desc
        '
        Me.ColMachine_Desc.Frozen = True
        Me.ColMachine_Desc.HeaderText = "Description"
        Me.ColMachine_Desc.MaxInputLength = 150
        Me.ColMachine_Desc.Name = "ColMachine_Desc"
        Me.ColMachine_Desc.ReadOnly = True
        Me.ColMachine_Desc.Width = 85
        '
        'ColMachine_Remarks
        '
        Me.ColMachine_Remarks.Frozen = True
        Me.ColMachine_Remarks.HeaderText = "Remarks"
        Me.ColMachine_Remarks.MaxInputLength = 150
        Me.ColMachine_Remarks.Name = "ColMachine_Remarks"
        Me.ColMachine_Remarks.Width = 74
        '
        'ColMachine_PrimaryMachine
        '
        Me.ColMachine_PrimaryMachine.Frozen = True
        Me.ColMachine_PrimaryMachine.HeaderText = "Primary Machine"
        Me.ColMachine_PrimaryMachine.Name = "ColMachine_PrimaryMachine"
        Me.ColMachine_PrimaryMachine.ReadOnly = True
        Me.ColMachine_PrimaryMachine.Visible = False
        Me.ColMachine_PrimaryMachine.Width = 110
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvEmployee)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(632, 114)
        Me.TabPage3.TabIndex = 1
        Me.TabPage3.Text = "Employee"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DgvEmployee
        '
        Me.DgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvEmployee.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColEmployee_No, Me.ColEmployee_ShiftCode, Me.ColEmployee_EmployeeCode, Me.ColEmployee_Name, Me.ColEmployee_Remarks, Me.ColEmployee_Primary, Me.ColEmployee_PrimaryShift})
        Me.DgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvEmployee.Location = New System.Drawing.Point(2, 2)
        Me.DgvEmployee.MultiSelect = False
        Me.DgvEmployee.Name = "DgvEmployee"
        Me.DgvEmployee.RowTemplate.Height = 24
        Me.DgvEmployee.Size = New System.Drawing.Size(628, 110)
        Me.DgvEmployee.TabIndex = 3
        '
        'ColEmployee_No
        '
        Me.ColEmployee_No.Frozen = True
        Me.ColEmployee_No.HeaderText = "No"
        Me.ColEmployee_No.Name = "ColEmployee_No"
        Me.ColEmployee_No.ReadOnly = True
        Me.ColEmployee_No.Width = 46
        '
        'ColEmployee_ShiftCode
        '
        Me.ColEmployee_ShiftCode.Frozen = True
        Me.ColEmployee_ShiftCode.HeaderText = "Shift"
        Me.ColEmployee_ShiftCode.MaxInputLength = 50
        Me.ColEmployee_ShiftCode.Name = "ColEmployee_ShiftCode"
        Me.ColEmployee_ShiftCode.Width = 53
        '
        'ColEmployee_EmployeeCode
        '
        Me.ColEmployee_EmployeeCode.Frozen = True
        Me.ColEmployee_EmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployee_EmployeeCode.MaxInputLength = 50
        Me.ColEmployee_EmployeeCode.Name = "ColEmployee_EmployeeCode"
        Me.ColEmployee_EmployeeCode.Width = 106
        '
        'ColEmployee_Name
        '
        Me.ColEmployee_Name.Frozen = True
        Me.ColEmployee_Name.HeaderText = "Name"
        Me.ColEmployee_Name.MaxInputLength = 150
        Me.ColEmployee_Name.Name = "ColEmployee_Name"
        Me.ColEmployee_Name.ReadOnly = True
        Me.ColEmployee_Name.Width = 60
        '
        'ColEmployee_Remarks
        '
        Me.ColEmployee_Remarks.Frozen = True
        Me.ColEmployee_Remarks.HeaderText = "Remarks"
        Me.ColEmployee_Remarks.MaxInputLength = 150
        Me.ColEmployee_Remarks.Name = "ColEmployee_Remarks"
        Me.ColEmployee_Remarks.Width = 74
        '
        'ColEmployee_Primary
        '
        Me.ColEmployee_Primary.Frozen = True
        Me.ColEmployee_Primary.HeaderText = "Primary Machine"
        Me.ColEmployee_Primary.Name = "ColEmployee_Primary"
        Me.ColEmployee_Primary.ReadOnly = True
        Me.ColEmployee_Primary.Visible = False
        Me.ColEmployee_Primary.Width = 110
        '
        'ColEmployee_PrimaryShift
        '
        Me.ColEmployee_PrimaryShift.Frozen = True
        Me.ColEmployee_PrimaryShift.HeaderText = "Primary Shift"
        Me.ColEmployee_PrimaryShift.Name = "ColEmployee_PrimaryShift"
        Me.ColEmployee_PrimaryShift.ReadOnly = True
        Me.ColEmployee_PrimaryShift.Visible = False
        Me.ColEmployee_PrimaryShift.Width = 90
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvSetuptime)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(632, 114)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "Setup time"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvSetuptime
        '
        Me.dgvSetuptime.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSetuptime.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSetuptime.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSetuptime.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSetupTime_No, Me.ColSetupTime_ProcessCode1, Me.ColSetupTime_CodeNo1, Me.ColSetupTime_Desc1, Me.ColSetupTime_ProcessCode2, Me.ColSetupTime_CodeNo2, Me.ColSetupTime_Desc2, Me.ColSetupTime_Code, Me.ColSetupTime_SetupTimeDuration, Me.ColSetupTime_Remarks, Me.ColSetupTime_PrimaryProcess1, Me.ColSetupTime_PrimaryCodeNo1, Me.ColSetupTime_PrimaryProses2, Me.ColSetupTime_PrimaryCodeNo2, Me.ColSetupTime_PrimarySetupTime})
        Me.dgvSetuptime.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSetuptime.Location = New System.Drawing.Point(2, 2)
        Me.dgvSetuptime.MultiSelect = False
        Me.dgvSetuptime.Name = "dgvSetuptime"
        Me.dgvSetuptime.RowTemplate.Height = 24
        Me.dgvSetuptime.Size = New System.Drawing.Size(628, 110)
        Me.dgvSetuptime.TabIndex = 4
        '
        'ColSetupTime_No
        '
        Me.ColSetupTime_No.Frozen = True
        Me.ColSetupTime_No.HeaderText = "No"
        Me.ColSetupTime_No.Name = "ColSetupTime_No"
        Me.ColSetupTime_No.ReadOnly = True
        Me.ColSetupTime_No.Width = 46
        '
        'ColSetupTime_ProcessCode1
        '
        Me.ColSetupTime_ProcessCode1.Frozen = True
        Me.ColSetupTime_ProcessCode1.HeaderText = "Process 1"
        Me.ColSetupTime_ProcessCode1.MaxInputLength = 50
        Me.ColSetupTime_ProcessCode1.Name = "ColSetupTime_ProcessCode1"
        Me.ColSetupTime_ProcessCode1.Width = 79
        '
        'ColSetupTime_CodeNo1
        '
        Me.ColSetupTime_CodeNo1.Frozen = True
        Me.ColSetupTime_CodeNo1.HeaderText = "Code No 1"
        Me.ColSetupTime_CodeNo1.MaxInputLength = 50
        Me.ColSetupTime_CodeNo1.Name = "ColSetupTime_CodeNo1"
        Me.ColSetupTime_CodeNo1.Width = 83
        '
        'ColSetupTime_Desc1
        '
        Me.ColSetupTime_Desc1.HeaderText = "Description 1"
        Me.ColSetupTime_Desc1.Name = "ColSetupTime_Desc1"
        Me.ColSetupTime_Desc1.Visible = False
        Me.ColSetupTime_Desc1.Width = 116
        '
        'ColSetupTime_ProcessCode2
        '
        Me.ColSetupTime_ProcessCode2.HeaderText = "Process 2"
        Me.ColSetupTime_ProcessCode2.MaxInputLength = 50
        Me.ColSetupTime_ProcessCode2.Name = "ColSetupTime_ProcessCode2"
        Me.ColSetupTime_ProcessCode2.ReadOnly = True
        Me.ColSetupTime_ProcessCode2.Width = 79
        '
        'ColSetupTime_CodeNo2
        '
        Me.ColSetupTime_CodeNo2.HeaderText = "Code No 2"
        Me.ColSetupTime_CodeNo2.MaxInputLength = 150
        Me.ColSetupTime_CodeNo2.Name = "ColSetupTime_CodeNo2"
        Me.ColSetupTime_CodeNo2.Width = 83
        '
        'ColSetupTime_Desc2
        '
        Me.ColSetupTime_Desc2.HeaderText = "Description 2"
        Me.ColSetupTime_Desc2.Name = "ColSetupTime_Desc2"
        Me.ColSetupTime_Desc2.Visible = False
        Me.ColSetupTime_Desc2.Width = 116
        '
        'ColSetupTime_Code
        '
        Me.ColSetupTime_Code.HeaderText = "Setup Time"
        Me.ColSetupTime_Code.Name = "ColSetupTime_Code"
        Me.ColSetupTime_Code.Width = 86
        '
        'ColSetupTime_SetupTimeDuration
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSetupTime_SetupTimeDuration.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSetupTime_SetupTimeDuration.HeaderText = "Duration"
        Me.ColSetupTime_SetupTimeDuration.Name = "ColSetupTime_SetupTimeDuration"
        Me.ColSetupTime_SetupTimeDuration.ReadOnly = True
        Me.ColSetupTime_SetupTimeDuration.Width = 72
        '
        'ColSetupTime_Remarks
        '
        Me.ColSetupTime_Remarks.HeaderText = "Remarks"
        Me.ColSetupTime_Remarks.Name = "ColSetupTime_Remarks"
        Me.ColSetupTime_Remarks.Width = 74
        '
        'ColSetupTime_PrimaryProcess1
        '
        Me.ColSetupTime_PrimaryProcess1.HeaderText = "Primary Process 1"
        Me.ColSetupTime_PrimaryProcess1.Name = "ColSetupTime_PrimaryProcess1"
        Me.ColSetupTime_PrimaryProcess1.ReadOnly = True
        Me.ColSetupTime_PrimaryProcess1.Visible = False
        Me.ColSetupTime_PrimaryProcess1.Width = 148
        '
        'ColSetupTime_PrimaryCodeNo1
        '
        Me.ColSetupTime_PrimaryCodeNo1.HeaderText = "Primary Code No 1"
        Me.ColSetupTime_PrimaryCodeNo1.Name = "ColSetupTime_PrimaryCodeNo1"
        Me.ColSetupTime_PrimaryCodeNo1.Visible = False
        Me.ColSetupTime_PrimaryCodeNo1.Width = 152
        '
        'ColSetupTime_PrimaryProses2
        '
        Me.ColSetupTime_PrimaryProses2.HeaderText = "Primary Proses 2"
        Me.ColSetupTime_PrimaryProses2.Name = "ColSetupTime_PrimaryProses2"
        Me.ColSetupTime_PrimaryProses2.Visible = False
        Me.ColSetupTime_PrimaryProses2.Width = 141
        '
        'ColSetupTime_PrimaryCodeNo2
        '
        Me.ColSetupTime_PrimaryCodeNo2.HeaderText = "Primary Code No 2"
        Me.ColSetupTime_PrimaryCodeNo2.Name = "ColSetupTime_PrimaryCodeNo2"
        Me.ColSetupTime_PrimaryCodeNo2.Visible = False
        Me.ColSetupTime_PrimaryCodeNo2.Width = 152
        '
        'ColSetupTime_PrimarySetupTime
        '
        Me.ColSetupTime_PrimarySetupTime.HeaderText = "Primary Setup Time"
        Me.ColSetupTime_PrimarySetupTime.Name = "ColSetupTime_PrimarySetupTime"
        Me.ColSetupTime_PrimarySetupTime.Visible = False
        Me.ColSetupTime_PrimarySetupTime.Width = 157
        '
        'ChkActive
        '
        Me.ChkActive.AutoSize = True
        Me.ChkActive.Location = New System.Drawing.Point(78, 194)
        Me.ChkActive.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkActive.Name = "ChkActive"
        Me.ChkActive.Size = New System.Drawing.Size(56, 17)
        Me.ChkActive.TabIndex = 15
        Me.ChkActive.Text = "Active"
        Me.ChkActive.UseVisualStyleBackColor = True
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(78, 69)
        Me.TxtRemarks.MaxLength = 100
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(559, 20)
        Me.TxtRemarks.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Remarks"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(343, 359)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 18
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'TxtUnitDesc
        '
        Me.TxtUnitDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnitDesc.Location = New System.Drawing.Point(78, 45)
        Me.TxtUnitDesc.MaxLength = 100
        Me.TxtUnitDesc.Name = "TxtUnitDesc"
        Me.TxtUnitDesc.Size = New System.Drawing.Size(559, 20)
        Me.TxtUnitDesc.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Description"
        '
        'TxtUnitCode
        '
        Me.TxtUnitCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnitCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnitCode.Location = New System.Drawing.Point(78, 19)
        Me.TxtUnitCode.MaxLength = 50
        Me.TxtUnitCode.Name = "TxtUnitCode"
        Me.TxtUnitCode.Size = New System.Drawing.Size(559, 20)
        Me.TxtUnitCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Code"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.BackColor = System.Drawing.SystemColors.Control
        Me.TxtReadOnly.Location = New System.Drawing.Point(78, 69)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.Size = New System.Drawing.Size(43, 20)
        Me.TxtReadOnly.TabIndex = 59
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtPrimaryWorkUnitGroup
        '
        Me.TxtPrimaryWorkUnitGroup.Location = New System.Drawing.Point(78, 93)
        Me.TxtPrimaryWorkUnitGroup.Name = "TxtPrimaryWorkUnitGroup"
        Me.TxtPrimaryWorkUnitGroup.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryWorkUnitGroup.TabIndex = 64
        Me.TxtPrimaryWorkUnitGroup.TabStop = False
        Me.TxtPrimaryWorkUnitGroup.Text = "-1"
        Me.TxtPrimaryWorkUnitGroup.Visible = False
        '
        'TxtPrimaryUnit
        '
        Me.TxtPrimaryUnit.Location = New System.Drawing.Point(78, 19)
        Me.TxtPrimaryUnit.Name = "TxtPrimaryUnit"
        Me.TxtPrimaryUnit.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryUnit.TabIndex = 32
        Me.TxtPrimaryUnit.TabStop = False
        Me.TxtPrimaryUnit.Visible = False
        '
        'txtPrimaryLocationIn
        '
        Me.txtPrimaryLocationIn.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrimaryLocationIn.Location = New System.Drawing.Point(78, 119)
        Me.txtPrimaryLocationIn.MaxLength = 100
        Me.txtPrimaryLocationIn.Name = "txtPrimaryLocationIn"
        Me.txtPrimaryLocationIn.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryLocationIn.TabIndex = 73
        Me.txtPrimaryLocationIn.Text = "-1"
        Me.txtPrimaryLocationIn.Visible = False
        '
        'txtPrimaryLocationOut
        '
        Me.txtPrimaryLocationOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrimaryLocationOut.Location = New System.Drawing.Point(78, 143)
        Me.txtPrimaryLocationOut.MaxLength = 100
        Me.txtPrimaryLocationOut.Name = "txtPrimaryLocationOut"
        Me.txtPrimaryLocationOut.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryLocationOut.TabIndex = 74
        Me.txtPrimaryLocationOut.Text = "-1"
        Me.txtPrimaryLocationOut.Visible = False
        '
        'TxtPrimaryProductionLine
        '
        Me.TxtPrimaryProductionLine.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryProductionLine.Location = New System.Drawing.Point(78, 169)
        Me.TxtPrimaryProductionLine.MaxLength = 100
        Me.TxtPrimaryProductionLine.Name = "TxtPrimaryProductionLine"
        Me.TxtPrimaryProductionLine.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryProductionLine.TabIndex = 80
        Me.TxtPrimaryProductionLine.Text = "-1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(332, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 434)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(679, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 37
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(332, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 456)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(679, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 38
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(332, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(471, 398)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(385, 398)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(213, 398)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
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
        'FrmMasterWorkUnit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 478)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(695, 516)
        Me.Name = "FrmMasterWorkUnit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Unit Data"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvMachine, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvSetuptime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents ChkActive As System.Windows.Forms.CheckBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryUnit As System.Windows.Forms.TextBox
    Friend WithEvents TxtUnitDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtUnitCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvMachine As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColMachine_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMachine_MachineCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMachine_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMachine_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMachine_PrimaryMachine As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_ShiftCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_EmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_PrimaryShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnFindGroup As System.Windows.Forms.Button
    Friend WithEvents txtGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtGroupDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryWorkUnitGroup As System.Windows.Forms.TextBox
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvSetuptime As System.Windows.Forms.DataGridView
    Friend WithEvents ColSetupTime_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_ProcessCode1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_CodeNo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Desc1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_ProcessCode2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_CodeNo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Desc2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_SetupTimeDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryProcess1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryCodeNo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryProses2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryCodeNo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimarySetupTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BtnFindLocationOut As System.Windows.Forms.Button
    Friend WithEvents BtnFindLocationIn As System.Windows.Forms.Button
    Friend WithEvents txtDescLocationOut As System.Windows.Forms.TextBox
    Friend WithEvents txtCodeLocationOut As System.Windows.Forms.TextBox
    Friend WithEvents txtDescLocationIn As System.Windows.Forms.TextBox
    Friend WithEvents txtCodeLocationIn As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryLocationOut As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryLocationIn As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryProductionLine As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BtnFindProductionLine As System.Windows.Forms.Button
    Friend WithEvents TxtProductionLineDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtProductionLineCode As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
End Class
