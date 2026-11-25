<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKaryawan_Kompetensi_Upd
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.cboDefValue = New System.Windows.Forms.ComboBox
        Me.DtpDefValue = New System.Windows.Forms.DateTimePicker
        Me.TxtDefValueText = New System.Windows.Forms.TextBox
        Me.cboAlfabetic = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboDataType = New System.Windows.Forms.ComboBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.BtnShow = New System.Windows.Forms.Button
        Me.BtnFindCompetace = New System.Windows.Forms.Button
        Me.ChkShowAll = New System.Windows.Forms.CheckBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtDefValueDigit = New System.Windows.Forms.TextBox
        Me.TxtLastPrimary = New System.Windows.Forms.TextBox
        Me.Panel2 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvDepartment = New System.Windows.Forms.DataGridView
        Me.ColDeptNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvJobPosition = New System.Windows.Forms.DataGridView
        Me.ColJobPosisitonNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPosisitonCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPosisitonDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPosisitonPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColEmployee_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_NIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.DgvSalFormula = New System.Windows.Forms.DataGridView
        Me.ColSalFor_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalFor_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalFor_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalFor_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.LlbUnchkAllExp = New System.Windows.Forms.LinkLabel
        Me.LlbChkAllExp = New System.Windows.Forms.LinkLabel
        Me.LlbUnchkAll = New System.Windows.Forms.LinkLabel
        Me.LlbChkAll = New System.Windows.Forms.LinkLabel
        Me.LlbAdvance = New System.Windows.Forms.LinkLabel
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_EmpCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_EmpNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_EmpName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_JobPos = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Dept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Value = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_NextCertDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Modify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_Del = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_Export = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_PrimaryEmp = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_PrimaryJobPost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteToAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.ImportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BgwJobPosition = New System.ComponentModel.BackgroundWorker
        Me.BgwDepartment = New System.ComponentModel.BackgroundWorker
        Me.BgwEmployee = New System.ComponentModel.BackgroundWorker
        Me.BgwSalFor = New System.ComponentModel.BackgroundWorker
        Me.frmDia = New System.Windows.Forms.SaveFileDialog
        Me.OpenDia = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DgvSalFormula, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(836, 162)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboDefValue)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDefValue)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDefValueText)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboAlfabetic)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboDataType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnShow)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindCompetace)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDefValueDigit)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtLastPrimary)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Size = New System.Drawing.Size(824, 144)
        Me.SplitContainer1.SplitterDistance = 330
        Me.SplitContainer1.TabIndex = 0
        '
        'cboDefValue
        '
        Me.cboDefValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDefValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDefValue.FormattingEnabled = True
        Me.cboDefValue.Items.AddRange(New Object() {"No", "Yes"})
        Me.cboDefValue.Location = New System.Drawing.Point(94, 81)
        Me.cboDefValue.Name = "cboDefValue"
        Me.cboDefValue.Size = New System.Drawing.Size(233, 21)
        Me.cboDefValue.TabIndex = 4
        '
        'DtpDefValue
        '
        Me.DtpDefValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDefValue.Location = New System.Drawing.Point(94, 81)
        Me.DtpDefValue.Name = "DtpDefValue"
        Me.DtpDefValue.Size = New System.Drawing.Size(233, 20)
        Me.DtpDefValue.TabIndex = 4
        '
        'TxtDefValueText
        '
        Me.TxtDefValueText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDefValueText.Location = New System.Drawing.Point(94, 81)
        Me.TxtDefValueText.MaxLength = 250
        Me.TxtDefValueText.Name = "TxtDefValueText"
        Me.TxtDefValueText.Size = New System.Drawing.Size(233, 20)
        Me.TxtDefValueText.TabIndex = 4
        '
        'cboAlfabetic
        '
        Me.cboAlfabetic.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAlfabetic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAlfabetic.FormattingEnabled = True
        Me.cboAlfabetic.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F"})
        Me.cboAlfabetic.Location = New System.Drawing.Point(94, 81)
        Me.cboAlfabetic.Name = "cboAlfabetic"
        Me.cboAlfabetic.Size = New System.Drawing.Size(233, 21)
        Me.cboAlfabetic.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Type Value"
        '
        'cboDataType
        '
        Me.cboDataType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDataType.Enabled = False
        Me.cboDataType.FormattingEnabled = True
        Me.cboDataType.Items.AddRange(New Object() {"Number", "Date", "Text", "True/False", "Alphabetic"})
        Me.cboDataType.Location = New System.Drawing.Point(94, 55)
        Me.cboDataType.Name = "cboDataType"
        Me.cboDataType.Size = New System.Drawing.Size(233, 21)
        Me.cboDataType.TabIndex = 3
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(42, 3)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimary.TabIndex = 10
        Me.TxtPrimary.Text = "-1"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Standar Value"
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(94, 3)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(186, 20)
        Me.TxtCode.TabIndex = 0
        '
        'BtnShow
        '
        Me.BtnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnShow.Location = New System.Drawing.Point(252, 111)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(75, 23)
        Me.BtnShow.TabIndex = 6
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'BtnFindCompetace
        '
        Me.BtnFindCompetace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCompetace.Location = New System.Drawing.Point(287, 3)
        Me.BtnFindCompetace.Name = "BtnFindCompetace"
        Me.BtnFindCompetace.Size = New System.Drawing.Size(41, 20)
        Me.BtnFindCompetace.TabIndex = 1
        Me.BtnFindCompetace.TabStop = False
        Me.BtnFindCompetace.Text = "Find"
        Me.BtnFindCompetace.UseVisualStyleBackColor = True
        '
        'ChkShowAll
        '
        Me.ChkShowAll.AutoSize = True
        Me.ChkShowAll.Location = New System.Drawing.Point(94, 115)
        Me.ChkShowAll.Name = "ChkShowAll"
        Me.ChkShowAll.Size = New System.Drawing.Size(116, 17)
        Me.ChkShowAll.TabIndex = 5
        Me.ChkShowAll.Text = "Show All Employee"
        Me.ChkShowAll.UseVisualStyleBackColor = True
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(94, 29)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(233, 20)
        Me.TxtDesc.TabIndex = 2
        Me.TxtDesc.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Description"
        '
        'TxtDefValueDigit
        '
        Me.TxtDefValueDigit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDefValueDigit.Location = New System.Drawing.Point(94, 81)
        Me.TxtDefValueDigit.MaxLength = 15
        Me.TxtDefValueDigit.Name = "TxtDefValueDigit"
        Me.TxtDefValueDigit.Size = New System.Drawing.Size(233, 20)
        Me.TxtDefValueDigit.TabIndex = 4
        Me.TxtDefValueDigit.Text = "0"
        Me.TxtDefValueDigit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtLastPrimary
        '
        Me.TxtLastPrimary.Location = New System.Drawing.Point(42, 24)
        Me.TxtLastPrimary.Name = "TxtLastPrimary"
        Me.TxtLastPrimary.Size = New System.Drawing.Size(33, 20)
        Me.TxtLastPrimary.TabIndex = 13
        Me.TxtLastPrimary.Text = "-1"
        Me.TxtLastPrimary.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.TabPage1)
        Me.Panel2.Controls.Add(Me.TabPage2)
        Me.Panel2.Controls.Add(Me.TabPage3)
        Me.Panel2.Controls.Add(Me.TabPage4)
        Me.Panel2.Location = New System.Drawing.Point(2, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.SelectedIndex = 0
        Me.Panel2.Size = New System.Drawing.Size(485, 138)
        Me.Panel2.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvDepartment)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(477, 112)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Department"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvDepartment
        '
        Me.dgvDepartment.AllowUserToResizeColumns = False
        Me.dgvDepartment.AllowUserToResizeRows = False
        Me.dgvDepartment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDepartment.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvDepartment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDepartment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDeptNo, Me.ColDeptCode, Me.ColDeptDesc, Me.ColDeptPrimary})
        Me.dgvDepartment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDepartment.Location = New System.Drawing.Point(3, 3)
        Me.dgvDepartment.Name = "dgvDepartment"
        Me.dgvDepartment.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDepartment.RowTemplate.Height = 24
        Me.dgvDepartment.Size = New System.Drawing.Size(471, 106)
        Me.dgvDepartment.TabIndex = 0
        '
        'ColDeptNo
        '
        Me.ColDeptNo.HeaderText = "No"
        Me.ColDeptNo.Name = "ColDeptNo"
        Me.ColDeptNo.ReadOnly = True
        Me.ColDeptNo.Width = 46
        '
        'ColDeptCode
        '
        Me.ColDeptCode.HeaderText = "Deptartment Code"
        Me.ColDeptCode.Name = "ColDeptCode"
        Me.ColDeptCode.Width = 118
        '
        'ColDeptDesc
        '
        Me.ColDeptDesc.HeaderText = "Description"
        Me.ColDeptDesc.Name = "ColDeptDesc"
        Me.ColDeptDesc.ReadOnly = True
        Me.ColDeptDesc.Width = 85
        '
        'ColDeptPrimary
        '
        Me.ColDeptPrimary.HeaderText = "Primary Department"
        Me.ColDeptPrimary.Name = "ColDeptPrimary"
        Me.ColDeptPrimary.ReadOnly = True
        Me.ColDeptPrimary.Visible = False
        Me.ColDeptPrimary.Width = 159
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvJobPosition)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(480, 112)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Job Position"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvJobPosition
        '
        Me.dgvJobPosition.AllowUserToResizeColumns = False
        Me.dgvJobPosition.AllowUserToResizeRows = False
        Me.dgvJobPosition.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvJobPosition.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvJobPosition.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvJobPosition.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColJobPosisitonNo, Me.ColJobPosisitonCode, Me.ColJobPosisitonDesc, Me.ColJobPosisitonPrimary})
        Me.dgvJobPosition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvJobPosition.Location = New System.Drawing.Point(3, 3)
        Me.dgvJobPosition.Name = "dgvJobPosition"
        Me.dgvJobPosition.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvJobPosition.RowTemplate.Height = 24
        Me.dgvJobPosition.Size = New System.Drawing.Size(474, 106)
        Me.dgvJobPosition.TabIndex = 1
        '
        'ColJobPosisitonNo
        '
        Me.ColJobPosisitonNo.HeaderText = "No"
        Me.ColJobPosisitonNo.Name = "ColJobPosisitonNo"
        Me.ColJobPosisitonNo.ReadOnly = True
        Me.ColJobPosisitonNo.Width = 46
        '
        'ColJobPosisitonCode
        '
        Me.ColJobPosisitonCode.HeaderText = "Job Position Code"
        Me.ColJobPosisitonCode.Name = "ColJobPosisitonCode"
        Me.ColJobPosisitonCode.Width = 117
        '
        'ColJobPosisitonDesc
        '
        Me.ColJobPosisitonDesc.HeaderText = "Description"
        Me.ColJobPosisitonDesc.Name = "ColJobPosisitonDesc"
        Me.ColJobPosisitonDesc.ReadOnly = True
        Me.ColJobPosisitonDesc.Width = 85
        '
        'ColJobPosisitonPrimary
        '
        Me.ColJobPosisitonPrimary.HeaderText = "Primary Job Position"
        Me.ColJobPosisitonPrimary.Name = "ColJobPosisitonPrimary"
        Me.ColJobPosisitonPrimary.ReadOnly = True
        Me.ColJobPosisitonPrimary.Visible = False
        Me.ColJobPosisitonPrimary.Width = 162
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvEmployee)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage3.Size = New System.Drawing.Size(480, 112)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Employee"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AllowUserToResizeColumns = False
        Me.dgvEmployee.AllowUserToResizeRows = False
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColEmployee_No, Me.ColEmployee_Code, Me.ColEmployee_Name, Me.ColEmployee_NIK, Me.ColEmployee_primary})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(3, 3)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(474, 106)
        Me.dgvEmployee.TabIndex = 3
        '
        'ColEmployee_No
        '
        Me.ColEmployee_No.HeaderText = "No"
        Me.ColEmployee_No.Name = "ColEmployee_No"
        Me.ColEmployee_No.ReadOnly = True
        Me.ColEmployee_No.Width = 46
        '
        'ColEmployee_Code
        '
        Me.ColEmployee_Code.HeaderText = "Employee Code"
        Me.ColEmployee_Code.Name = "ColEmployee_Code"
        Me.ColEmployee_Code.Width = 106
        '
        'ColEmployee_Name
        '
        Me.ColEmployee_Name.HeaderText = "Employee Name"
        Me.ColEmployee_Name.Name = "ColEmployee_Name"
        Me.ColEmployee_Name.ReadOnly = True
        Me.ColEmployee_Name.Width = 109
        '
        'ColEmployee_NIK
        '
        Me.ColEmployee_NIK.HeaderText = "Employee NIK"
        Me.ColEmployee_NIK.Name = "ColEmployee_NIK"
        Me.ColEmployee_NIK.ReadOnly = True
        Me.ColEmployee_NIK.Width = 99
        '
        'ColEmployee_primary
        '
        Me.ColEmployee_primary.HeaderText = "Primary Employee"
        Me.ColEmployee_primary.Name = "ColEmployee_primary"
        Me.ColEmployee_primary.ReadOnly = True
        Me.ColEmployee_primary.Visible = False
        Me.ColEmployee_primary.Width = 147
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DgvSalFormula)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage4.Size = New System.Drawing.Size(480, 112)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Salary Formula"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DgvSalFormula
        '
        Me.DgvSalFormula.AllowUserToResizeColumns = False
        Me.DgvSalFormula.AllowUserToResizeRows = False
        Me.DgvSalFormula.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalFormula.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvSalFormula.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSalFormula.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSalFor_No, Me.ColSalFor_Code, Me.ColSalFor_Desc, Me.ColSalFor_Primary})
        Me.DgvSalFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalFormula.Location = New System.Drawing.Point(3, 3)
        Me.DgvSalFormula.Name = "DgvSalFormula"
        Me.DgvSalFormula.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSalFormula.RowTemplate.Height = 24
        Me.DgvSalFormula.Size = New System.Drawing.Size(474, 106)
        Me.DgvSalFormula.TabIndex = 2
        '
        'ColSalFor_No
        '
        Me.ColSalFor_No.HeaderText = "No"
        Me.ColSalFor_No.Name = "ColSalFor_No"
        Me.ColSalFor_No.ReadOnly = True
        Me.ColSalFor_No.Width = 46
        '
        'ColSalFor_Code
        '
        Me.ColSalFor_Code.HeaderText = "Salary Formula Code"
        Me.ColSalFor_Code.Name = "ColSalFor_Code"
        Me.ColSalFor_Code.Width = 129
        '
        'ColSalFor_Desc
        '
        Me.ColSalFor_Desc.HeaderText = "Description"
        Me.ColSalFor_Desc.Name = "ColSalFor_Desc"
        Me.ColSalFor_Desc.ReadOnly = True
        Me.ColSalFor_Desc.Width = 85
        '
        'ColSalFor_Primary
        '
        Me.ColSalFor_Primary.HeaderText = "Primary SF"
        Me.ColSalFor_Primary.Name = "ColSalFor_Primary"
        Me.ColSalFor_Primary.ReadOnly = True
        Me.ColSalFor_Primary.Visible = False
        Me.ColSalFor_Primary.Width = 102
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.LlbUnchkAllExp)
        Me.GroupBox2.Controls.Add(Me.LlbChkAllExp)
        Me.GroupBox2.Controls.Add(Me.LlbUnchkAll)
        Me.GroupBox2.Controls.Add(Me.LlbChkAll)
        Me.GroupBox2.Controls.Add(Me.LlbAdvance)
        Me.GroupBox2.Controls.Add(Me.DgvItem)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 163)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(836, 322)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'LlbUnchkAllExp
        '
        Me.LlbUnchkAllExp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUnchkAllExp.AutoSize = True
        Me.LlbUnchkAllExp.Location = New System.Drawing.Point(679, 0)
        Me.LlbUnchkAllExp.Name = "LlbUnchkAllExp"
        Me.LlbUnchkAllExp.Size = New System.Drawing.Size(98, 13)
        Me.LlbUnchkAllExp.TabIndex = 5
        Me.LlbUnchkAllExp.TabStop = True
        Me.LlbUnchkAllExp.Text = "Uncheck All Export"
        '
        'LlbChkAllExp
        '
        Me.LlbChkAllExp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbChkAllExp.AutoSize = True
        Me.LlbChkAllExp.Location = New System.Drawing.Point(590, 0)
        Me.LlbChkAllExp.Name = "LlbChkAllExp"
        Me.LlbChkAllExp.Size = New System.Drawing.Size(85, 13)
        Me.LlbChkAllExp.TabIndex = 4
        Me.LlbChkAllExp.TabStop = True
        Me.LlbChkAllExp.Text = "Check All Export"
        '
        'LlbUnchkAll
        '
        Me.LlbUnchkAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUnchkAll.AutoSize = True
        Me.LlbUnchkAll.Location = New System.Drawing.Point(489, 0)
        Me.LlbUnchkAll.Name = "LlbUnchkAll"
        Me.LlbUnchkAll.Size = New System.Drawing.Size(99, 13)
        Me.LlbUnchkAll.TabIndex = 1
        Me.LlbUnchkAll.TabStop = True
        Me.LlbUnchkAll.Text = "Uncheck All Modify"
        '
        'LlbChkAll
        '
        Me.LlbChkAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbChkAll.AutoSize = True
        Me.LlbChkAll.Location = New System.Drawing.Point(400, 0)
        Me.LlbChkAll.Name = "LlbChkAll"
        Me.LlbChkAll.Size = New System.Drawing.Size(86, 13)
        Me.LlbChkAll.TabIndex = 0
        Me.LlbChkAll.TabStop = True
        Me.LlbChkAll.Text = "Check All Modify"
        '
        'LlbAdvance
        '
        Me.LlbAdvance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbAdvance.AutoSize = True
        Me.LlbAdvance.Location = New System.Drawing.Point(779, 0)
        Me.LlbAdvance.Name = "LlbAdvance"
        Me.LlbAdvance.Size = New System.Drawing.Size(50, 13)
        Me.LlbAdvance.TabIndex = 2
        Me.LlbAdvance.TabStop = True
        Me.LlbAdvance.Text = "Advance"
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_EmpCode, Me.ColItem_EmpNIK, Me.ColItem_EmpName, Me.ColItem_JobPos, Me.ColItem_Dept, Me.ColItem_Value, Me.ColItem_NextCertDate, Me.ColItem_Remarks, Me.ColItem_Modify, Me.ColItem_Del, Me.ColItem_Export, Me.ColItem_PrimaryEmp, Me.ColItem_PrimaryJobPost})
        Me.DgvItem.Location = New System.Drawing.Point(3, 16)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(830, 303)
        Me.DgvItem.TabIndex = 3
        '
        'ColItem_No
        '
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.ReadOnly = True
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 51
        '
        'ColItem_EmpCode
        '
        Me.ColItem_EmpCode.HeaderText = "Employee Code"
        Me.ColItem_EmpCode.Name = "ColItem_EmpCode"
        Me.ColItem_EmpCode.ReadOnly = True
        Me.ColItem_EmpCode.Width = 132
        '
        'ColItem_EmpNIK
        '
        Me.ColItem_EmpNIK.HeaderText = "Employee NIK"
        Me.ColItem_EmpNIK.Name = "ColItem_EmpNIK"
        Me.ColItem_EmpNIK.ReadOnly = True
        Me.ColItem_EmpNIK.Width = 121
        '
        'ColItem_EmpName
        '
        Me.ColItem_EmpName.HeaderText = "Employee Name"
        Me.ColItem_EmpName.Name = "ColItem_EmpName"
        Me.ColItem_EmpName.ReadOnly = True
        Me.ColItem_EmpName.Width = 136
        '
        'ColItem_JobPos
        '
        Me.ColItem_JobPos.HeaderText = "Job Position"
        Me.ColItem_JobPos.Name = "ColItem_JobPos"
        Me.ColItem_JobPos.ReadOnly = True
        Me.ColItem_JobPos.Width = 110
        '
        'ColItem_Dept
        '
        Me.ColItem_Dept.HeaderText = "Department"
        Me.ColItem_Dept.Name = "ColItem_Dept"
        Me.ColItem_Dept.ReadOnly = True
        Me.ColItem_Dept.Width = 107
        '
        'ColItem_Value
        '
        Me.ColItem_Value.HeaderText = "Value"
        Me.ColItem_Value.Name = "ColItem_Value"
        Me.ColItem_Value.Width = 69
        '
        'ColItem_NextCertDate
        '
        Me.ColItem_NextCertDate.HeaderText = "Next Certificate Date"
        Me.ColItem_NextCertDate.Name = "ColItem_NextCertDate"
        Me.ColItem_NextCertDate.Width = 162
        '
        'ColItem_Remarks
        '
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        Me.ColItem_Remarks.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItem_Remarks.HeaderText = "Remarks"
        Me.ColItem_Remarks.Name = "ColItem_Remarks"
        Me.ColItem_Remarks.Width = 89
        '
        'ColItem_Modify
        '
        Me.ColItem_Modify.HeaderText = "Modify"
        Me.ColItem_Modify.Name = "ColItem_Modify"
        Me.ColItem_Modify.Width = 55
        '
        'ColItem_Del
        '
        Me.ColItem_Del.HeaderText = "Delete"
        Me.ColItem_Del.Name = "ColItem_Del"
        Me.ColItem_Del.Width = 55
        '
        'ColItem_Export
        '
        Me.ColItem_Export.HeaderText = "Export"
        Me.ColItem_Export.Name = "ColItem_Export"
        Me.ColItem_Export.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'ColItem_PrimaryEmp
        '
        Me.ColItem_PrimaryEmp.HeaderText = "Primary Employee"
        Me.ColItem_PrimaryEmp.Name = "ColItem_PrimaryEmp"
        Me.ColItem_PrimaryEmp.Visible = False
        Me.ColItem_PrimaryEmp.Width = 96
        '
        'ColItem_PrimaryJobPost
        '
        Me.ColItem_PrimaryJobPost.HeaderText = "Primary Job"
        Me.ColItem_PrimaryJobPost.Name = "ColItem_PrimaryJobPost"
        Me.ColItem_PrimaryJobPost.ReadOnly = True
        Me.ColItem_PrimaryJobPost.Visible = False
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllToolStripMenuItem, Me.DeleteToAllToolStripMenuItem, Me.ToolStripSeparator1, Me.ImportToolStripMenuItem, Me.ExportToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(142, 98)
        '
        'ApplyToAllToolStripMenuItem
        '
        Me.ApplyToAllToolStripMenuItem.Name = "ApplyToAllToolStripMenuItem"
        Me.ApplyToAllToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.ApplyToAllToolStripMenuItem.Text = "Apply To All"
        '
        'DeleteToAllToolStripMenuItem
        '
        Me.DeleteToAllToolStripMenuItem.Name = "DeleteToAllToolStripMenuItem"
        Me.DeleteToAllToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.DeleteToAllToolStripMenuItem.Text = "Delete To All"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(138, 6)
        '
        'ImportToolStripMenuItem
        '
        Me.ImportToolStripMenuItem.Name = "ImportToolStripMenuItem"
        Me.ImportToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.ImportToolStripMenuItem.Text = "Import File"
        '
        'ExportToolStripMenuItem
        '
        Me.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem"
        Me.ExportToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.ExportToolStripMenuItem.Text = "Export File"
        '
        'BgwJobPosition
        '
        '
        'BgwDepartment
        '
        '
        'BgwEmployee
        '
        '
        'BgwSalFor
        '
        '
        'OpenDia
        '
        Me.OpenDia.FileName = "OpenFileDialog1"
        '
        'FrmMasterKaryawan_Kompetensi_Upd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(845, 494)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimumSize = New System.Drawing.Size(861, 532)
        Me.Name = "FrmMasterKaryawan_Kompetensi_Upd"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Competence Employee Update"
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DgvSalFormula, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvDepartment As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvJobPosition As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindCompetace As System.Windows.Forms.Button
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents ChkShowAll As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LlbAdvance As System.Windows.Forms.LinkLabel
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LlbUnchkAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbChkAll As System.Windows.Forms.LinkLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyToAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteToAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents DgvSalFormula As System.Windows.Forms.DataGridView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BgwJobPosition As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwDepartment As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwEmployee As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColDeptNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosisitonNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosisitonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosisitonDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosisitonPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_NIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalFor_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalFor_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalFor_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalFor_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwSalFor As System.ComponentModel.BackgroundWorker
    Friend WithEvents cboDataType As System.Windows.Forms.ComboBox
    Friend WithEvents cboDefValue As System.Windows.Forms.ComboBox
    Friend WithEvents DtpDefValue As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtDefValueText As System.Windows.Forms.TextBox
    Friend WithEvents cboAlfabetic As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDefValueDigit As System.Windows.Forms.TextBox
    Friend WithEvents frmDia As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenDia As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TxtLastPrimary As System.Windows.Forms.TextBox
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_EmpCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_EmpNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_EmpName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_JobPos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Dept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_NextCertDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Modify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_Del As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_Export As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_PrimaryEmp As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_PrimaryJobPost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LlbUnchkAllExp As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbChkAllExp As System.Windows.Forms.LinkLabel
End Class
