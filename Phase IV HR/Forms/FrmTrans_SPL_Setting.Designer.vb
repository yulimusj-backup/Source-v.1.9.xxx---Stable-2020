<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL_Setting
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
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColUserNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserEmail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserAllowOverBudgetSPL = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColUserApproved = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserApprovedPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserPrimaryUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.chkApplyToAllEmployee = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.txtReadonly = New System.Windows.Forms.TextBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvScope = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartmentCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartmentDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllWokUnitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyToWorkUnitGroupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyByJobPosirionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CustomEmployeeListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker_CloseForm = New System.ComponentModel.BackgroundWorker
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvScope, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(77, 2)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(269, 20)
        Me.txtRemarks.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvUser)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(696, 324)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "User"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToAddRows = False
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvUser.ColumnHeadersHeight = 25
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColUserNo, Me.ColUserCode, Me.ColUserName, Me.ColUserEmail, Me.ColUserAllowOverBudgetSPL, Me.ColUserApproved, Me.ColUserApprovedPosition, Me.ColUserRemarks, Me.ColUserPrimaryUser})
        Me.DgvUser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvUser.Location = New System.Drawing.Point(2, 2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(692, 320)
        Me.DgvUser.TabIndex = 1
        '
        'ColUserNo
        '
        Me.ColUserNo.Frozen = True
        Me.ColUserNo.HeaderText = "No"
        Me.ColUserNo.Name = "ColUserNo"
        Me.ColUserNo.ReadOnly = True
        Me.ColUserNo.Width = 46
        '
        'ColUserCode
        '
        Me.ColUserCode.HeaderText = "User Code"
        Me.ColUserCode.Name = "ColUserCode"
        Me.ColUserCode.Width = 82
        '
        'ColUserName
        '
        Me.ColUserName.HeaderText = "Name"
        Me.ColUserName.Name = "ColUserName"
        Me.ColUserName.Width = 60
        '
        'ColUserEmail
        '
        Me.ColUserEmail.HeaderText = "Email Address"
        Me.ColUserEmail.Name = "ColUserEmail"
        Me.ColUserEmail.Width = 98
        '
        'ColUserAllowOverBudgetSPL
        '
        Me.ColUserAllowOverBudgetSPL.HeaderText = "Approve Over Budget Overtime"
        Me.ColUserAllowOverBudgetSPL.Name = "ColUserAllowOverBudgetSPL"
        Me.ColUserAllowOverBudgetSPL.Width = 161
        '
        'ColUserApproved
        '
        Me.ColUserApproved.HeaderText = "Approved By"
        Me.ColUserApproved.Name = "ColUserApproved"
        Me.ColUserApproved.Width = 93
        '
        'ColUserApprovedPosition
        '
        Me.ColUserApprovedPosition.HeaderText = "Approved By Position"
        Me.ColUserApprovedPosition.Name = "ColUserApprovedPosition"
        Me.ColUserApprovedPosition.Width = 133
        '
        'ColUserRemarks
        '
        Me.ColUserRemarks.HeaderText = "Remarks"
        Me.ColUserRemarks.Name = "ColUserRemarks"
        Me.ColUserRemarks.Width = 74
        '
        'ColUserPrimaryUser
        '
        Me.ColUserPrimaryUser.HeaderText = "Primary User"
        Me.ColUserPrimaryUser.Name = "ColUserPrimaryUser"
        Me.ColUserPrimaryUser.ReadOnly = True
        Me.ColUserPrimaryUser.Visible = False
        Me.ColUserPrimaryUser.Width = 91
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code"
        '
        'txtCode
        '
        Me.txtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCode.BackColor = System.Drawing.Color.White
        Me.txtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCode.Location = New System.Drawing.Point(80, 2)
        Me.txtCode.MaxLength = 250
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(267, 20)
        Me.txtCode.TabIndex = 1
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 10)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkApplyToAllEmployee)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Size = New System.Drawing.Size(702, 48)
        Me.SplitContainer1.SplitterDistance = 349
        Me.SplitContainer1.TabIndex = 99999999
        Me.SplitContainer1.TabStop = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.BackColor = System.Drawing.SystemColors.Window
        Me.TxtDesc.Location = New System.Drawing.Point(80, 26)
        Me.TxtDesc.MaxLength = 250
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(267, 20)
        Me.TxtDesc.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 29)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Description"
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(10, 0)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimary.TabIndex = 0
        Me.txtPrimary.TabStop = False
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'chkApplyToAllEmployee
        '
        Me.chkApplyToAllEmployee.AutoSize = True
        Me.chkApplyToAllEmployee.Location = New System.Drawing.Point(78, 27)
        Me.chkApplyToAllEmployee.Margin = New System.Windows.Forms.Padding(2)
        Me.chkApplyToAllEmployee.Name = "chkApplyToAllEmployee"
        Me.chkApplyToAllEmployee.Size = New System.Drawing.Size(86, 17)
        Me.chkApplyToAllEmployee.TabIndex = 1
        Me.chkApplyToAllEmployee.Text = "All Employee"
        Me.chkApplyToAllEmployee.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Remarks"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(180, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'txtReadonly
        '
        Me.txtReadonly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReadonly.Location = New System.Drawing.Point(693, 466)
        Me.txtReadonly.Name = "txtReadonly"
        Me.txtReadonly.ReadOnly = True
        Me.txtReadonly.Size = New System.Drawing.Size(29, 20)
        Me.txtReadonly.TabIndex = 38
        Me.txtReadonly.TabStop = False
        Me.txtReadonly.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 492)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(738, 22)
        Me.StatusStrip1.TabIndex = 40
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(180, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(180, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(180, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(5, 63)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(704, 350)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvScope)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(696, 324)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Scope"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvScope
        '
        Me.dgvScope.AllowUserToAddRows = False
        Me.dgvScope.AllowUserToResizeColumns = False
        Me.dgvScope.AllowUserToResizeRows = False
        Me.dgvScope.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvScope.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvScope.ColumnHeadersHeight = 25
        Me.dgvScope.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvScope.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.ColDepartmentCode, Me.ColDepartmentDesc, Me.ColJobPositionCode, Me.ColJobPositionDesc, Me.ColEmployeeCode, Me.ColEmployeeName, Me.colRemarks, Me.ColPrimaryDept, Me.ColPrimaryPosition, Me.ColPrimaryEmp})
        Me.dgvScope.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvScope.Location = New System.Drawing.Point(2, 2)
        Me.dgvScope.Name = "dgvScope"
        Me.dgvScope.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvScope.RowTemplate.Height = 24
        Me.dgvScope.Size = New System.Drawing.Size(692, 320)
        Me.dgvScope.TabIndex = 0
        '
        'colNo
        '
        Me.colNo.Frozen = True
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 46
        '
        'ColDepartmentCode
        '
        Me.ColDepartmentCode.HeaderText = "Departement Code"
        Me.ColDepartmentCode.Name = "ColDepartmentCode"
        Me.ColDepartmentCode.Width = 121
        '
        'ColDepartmentDesc
        '
        Me.ColDepartmentDesc.HeaderText = "Description"
        Me.ColDepartmentDesc.Name = "ColDepartmentDesc"
        Me.ColDepartmentDesc.Width = 85
        '
        'ColJobPositionCode
        '
        Me.ColJobPositionCode.HeaderText = "Job Position Code"
        Me.ColJobPositionCode.Name = "ColJobPositionCode"
        Me.ColJobPositionCode.Width = 117
        '
        'ColJobPositionDesc
        '
        Me.ColJobPositionDesc.HeaderText = "Description"
        Me.ColJobPositionDesc.Name = "ColJobPositionDesc"
        Me.ColJobPositionDesc.Width = 85
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
        Me.ColEmployeeName.Visible = False
        Me.ColEmployeeName.Width = 109
        '
        'colRemarks
        '
        Me.colRemarks.HeaderText = "Remarks"
        Me.colRemarks.Name = "colRemarks"
        Me.colRemarks.Width = 74
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.ReadOnly = True
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 92
        '
        'ColPrimaryPosition
        '
        Me.ColPrimaryPosition.HeaderText = "Primary Position"
        Me.ColPrimaryPosition.Name = "ColPrimaryPosition"
        Me.ColPrimaryPosition.ReadOnly = True
        Me.ColPrimaryPosition.Visible = False
        Me.ColPrimaryPosition.Width = 106
        '
        'ColPrimaryEmp
        '
        Me.ColPrimaryEmp.HeaderText = "Primary Employee"
        Me.ColPrimaryEmp.Name = "ColPrimaryEmp"
        Me.ColPrimaryEmp.ReadOnly = True
        Me.ColPrimaryEmp.Visible = False
        Me.ColPrimaryEmp.Width = 115
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllWokUnitToolStripMenuItem, Me.ApplyToWorkUnitGroupToolStripMenuItem, Me.ApplyByJobPosirionToolStripMenuItem, Me.CustomEmployeeListToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(194, 92)
        '
        'ApplyToAllWokUnitToolStripMenuItem
        '
        Me.ApplyToAllWokUnitToolStripMenuItem.Name = "ApplyToAllWokUnitToolStripMenuItem"
        Me.ApplyToAllWokUnitToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ApplyToAllWokUnitToolStripMenuItem.Text = "Apply to All Employee"
        '
        'ApplyToWorkUnitGroupToolStripMenuItem
        '
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Name = "ApplyToWorkUnitGroupToolStripMenuItem"
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Text = "Apply By Departement"
        '
        'ApplyByJobPosirionToolStripMenuItem
        '
        Me.ApplyByJobPosirionToolStripMenuItem.Name = "ApplyByJobPosirionToolStripMenuItem"
        Me.ApplyByJobPosirionToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ApplyByJobPosirionToolStripMenuItem.Text = "Apply By Job Position"
        '
        'CustomEmployeeListToolStripMenuItem
        '
        Me.CustomEmployeeListToolStripMenuItem.Name = "CustomEmployeeListToolStripMenuItem"
        Me.CustomEmployeeListToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.CustomEmployeeListToolStripMenuItem.Text = "Custom Employee List"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(500, 457)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.TabStop = False
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        Me.btnPrint.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(253, 419)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(101, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(360, 419)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(101, 26)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(501, 457)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(329, 457)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(155, 457)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(241, 457)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(415, 457)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(714, 451)
        Me.gboDetail.TabIndex = 32
        Me.gboDetail.TabStop = False
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker_CloseForm
        '
        '
        'FrmTrans_SPL_Setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 514)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.txtReadonly)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmTrans_SPL_Setting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Approval Group"
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvScope, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.gboDetail.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtReadonly As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyToAllWokUnitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToWorkUnitGroupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyByJobPosirionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomEmployeeListToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents chkApplyToAllEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvScope As System.Windows.Forms.DataGridView
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartmentCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartmentDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColUserNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserEmail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserAllowOverBudgetSPL As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColUserApproved As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserApprovedPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserPrimaryUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker_CloseForm As System.ComponentModel.BackgroundWorker
End Class
