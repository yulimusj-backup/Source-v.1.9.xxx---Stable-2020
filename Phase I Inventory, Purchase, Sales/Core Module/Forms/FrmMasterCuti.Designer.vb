<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCuti
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.gboEmployeeData = New System.Windows.Forms.GroupBox
        Me.LblAdvance = New System.Windows.Forms.LinkLabel
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllEmployeeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyByDepartementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyToAJobPositionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CopyFromOtherDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.AddMonthlyLeaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AddLongLeaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.ChkIsDefault = New System.Windows.Forms.CheckBox
        Me.LblDays = New System.Windows.Forms.Label
        Me.TxtLeaveQty = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeaveAddition = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColLeaveDeduction = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboEmployeeData.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(269, 152)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(77, 20)
        Me.TxtReadOnly.TabIndex = 35
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(363, 332)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(59, 71)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(549, 40)
        Me.TxtRemarks.TabIndex = 2
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(191, 332)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Remarks"
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(209, 293)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(247, 19)
        Me.TxtDesc.MaxLength = 250
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(362, 20)
        Me.TxtDesc.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(181, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Description"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Code"
        '
        'TxtCode
        '
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(59, 19)
        Me.TxtCode.MaxLength = 150
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(116, 20)
        Me.TxtCode.TabIndex = 0
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(314, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 369)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(644, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 54
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(314, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(325, 293)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(314, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 391)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(644, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 53
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(314, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(449, 332)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(277, 332)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AllowUserToResizeColumns = False
        Me.dgvEmployee.AllowUserToResizeRows = False
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCode, Me.ColName, Me.ColLeaveAddition, Me.ColLeaveDeduction, Me.ColBalance, Me.ColRemarks, Me.ColPrimary, Me.ColIdentifier})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(3, 16)
        Me.dgvEmployee.MultiSelect = False
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(597, 151)
        Me.dgvEmployee.TabIndex = 0
        '
        'gboEmployeeData
        '
        Me.gboEmployeeData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboEmployeeData.Controls.Add(Me.dgvEmployee)
        Me.gboEmployeeData.Controls.Add(Me.LblAdvance)
        Me.gboEmployeeData.Location = New System.Drawing.Point(6, 117)
        Me.gboEmployeeData.Name = "gboEmployeeData"
        Me.gboEmployeeData.Size = New System.Drawing.Size(603, 170)
        Me.gboEmployeeData.TabIndex = 6
        Me.gboEmployeeData.TabStop = False
        Me.gboEmployeeData.Text = "Employee Data"
        '
        'LblAdvance
        '
        Me.LblAdvance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAdvance.AutoSize = True
        Me.LblAdvance.ContextMenuStrip = Me.ContextMenuStrip1
        Me.LblAdvance.Location = New System.Drawing.Point(547, 0)
        Me.LblAdvance.Name = "LblAdvance"
        Me.LblAdvance.Size = New System.Drawing.Size(50, 13)
        Me.LblAdvance.TabIndex = 0
        Me.LblAdvance.TabStop = True
        Me.LblAdvance.Text = "Advance"
        Me.ToolTip1.SetToolTip(Me.LblAdvance, "Click to show menu")
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllEmployeeToolStripMenuItem, Me.ApplyByDepartementToolStripMenuItem, Me.ApplyToAJobPositionToolStripMenuItem, Me.CopyFromOtherDataToolStripMenuItem, Me.ToolStripSeparator1, Me.AddMonthlyLeaveToolStripMenuItem, Me.AddLongLeaveToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(201, 142)
        '
        'ApplyToAllEmployeeToolStripMenuItem
        '
        Me.ApplyToAllEmployeeToolStripMenuItem.Name = "ApplyToAllEmployeeToolStripMenuItem"
        Me.ApplyToAllEmployeeToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ApplyToAllEmployeeToolStripMenuItem.Text = "Apply To All Employee"
        '
        'ApplyByDepartementToolStripMenuItem
        '
        Me.ApplyByDepartementToolStripMenuItem.Name = "ApplyByDepartementToolStripMenuItem"
        Me.ApplyByDepartementToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ApplyByDepartementToolStripMenuItem.Text = "Apply To A Department"
        '
        'ApplyToAJobPositionToolStripMenuItem
        '
        Me.ApplyToAJobPositionToolStripMenuItem.Name = "ApplyToAJobPositionToolStripMenuItem"
        Me.ApplyToAJobPositionToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ApplyToAJobPositionToolStripMenuItem.Text = "Apply To A Job Position"
        '
        'CopyFromOtherDataToolStripMenuItem
        '
        Me.CopyFromOtherDataToolStripMenuItem.Name = "CopyFromOtherDataToolStripMenuItem"
        Me.CopyFromOtherDataToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.CopyFromOtherDataToolStripMenuItem.Text = "Copy From Other Data"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(197, 6)
        '
        'AddMonthlyLeaveToolStripMenuItem
        '
        Me.AddMonthlyLeaveToolStripMenuItem.Name = "AddMonthlyLeaveToolStripMenuItem"
        Me.AddMonthlyLeaveToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.AddMonthlyLeaveToolStripMenuItem.Text = "Add Monthly Leave"
        '
        'AddLongLeaveToolStripMenuItem
        '
        Me.AddLongLeaveToolStripMenuItem.Name = "AddLongLeaveToolStripMenuItem"
        Me.AddLongLeaveToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.AddLongLeaveToolStripMenuItem.Text = "Add Long Leave"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(57, 19)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(105, 332)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.ChkIsDefault)
        Me.gboDetail.Controls.Add(Me.LblDays)
        Me.gboDetail.Controls.Add(Me.TxtLeaveQty)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.dtpPeriod)
        Me.gboDetail.Controls.Add(Me.gboEmployeeData)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtDesc)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Controls.Add(Me.TxtCode)
        Me.gboDetail.Location = New System.Drawing.Point(9, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(615, 325)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'ChkIsDefault
        '
        Me.ChkIsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkIsDefault.AutoSize = True
        Me.ChkIsDefault.Location = New System.Drawing.Point(508, 49)
        Me.ChkIsDefault.Name = "ChkIsDefault"
        Me.ChkIsDefault.Size = New System.Drawing.Size(93, 17)
        Me.ChkIsDefault.TabIndex = 5
        Me.ChkIsDefault.Text = "Default Leave"
        Me.ChkIsDefault.UseVisualStyleBackColor = True
        '
        'LblDays
        '
        Me.LblDays.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblDays.AutoSize = True
        Me.LblDays.Location = New System.Drawing.Point(476, 49)
        Me.LblDays.Name = "LblDays"
        Me.LblDays.Size = New System.Drawing.Size(31, 13)
        Me.LblDays.TabIndex = 41
        Me.LblDays.Text = "Days"
        '
        'TxtLeaveQty
        '
        Me.TxtLeaveQty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLeaveQty.Location = New System.Drawing.Point(247, 45)
        Me.TxtLeaveQty.MaxLength = 12
        Me.TxtLeaveQty.Name = "TxtLeaveQty"
        Me.TxtLeaveQty.Size = New System.Drawing.Size(225, 20)
        Me.TxtLeaveQty.TabIndex = 4
        Me.TxtLeaveQty.Text = "0.0"
        Me.TxtLeaveQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(181, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "Leave Qty"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 13)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Period"
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(59, 45)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.ShowUpDown = True
        Me.dtpPeriod.Size = New System.Drawing.Size(116, 20)
        Me.dtpPeriod.TabIndex = 1
        '
        'BackgroundWorker1
        '
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.Width = 57
        '
        'ColName
        '
        Me.ColName.HeaderText = "Name"
        Me.ColName.Name = "ColName"
        Me.ColName.Width = 60
        '
        'ColLeaveAddition
        '
        Me.ColLeaveAddition.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLeaveAddition.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColLeaveAddition.HeaderText = "Leave Addition"
        Me.ColLeaveAddition.Name = "ColLeaveAddition"
        Me.ColLeaveAddition.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLeaveAddition.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColLeaveAddition.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColLeaveAddition.Width = 103
        '
        'ColLeaveDeduction
        '
        Me.ColLeaveDeduction.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLeaveDeduction.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColLeaveDeduction.HeaderText = "Leave Deduction"
        Me.ColLeaveDeduction.Name = "ColLeaveDeduction"
        Me.ColLeaveDeduction.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLeaveDeduction.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColLeaveDeduction.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColLeaveDeduction.Width = 114
        '
        'ColBalance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColBalance.TrackVisitedState = False
        Me.ColBalance.Width = 71
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary Karyawan"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 116
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.ReadOnly = True
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 72
        '
        'FrmMasterCuti
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(644, 413)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(654, 450)
        Me.Name = "FrmMasterCuti"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave Data"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboEmployeeData.ResumeLayout(False)
        Me.gboEmployeeData.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents gboEmployeeData As System.Windows.Forms.GroupBox
    Friend WithEvents LblAdvance As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents TxtLeaveQty As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents ApplyToAllEmployeeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyByDepartementToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToAJobPositionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyFromOtherDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LblDays As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ChkIsDefault As System.Windows.Forms.CheckBox
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AddMonthlyLeaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddLongLeaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveAddition As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColLeaveDeduction As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
