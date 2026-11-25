<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_SplitMerge
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColMoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartHour = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndHour = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProgress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPlanStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.ChkSetAsLast = New System.Windows.Forms.CheckBox
        Me.TxtWUDesc = New System.Windows.Forms.TextBox
        Me.BtnFindWU = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtWUCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryWU = New System.Windows.Forms.TextBox
        Me.DgvItem2 = New System.Windows.Forms.DataGridView
        Me.Col2Insert = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Col2IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2MO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2JO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2WO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2Rate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2UnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2WorkHour = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2StartDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2EndDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2Progress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2PrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2ReadOnly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2PlanStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Btn1Back = New System.Windows.Forms.Button
        Me.Btn1Next = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(173, 602)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(103, 17)
        Me.LinkLabel3.TabIndex = 20
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Column Setting"
        Me.LinkLabel3.Visible = False
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(857, 2)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(112, 17)
        Me.LinkLabel2.TabIndex = 19
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All Item"
        Me.LinkLabel2.Visible = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(733, 2)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(118, 17)
        Me.LinkLabel1.TabIndex = 18
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All JO Item"
        Me.LinkLabel1.Visible = False
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Enabled = False
        Me.BtnOK.Location = New System.Drawing.Point(827, 597)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(75, 26)
        Me.BtnOK.TabIndex = 13
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(908, 597)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(75, 26)
        Me.BtnCancel.TabIndex = 12
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColModify, Me.ColMoNo, Me.ColJONo, Me.ColProcessCode, Me.ColWONo, Me.ColCodeNo, Me.ColDesc, Me.ColRate, Me.ColQty, Me.ColUnit, Me.ColWH, Me.ColStartDay, Me.ColStartDate, Me.ColStartHour, Me.ColEndDay, Me.ColEndDate, Me.ColEndHour, Me.ColProgress, Me.ColPrimaryMO, Me.ColIdentifierJO, Me.ColPrimaryProcess, Me.colIdentifierWO, Me.ColPrimaryCodeNo, Me.ColPrimaryUnit, Me.ColPlanStartTime})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvItem.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvItem.Location = New System.Drawing.Point(5, 22)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(963, 527)
        Me.dgvItem.TabIndex = 11
        '
        'ColModify
        '
        Me.ColModify.Frozen = True
        Me.ColModify.HeaderText = "Modify"
        Me.ColModify.Name = "ColModify"
        Me.ColModify.Width = 55
        '
        'ColMoNo
        '
        Me.ColMoNo.Frozen = True
        Me.ColMoNo.HeaderText = "Mo No"
        Me.ColMoNo.Name = "ColMoNo"
        Me.ColMoNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMoNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMoNo.Width = 55
        '
        'ColJONo
        '
        Me.ColJONo.Frozen = True
        Me.ColJONo.HeaderText = "JO No"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColJONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJONo.Width = 54
        '
        'ColProcessCode
        '
        Me.ColProcessCode.Frozen = True
        Me.ColProcessCode.HeaderText = "Process"
        Me.ColProcessCode.Name = "ColProcessCode"
        Me.ColProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProcessCode.Width = 65
        '
        'ColWONo
        '
        Me.ColWONo.Frozen = True
        Me.ColWONo.HeaderText = "WO No"
        Me.ColWONo.Name = "ColWONo"
        Me.ColWONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWONo.Width = 60
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 69
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 85
        '
        'ColRate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColRate.HeaderText = "Rate (Unit/Hour)"
        Me.ColRate.Name = "ColRate"
        Me.ColRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRate.Width = 118
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 36
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 39
        '
        'ColWH
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWH.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColWH.HeaderText = "W/H"
        Me.ColWH.Name = "ColWH"
        Me.ColWH.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWH.Width = 41
        '
        'ColStartDay
        '
        Me.ColStartDay.HeaderText = "Start"
        Me.ColStartDay.Name = "ColStartDay"
        Me.ColStartDay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStartDay.Width = 44
        '
        'ColStartDate
        '
        Me.ColStartDate.HeaderText = ""
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStartDate.Width = 5
        '
        'ColStartHour
        '
        Me.ColStartHour.HeaderText = ""
        Me.ColStartHour.Name = "ColStartHour"
        Me.ColStartHour.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStartHour.Width = 5
        '
        'ColEndDay
        '
        Me.ColEndDay.HeaderText = "End"
        Me.ColEndDay.Name = "ColEndDay"
        Me.ColEndDay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEndDay.Width = 39
        '
        'ColEndDate
        '
        Me.ColEndDate.HeaderText = ""
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEndDate.Width = 5
        '
        'ColEndHour
        '
        Me.ColEndHour.HeaderText = ""
        Me.ColEndHour.Name = "ColEndHour"
        Me.ColEndHour.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEndHour.Width = 5
        '
        'ColProgress
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColProgress.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColProgress.HeaderText = "Progress"
        Me.ColProgress.Name = "ColProgress"
        Me.ColProgress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProgress.Width = 71
        '
        'ColPrimaryMO
        '
        Me.ColPrimaryMO.HeaderText = "Primary MO"
        Me.ColPrimaryMO.Name = "ColPrimaryMO"
        Me.ColPrimaryMO.ReadOnly = True
        Me.ColPrimaryMO.Visible = False
        Me.ColPrimaryMO.Width = 107
        '
        'ColIdentifierJO
        '
        Me.ColIdentifierJO.HeaderText = "Identifier JO"
        Me.ColIdentifierJO.Name = "ColIdentifierJO"
        Me.ColIdentifierJO.ReadOnly = True
        Me.ColIdentifierJO.Visible = False
        Me.ColIdentifierJO.Width = 109
        '
        'ColPrimaryProcess
        '
        Me.ColPrimaryProcess.HeaderText = "Primary Process"
        Me.ColPrimaryProcess.Name = "ColPrimaryProcess"
        Me.ColPrimaryProcess.ReadOnly = True
        Me.ColPrimaryProcess.Visible = False
        Me.ColPrimaryProcess.Width = 136
        '
        'colIdentifierWO
        '
        Me.colIdentifierWO.HeaderText = "Identifier WO"
        Me.colIdentifierWO.Name = "colIdentifierWO"
        Me.colIdentifierWO.ReadOnly = True
        Me.colIdentifierWO.Visible = False
        Me.colIdentifierWO.Width = 115
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 140
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 110
        '
        'ColPlanStartTime
        '
        Me.ColPlanStartTime.HeaderText = "Plan Start Time"
        Me.ColPlanStartTime.Name = "ColPlanStartTime"
        Me.ColPlanStartTime.ReadOnly = True
        Me.ColPlanStartTime.Visible = False
        Me.ColPlanStartTime.Width = 130
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(7, 10)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(983, 583)
        Me.TabControl1.TabIndex = 22
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Controls.Add(Me.LinkLabel2)
        Me.TabPage1.Controls.Add(Me.LinkLabel1)
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Size = New System.Drawing.Size(975, 554)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Select item to split or move"
        '
        'Panel1
        '
        Me.Panel1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Location = New System.Drawing.Point(235, 246)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 62)
        Me.Panel1.TabIndex = 21
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 34)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Label2"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(4, 9)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(497, 22)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 0
        Me.ProgressBar1.Value = 25
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.ChkSetAsLast)
        Me.TabPage2.Controls.Add(Me.TxtWUDesc)
        Me.TabPage2.Controls.Add(Me.BtnFindWU)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.TxtWUCode)
        Me.TabPage2.Controls.Add(Me.TxtPrimaryWU)
        Me.TabPage2.Controls.Add(Me.DgvItem2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Size = New System.Drawing.Size(975, 554)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Select workunit target"
        '
        'ChkSetAsLast
        '
        Me.ChkSetAsLast.AutoSize = True
        Me.ChkSetAsLast.Checked = True
        Me.ChkSetAsLast.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkSetAsLast.Location = New System.Drawing.Point(629, 7)
        Me.ChkSetAsLast.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkSetAsLast.Name = "ChkSetAsLast"
        Me.ChkSetAsLast.Size = New System.Drawing.Size(131, 21)
        Me.ChkSetAsLast.TabIndex = 3
        Me.ChkSetAsLast.Text = "Set as Last Item"
        Me.ChkSetAsLast.UseVisualStyleBackColor = True
        '
        'TxtWUDesc
        '
        Me.TxtWUDesc.Location = New System.Drawing.Point(288, 5)
        Me.TxtWUDesc.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtWUDesc.Name = "TxtWUDesc"
        Me.TxtWUDesc.ReadOnly = True
        Me.TxtWUDesc.Size = New System.Drawing.Size(336, 22)
        Me.TxtWUDesc.TabIndex = 2
        Me.TxtWUDesc.TabStop = False
        '
        'BtnFindWU
        '
        Me.BtnFindWU.Location = New System.Drawing.Point(207, 5)
        Me.BtnFindWU.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnFindWU.Name = "BtnFindWU"
        Me.BtnFindWU.Size = New System.Drawing.Size(75, 26)
        Me.BtnFindWU.TabIndex = 1
        Me.BtnFindWU.Text = "Find"
        Me.BtnFindWU.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Workunit"
        '
        'TxtWUCode
        '
        Me.TxtWUCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtWUCode.Location = New System.Drawing.Point(83, 5)
        Me.TxtWUCode.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtWUCode.Name = "TxtWUCode"
        Me.TxtWUCode.Size = New System.Drawing.Size(119, 22)
        Me.TxtWUCode.TabIndex = 0
        '
        'TxtPrimaryWU
        '
        Me.TxtPrimaryWU.Location = New System.Drawing.Point(157, 5)
        Me.TxtPrimaryWU.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtPrimaryWU.Name = "TxtPrimaryWU"
        Me.TxtPrimaryWU.Size = New System.Drawing.Size(43, 22)
        Me.TxtPrimaryWU.TabIndex = 17
        Me.TxtPrimaryWU.TabStop = False
        Me.TxtPrimaryWU.Text = "-1"
        '
        'DgvItem2
        '
        Me.DgvItem2.AllowUserToAddRows = False
        Me.DgvItem2.AllowUserToDeleteRows = False
        Me.DgvItem2.AllowUserToResizeColumns = False
        Me.DgvItem2.AllowUserToResizeRows = False
        Me.DgvItem2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem2.ColumnHeadersHeight = 25
        Me.DgvItem2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col2Insert, Me.Col2IdentifierJO, Me.Col2IdentifierWO, Me.Col2MO_No, Me.Col2JO_No, Me.Col2ProcessCode, Me.Col2WO_No, Me.Col2CodeNo, Me.Col2Desc, Me.Col2Rate, Me.Col2Qty, Me.Col2UnitCode, Me.Col2WorkHour, Me.Col2StartDay, Me.Col2StartDate, Me.Col2StartTime, Me.Col2EndDay, Me.Col2EndDate, Me.Col2EndTime, Me.Col2Progress, Me.Col2PrimaryMO, Me.Col2PrimaryProcess, Me.Col2PrimaryCodeNo, Me.Col2PrimaryUnit, Me.Col2ReadOnly, Me.Col2PlanStartTime})
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgvItem2.DefaultCellStyle = DataGridViewCellStyle9
        Me.DgvItem2.Location = New System.Drawing.Point(5, 33)
        Me.DgvItem2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvItem2.Name = "DgvItem2"
        Me.DgvItem2.RowHeadersVisible = False
        Me.DgvItem2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem2.RowTemplate.Height = 24
        Me.DgvItem2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvItem2.Size = New System.Drawing.Size(963, 516)
        Me.DgvItem2.TabIndex = 4
        '
        'Col2Insert
        '
        Me.Col2Insert.Frozen = True
        Me.Col2Insert.HeaderText = "Insert"
        Me.Col2Insert.Name = "Col2Insert"
        Me.Col2Insert.Width = 49
        '
        'Col2IdentifierJO
        '
        Me.Col2IdentifierJO.HeaderText = "identifier jo"
        Me.Col2IdentifierJO.Name = "Col2IdentifierJO"
        Me.Col2IdentifierJO.ReadOnly = True
        Me.Col2IdentifierJO.Visible = False
        Me.Col2IdentifierJO.Width = 102
        '
        'Col2IdentifierWO
        '
        Me.Col2IdentifierWO.HeaderText = "identifier wo"
        Me.Col2IdentifierWO.Name = "Col2IdentifierWO"
        Me.Col2IdentifierWO.ReadOnly = True
        Me.Col2IdentifierWO.Visible = False
        Me.Col2IdentifierWO.Width = 108
        '
        'Col2MO_No
        '
        Me.Col2MO_No.Frozen = True
        Me.Col2MO_No.HeaderText = "Mo No"
        Me.Col2MO_No.Name = "Col2MO_No"
        Me.Col2MO_No.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col2MO_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2MO_No.Width = 55
        '
        'Col2JO_No
        '
        Me.Col2JO_No.Frozen = True
        Me.Col2JO_No.HeaderText = "JO No"
        Me.Col2JO_No.Name = "Col2JO_No"
        Me.Col2JO_No.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col2JO_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2JO_No.Width = 54
        '
        'Col2ProcessCode
        '
        Me.Col2ProcessCode.Frozen = True
        Me.Col2ProcessCode.HeaderText = "Process"
        Me.Col2ProcessCode.Name = "Col2ProcessCode"
        Me.Col2ProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2ProcessCode.Width = 65
        '
        'Col2WO_No
        '
        Me.Col2WO_No.Frozen = True
        Me.Col2WO_No.HeaderText = "WO No"
        Me.Col2WO_No.Name = "Col2WO_No"
        Me.Col2WO_No.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col2WO_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2WO_No.Width = 60
        '
        'Col2CodeNo
        '
        Me.Col2CodeNo.Frozen = True
        Me.Col2CodeNo.HeaderText = "Code No"
        Me.Col2CodeNo.Name = "Col2CodeNo"
        Me.Col2CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2CodeNo.Width = 69
        '
        'Col2Desc
        '
        Me.Col2Desc.HeaderText = "Description"
        Me.Col2Desc.Name = "Col2Desc"
        Me.Col2Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2Desc.Width = 85
        '
        'Col2Rate
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col2Rate.DefaultCellStyle = DataGridViewCellStyle6
        Me.Col2Rate.HeaderText = "Rate (Unit/Hour)"
        Me.Col2Rate.Name = "Col2Rate"
        Me.Col2Rate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2Rate.Width = 118
        '
        'Col2Qty
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col2Qty.DefaultCellStyle = DataGridViewCellStyle7
        Me.Col2Qty.HeaderText = "Qty"
        Me.Col2Qty.Name = "Col2Qty"
        Me.Col2Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2Qty.Width = 36
        '
        'Col2UnitCode
        '
        Me.Col2UnitCode.HeaderText = "Unit"
        Me.Col2UnitCode.Name = "Col2UnitCode"
        Me.Col2UnitCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2UnitCode.Width = 39
        '
        'Col2WorkHour
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col2WorkHour.DefaultCellStyle = DataGridViewCellStyle8
        Me.Col2WorkHour.HeaderText = "W/H"
        Me.Col2WorkHour.Name = "Col2WorkHour"
        Me.Col2WorkHour.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2WorkHour.Width = 41
        '
        'Col2StartDay
        '
        Me.Col2StartDay.HeaderText = "Start"
        Me.Col2StartDay.Name = "Col2StartDay"
        Me.Col2StartDay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2StartDay.Width = 44
        '
        'Col2StartDate
        '
        Me.Col2StartDate.HeaderText = ""
        Me.Col2StartDate.Name = "Col2StartDate"
        Me.Col2StartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2StartDate.Width = 5
        '
        'Col2StartTime
        '
        Me.Col2StartTime.HeaderText = ""
        Me.Col2StartTime.Name = "Col2StartTime"
        Me.Col2StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2StartTime.Width = 5
        '
        'Col2EndDay
        '
        Me.Col2EndDay.HeaderText = "End"
        Me.Col2EndDay.Name = "Col2EndDay"
        Me.Col2EndDay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2EndDay.Width = 39
        '
        'Col2EndDate
        '
        Me.Col2EndDate.HeaderText = ""
        Me.Col2EndDate.Name = "Col2EndDate"
        Me.Col2EndDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2EndDate.Width = 5
        '
        'Col2EndTime
        '
        Me.Col2EndTime.HeaderText = ""
        Me.Col2EndTime.Name = "Col2EndTime"
        Me.Col2EndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2EndTime.Width = 5
        '
        'Col2Progress
        '
        Me.Col2Progress.HeaderText = "Progress"
        Me.Col2Progress.Name = "Col2Progress"
        Me.Col2Progress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col2Progress.Width = 71
        '
        'Col2PrimaryMO
        '
        Me.Col2PrimaryMO.HeaderText = "Primary MO"
        Me.Col2PrimaryMO.Name = "Col2PrimaryMO"
        Me.Col2PrimaryMO.Visible = False
        Me.Col2PrimaryMO.Width = 107
        '
        'Col2PrimaryProcess
        '
        Me.Col2PrimaryProcess.HeaderText = "Primary Process"
        Me.Col2PrimaryProcess.Name = "Col2PrimaryProcess"
        Me.Col2PrimaryProcess.Visible = False
        Me.Col2PrimaryProcess.Width = 136
        '
        'Col2PrimaryCodeNo
        '
        Me.Col2PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.Col2PrimaryCodeNo.Name = "Col2PrimaryCodeNo"
        Me.Col2PrimaryCodeNo.Visible = False
        Me.Col2PrimaryCodeNo.Width = 140
        '
        'Col2PrimaryUnit
        '
        Me.Col2PrimaryUnit.HeaderText = "Primary Unit"
        Me.Col2PrimaryUnit.Name = "Col2PrimaryUnit"
        Me.Col2PrimaryUnit.Visible = False
        Me.Col2PrimaryUnit.Width = 110
        '
        'Col2ReadOnly
        '
        Me.Col2ReadOnly.HeaderText = "Read Only"
        Me.Col2ReadOnly.Name = "Col2ReadOnly"
        Me.Col2ReadOnly.ReadOnly = True
        Me.Col2ReadOnly.Visible = False
        '
        'Col2PlanStartTime
        '
        Me.Col2PlanStartTime.HeaderText = "Plan Start Time"
        Me.Col2PlanStartTime.Name = "Col2PlanStartTime"
        Me.Col2PlanStartTime.ReadOnly = True
        Me.Col2PlanStartTime.Visible = False
        Me.Col2PlanStartTime.Width = 130
        '
        'Btn1Back
        '
        Me.Btn1Back.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Btn1Back.Enabled = False
        Me.Btn1Back.Location = New System.Drawing.Point(12, 597)
        Me.Btn1Back.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Btn1Back.Name = "Btn1Back"
        Me.Btn1Back.Size = New System.Drawing.Size(75, 26)
        Me.Btn1Back.TabIndex = 23
        Me.Btn1Back.Text = "Back"
        Me.Btn1Back.UseVisualStyleBackColor = True
        '
        'Btn1Next
        '
        Me.Btn1Next.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Btn1Next.Location = New System.Drawing.Point(93, 597)
        Me.Btn1Next.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Btn1Next.Name = "Btn1Next"
        Me.Btn1Next.Size = New System.Drawing.Size(75, 26)
        Me.Btn1Next.TabIndex = 24
        Me.Btn1Next.Text = "Next"
        Me.Btn1Next.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTrans_ProductionSchedule_SplitMerge
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(997, 635)
        Me.Controls.Add(Me.LinkLabel3)
        Me.Controls.Add(Me.Btn1Next)
        Me.Controls.Add(Me.Btn1Back)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.TabControl1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1005, 672)
        Me.Name = "FrmTrans_ProductionSchedule_SplitMerge"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Split-Merge Schedule"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DgvItem2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Btn1Back As System.Windows.Forms.Button
    Friend WithEvents Btn1Next As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColMoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartHour As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndHour As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProgress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPlanStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ChkSetAsLast As System.Windows.Forms.CheckBox
    Friend WithEvents TxtWUDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindWU As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtWUCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryWU As System.Windows.Forms.TextBox
    Friend WithEvents DgvItem2 As System.Windows.Forms.DataGridView
    Friend WithEvents Col2Insert As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Col2IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2MO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2JO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2WO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2Rate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2UnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2WorkHour As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2StartDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2EndDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2Progress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2PrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2ReadOnly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2PlanStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
End Class
