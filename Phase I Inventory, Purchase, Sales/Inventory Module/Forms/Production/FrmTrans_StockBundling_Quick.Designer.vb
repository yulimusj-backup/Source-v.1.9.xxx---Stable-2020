<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_StockBundling_Quick
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItemWO = New System.Windows.Forms.DataGridView
        Me.ColWOTbl_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_Progress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_WONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_ProcessDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_CreateLog = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColWOTbl_View = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColWOTbl_Edit = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColWOTbl_Delete = New System.Windows.Forms.DataGridViewButtonColumn
        Me.TxtMONumber = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtJONumber = New System.Windows.Forms.TextBox
        Me.llbColumnSetting = New System.Windows.Forms.LinkLabel
        Me.btnFindJO = New System.Windows.Forms.Button
        Me.TxtIdentifierJO = New System.Windows.Forms.TextBox
        Me.TxtPrimaryMO = New System.Windows.Forms.TextBox
        Me.btnCreateLog = New System.Windows.Forms.Button
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.txtPrimaryProses = New System.Windows.Forms.TextBox
        Me.btnFilterProcess = New System.Windows.Forms.Button
        Me.btnFindProcess = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtProcessCode = New System.Windows.Forms.TextBox
        CType(Me.dgvItemWO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "JO Number"
        '
        'dgvItemWO
        '
        Me.dgvItemWO.AllowUserToAddRows = False
        Me.dgvItemWO.AllowUserToDeleteRows = False
        Me.dgvItemWO.AllowUserToResizeColumns = False
        Me.dgvItemWO.AllowUserToResizeRows = False
        Me.dgvItemWO.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemWO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemWO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemWO.ColumnHeadersHeight = 25
        Me.dgvItemWO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemWO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWOTbl_No, Me.ColWOTbl_Progress, Me.ColWOTbl_WONO, Me.ColWOTbl_CodeNo, Me.ColWOTbl_Desc, Me.ColWOTbl_ProcessCode, Me.ColWOTbl_ProcessDesc, Me.ColWOTbl_WorkUnit, Me.ColWOTbl_PrimaryProcess, Me.ColWOTbl_IdentifierWO, Me.ColWOTbl_IdentifierJO, Me.ColWOTbl_StartDate, Me.ColWOTbl_StartTime, Me.ColWOTbl_EndDate, Me.ColWOTbl_EndTime, Me.ColWOTbl_Remarks, Me.ColWOTbl_CreateLog, Me.ColWOTbl_View, Me.ColWOTbl_Edit, Me.ColWOTbl_Delete})
        Me.dgvItemWO.Location = New System.Drawing.Point(10, 76)
        Me.dgvItemWO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItemWO.Name = "dgvItemWO"
        Me.dgvItemWO.RowHeadersVisible = False
        Me.dgvItemWO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemWO.RowTemplate.Height = 24
        Me.dgvItemWO.Size = New System.Drawing.Size(820, 275)
        Me.dgvItemWO.TabIndex = 7
        '
        'ColWOTbl_No
        '
        Me.ColWOTbl_No.HeaderText = "No"
        Me.ColWOTbl_No.Name = "ColWOTbl_No"
        Me.ColWOTbl_No.Width = 46
        '
        'ColWOTbl_Progress
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWOTbl_Progress.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColWOTbl_Progress.HeaderText = "Progress (%)"
        Me.ColWOTbl_Progress.Name = "ColWOTbl_Progress"
        Me.ColWOTbl_Progress.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_Progress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_Progress.Width = 71
        '
        'ColWOTbl_WONO
        '
        Me.ColWOTbl_WONO.HeaderText = "WO No"
        Me.ColWOTbl_WONO.Name = "ColWOTbl_WONO"
        Me.ColWOTbl_WONO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_WONO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_WONO.Width = 49
        '
        'ColWOTbl_CodeNo
        '
        Me.ColWOTbl_CodeNo.HeaderText = "Code No"
        Me.ColWOTbl_CodeNo.Name = "ColWOTbl_CodeNo"
        Me.ColWOTbl_CodeNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_CodeNo.Width = 55
        '
        'ColWOTbl_Desc
        '
        Me.ColWOTbl_Desc.HeaderText = "Description"
        Me.ColWOTbl_Desc.Name = "ColWOTbl_Desc"
        Me.ColWOTbl_Desc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_Desc.Width = 66
        '
        'ColWOTbl_ProcessCode
        '
        Me.ColWOTbl_ProcessCode.HeaderText = "Process Code"
        Me.ColWOTbl_ProcessCode.Name = "ColWOTbl_ProcessCode"
        Me.ColWOTbl_ProcessCode.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_ProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_ProcessCode.Width = 79
        '
        'ColWOTbl_ProcessDesc
        '
        Me.ColWOTbl_ProcessDesc.HeaderText = "Process Description"
        Me.ColWOTbl_ProcessDesc.Name = "ColWOTbl_ProcessDesc"
        Me.ColWOTbl_ProcessDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_ProcessDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_ProcessDesc.Width = 107
        '
        'ColWOTbl_WorkUnit
        '
        Me.ColWOTbl_WorkUnit.HeaderText = "Work Unit"
        Me.ColWOTbl_WorkUnit.Name = "ColWOTbl_WorkUnit"
        Me.ColWOTbl_WorkUnit.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_WorkUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_WorkUnit.Width = 61
        '
        'ColWOTbl_PrimaryProcess
        '
        Me.ColWOTbl_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColWOTbl_PrimaryProcess.Name = "ColWOTbl_PrimaryProcess"
        Me.ColWOTbl_PrimaryProcess.Visible = False
        Me.ColWOTbl_PrimaryProcess.Width = 107
        '
        'ColWOTbl_IdentifierWO
        '
        Me.ColWOTbl_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColWOTbl_IdentifierWO.Name = "ColWOTbl_IdentifierWO"
        Me.ColWOTbl_IdentifierWO.Visible = False
        Me.ColWOTbl_IdentifierWO.Width = 94
        '
        'ColWOTbl_IdentifierJO
        '
        Me.ColWOTbl_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColWOTbl_IdentifierJO.Name = "ColWOTbl_IdentifierJO"
        Me.ColWOTbl_IdentifierJO.Visible = False
        Me.ColWOTbl_IdentifierJO.Width = 88
        '
        'ColWOTbl_StartDate
        '
        Me.ColWOTbl_StartDate.HeaderText = "Start"
        Me.ColWOTbl_StartDate.Name = "ColWOTbl_StartDate"
        Me.ColWOTbl_StartDate.Width = 54
        '
        'ColWOTbl_StartTime
        '
        Me.ColWOTbl_StartTime.HeaderText = ""
        Me.ColWOTbl_StartTime.Name = "ColWOTbl_StartTime"
        Me.ColWOTbl_StartTime.Width = 19
        '
        'ColWOTbl_EndDate
        '
        Me.ColWOTbl_EndDate.HeaderText = "End"
        Me.ColWOTbl_EndDate.Name = "ColWOTbl_EndDate"
        Me.ColWOTbl_EndDate.Width = 51
        '
        'ColWOTbl_EndTime
        '
        Me.ColWOTbl_EndTime.HeaderText = ""
        Me.ColWOTbl_EndTime.Name = "ColWOTbl_EndTime"
        Me.ColWOTbl_EndTime.Width = 19
        '
        'ColWOTbl_Remarks
        '
        Me.ColWOTbl_Remarks.HeaderText = "Remarks"
        Me.ColWOTbl_Remarks.Name = "ColWOTbl_Remarks"
        Me.ColWOTbl_Remarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_Remarks.Width = 55
        '
        'ColWOTbl_CreateLog
        '
        Me.ColWOTbl_CreateLog.HeaderText = "Create Log"
        Me.ColWOTbl_CreateLog.Name = "ColWOTbl_CreateLog"
        Me.ColWOTbl_CreateLog.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_CreateLog.Width = 65
        '
        'ColWOTbl_View
        '
        Me.ColWOTbl_View.HeaderText = "View Log"
        Me.ColWOTbl_View.Name = "ColWOTbl_View"
        Me.ColWOTbl_View.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_View.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColWOTbl_View.Width = 76
        '
        'ColWOTbl_Edit
        '
        Me.ColWOTbl_Edit.HeaderText = "Edit Log"
        Me.ColWOTbl_Edit.Name = "ColWOTbl_Edit"
        Me.ColWOTbl_Edit.Width = 52
        '
        'ColWOTbl_Delete
        '
        Me.ColWOTbl_Delete.HeaderText = "Delete Log"
        Me.ColWOTbl_Delete.Name = "ColWOTbl_Delete"
        Me.ColWOTbl_Delete.Width = 65
        '
        'TxtMONumber
        '
        Me.TxtMONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMONumber.Location = New System.Drawing.Point(75, 42)
        Me.TxtMONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtMONumber.Name = "TxtMONumber"
        Me.TxtMONumber.ReadOnly = True
        Me.TxtMONumber.Size = New System.Drawing.Size(134, 20)
        Me.TxtMONumber.TabIndex = 2
        Me.TxtMONumber.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 45)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "MO Number"
        '
        'TxtJONumber
        '
        Me.TxtJONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJONumber.Location = New System.Drawing.Point(75, 21)
        Me.TxtJONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtJONumber.Name = "TxtJONumber"
        Me.TxtJONumber.Size = New System.Drawing.Size(134, 20)
        Me.TxtJONumber.TabIndex = 0
        '
        'llbColumnSetting
        '
        Me.llbColumnSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbColumnSetting.AutoSize = True
        Me.llbColumnSetting.Location = New System.Drawing.Point(8, 353)
        Me.llbColumnSetting.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbColumnSetting.Name = "llbColumnSetting"
        Me.llbColumnSetting.Size = New System.Drawing.Size(88, 13)
        Me.llbColumnSetting.TabIndex = 8
        Me.llbColumnSetting.TabStop = True
        Me.llbColumnSetting.Text = "Columns Settings"
        '
        'btnFindJO
        '
        Me.btnFindJO.Location = New System.Drawing.Point(214, 21)
        Me.btnFindJO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindJO.Name = "btnFindJO"
        Me.btnFindJO.Size = New System.Drawing.Size(81, 19)
        Me.btnFindJO.TabIndex = 1
        Me.btnFindJO.TabStop = False
        Me.btnFindJO.Text = "Find"
        Me.btnFindJO.UseVisualStyleBackColor = True
        '
        'TxtIdentifierJO
        '
        Me.TxtIdentifierJO.Location = New System.Drawing.Point(75, 43)
        Me.TxtIdentifierJO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtIdentifierJO.Name = "TxtIdentifierJO"
        Me.TxtIdentifierJO.Size = New System.Drawing.Size(39, 20)
        Me.TxtIdentifierJO.TabIndex = 31
        Me.TxtIdentifierJO.TabStop = False
        Me.TxtIdentifierJO.Text = "-1"
        Me.TxtIdentifierJO.Visible = False
        '
        'TxtPrimaryMO
        '
        Me.TxtPrimaryMO.Location = New System.Drawing.Point(118, 43)
        Me.TxtPrimaryMO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryMO.Name = "TxtPrimaryMO"
        Me.TxtPrimaryMO.Size = New System.Drawing.Size(39, 20)
        Me.TxtPrimaryMO.TabIndex = 32
        Me.TxtPrimaryMO.TabStop = False
        Me.TxtPrimaryMO.Text = "-1"
        Me.TxtPrimaryMO.Visible = False
        '
        'btnCreateLog
        '
        Me.btnCreateLog.Location = New System.Drawing.Point(214, 43)
        Me.btnCreateLog.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCreateLog.Name = "btnCreateLog"
        Me.btnCreateLog.Size = New System.Drawing.Size(81, 20)
        Me.btnCreateLog.TabIndex = 3
        Me.btnCreateLog.Text = "Create Log"
        Me.btnCreateLog.UseVisualStyleBackColor = True
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(765, 62)
        Me.llbUncheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 6
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(711, 62)
        Me.llbCheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 5
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Location = New System.Drawing.Point(299, 10)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(531, 63)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(2, 15)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCodeNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCodeNo)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryProses)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilterProcess)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindProcess)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtProcessCode)
        Me.SplitContainer1.Size = New System.Drawing.Size(527, 46)
        Me.SplitContainer1.SplitterDistance = 262
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 25)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Description"
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDesc.Location = New System.Drawing.Point(68, 23)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(192, 20)
        Me.TxtDesc.TabIndex = 2
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCodeNo.Location = New System.Drawing.Point(203, 2)
        Me.btnFindCodeNo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(56, 19)
        Me.btnFindCodeNo.TabIndex = 1
        Me.btnFindCodeNo.TabStop = False
        Me.btnFindCodeNo.Text = "Find"
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 4)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Code No"
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(68, 2)
        Me.TxtCodeNo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(130, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'txtPrimaryProses
        '
        Me.txtPrimaryProses.Location = New System.Drawing.Point(25, 0)
        Me.txtPrimaryProses.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtPrimaryProses.Name = "txtPrimaryProses"
        Me.txtPrimaryProses.Size = New System.Drawing.Size(39, 20)
        Me.txtPrimaryProses.TabIndex = 37
        Me.txtPrimaryProses.TabStop = False
        Me.txtPrimaryProses.Text = "-1"
        Me.txtPrimaryProses.Visible = False
        '
        'btnFilterProcess
        '
        Me.btnFilterProcess.Location = New System.Drawing.Point(68, 23)
        Me.btnFilterProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFilterProcess.Name = "btnFilterProcess"
        Me.btnFilterProcess.Size = New System.Drawing.Size(56, 19)
        Me.btnFilterProcess.TabIndex = 2
        Me.btnFilterProcess.Text = "Filter"
        Me.btnFilterProcess.UseVisualStyleBackColor = True
        '
        'btnFindProcess
        '
        Me.btnFindProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProcess.Location = New System.Drawing.Point(204, 2)
        Me.btnFindProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindProcess.Name = "btnFindProcess"
        Me.btnFindProcess.Size = New System.Drawing.Size(56, 19)
        Me.btnFindProcess.TabIndex = 1
        Me.btnFindProcess.TabStop = False
        Me.btnFindProcess.Text = "Find"
        Me.btnFindProcess.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 4)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Process"
        '
        'txtProcessCode
        '
        Me.txtProcessCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProcessCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProcessCode.Location = New System.Drawing.Point(68, 2)
        Me.txtProcessCode.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtProcessCode.Name = "txtProcessCode"
        Me.txtProcessCode.Size = New System.Drawing.Size(132, 20)
        Me.txtProcessCode.TabIndex = 0
        '
        'FrmTrans_StockBundling_Quick
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(839, 376)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtPrimaryMO)
        Me.Controls.Add(Me.TxtIdentifierJO)
        Me.Controls.Add(Me.btnFindJO)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvItemWO)
        Me.Controls.Add(Me.TxtMONumber)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtJONumber)
        Me.Controls.Add(Me.llbColumnSetting)
        Me.Controls.Add(Me.btnCreateLog)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(855, 414)
        Me.Name = "FrmTrans_StockBundling_Quick"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Log Input"
        CType(Me.dgvItemWO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItemWO As System.Windows.Forms.DataGridView
    Friend WithEvents TxtMONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtJONumber As System.Windows.Forms.TextBox
    Friend WithEvents llbColumnSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents btnFindJO As System.Windows.Forms.Button
    Friend WithEvents TxtIdentifierJO As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryMO As System.Windows.Forms.TextBox
    Friend WithEvents btnCreateLog As System.Windows.Forms.Button
    Friend WithEvents ColWOTbl_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_Progress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_WONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_ProcessDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_CreateLog As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColWOTbl_View As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColWOTbl_Edit As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColWOTbl_Delete As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents btnFilterProcess As System.Windows.Forms.Button
    Friend WithEvents btnFindProcess As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtProcessCode As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryProses As System.Windows.Forms.TextBox
End Class
