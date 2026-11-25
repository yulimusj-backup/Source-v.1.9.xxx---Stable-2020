<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterProductionRate
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
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.TslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.BtnOk = New System.Windows.Forms.Button
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.BtnNew = New System.Windows.Forms.Button
        Me.BtnFind = New System.Windows.Forms.Button
        Me.GboDetail = New System.Windows.Forms.GroupBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvWorkUnit = New System.Windows.Forms.DataGridView
        Me.ColWorkUnit_Priority = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_Rate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_Efficiency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_EffectiveRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnit_PrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtProcessDesc = New System.Windows.Forms.TextBox
        Me.btnFindProcess = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtProcessCode = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtDocumentNo = New System.Windows.Forms.TextBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.txtPrimaryProcess = New System.Windows.Forms.TextBox
        Me.txtDependency = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.cboDependency = New System.Windows.Forms.ComboBox
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnApply = New System.Windows.Forms.Button
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BgwDelDgvWorkUnit = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.GboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvWorkUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnExit
        '
        Me.BtnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnExit.Location = New System.Drawing.Point(522, 431)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(88, 26)
        Me.BtnExit.TabIndex = 6
        Me.BtnExit.Text = "E&xit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnCancel.Location = New System.Drawing.Point(326, 388)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(88, 26)
        Me.BtnCancel.TabIndex = 3
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslModifiedBy, Me.TslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 472)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(670, 22)
        Me.StatusStrip1.TabIndex = 29
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TslModifiedBy
        '
        Me.TslModifiedBy.Name = "TslModifiedBy"
        Me.TslModifiedBy.Size = New System.Drawing.Size(327, 17)
        Me.TslModifiedBy.Spring = True
        Me.TslModifiedBy.Text = "Modified By:"
        '
        'TslModifiedDate
        '
        Me.TslModifiedDate.Name = "TslModifiedDate"
        Me.TslModifiedDate.Size = New System.Drawing.Size(327, 17)
        Me.TslModifiedDate.Spring = True
        Me.TslModifiedDate.Text = "Modified Date:"
        '
        'TslCreatedDate
        '
        Me.TslCreatedDate.Name = "TslCreatedDate"
        Me.TslCreatedDate.Size = New System.Drawing.Size(327, 17)
        Me.TslCreatedDate.Spring = True
        Me.TslCreatedDate.Text = "Created Date:"
        '
        'TslCreatedBy
        '
        Me.TslCreatedBy.Name = "TslCreatedBy"
        Me.TslCreatedBy.Size = New System.Drawing.Size(327, 17)
        Me.TslCreatedBy.Spring = True
        Me.TslCreatedBy.Text = "Created By:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslCreatedBy, Me.TslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 494)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(670, 22)
        Me.StatusStrip2.TabIndex = 30
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'BtnOk
        '
        Me.BtnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnOk.Location = New System.Drawing.Point(228, 388)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(88, 26)
        Me.BtnOk.TabIndex = 2
        Me.BtnOk.Text = "&OK"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnDelete.Location = New System.Drawing.Point(334, 431)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(88, 26)
        Me.BtnDelete.TabIndex = 4
        Me.BtnDelete.Text = "&Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnEdit.Location = New System.Drawing.Point(239, 431)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(88, 26)
        Me.BtnEdit.TabIndex = 3
        Me.BtnEdit.Text = "&Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnNew
        '
        Me.BtnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnNew.Location = New System.Drawing.Point(145, 431)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(88, 26)
        Me.BtnNew.TabIndex = 2
        Me.BtnNew.Text = "&New"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnFind.Location = New System.Drawing.Point(52, 431)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(88, 26)
        Me.BtnFind.TabIndex = 1
        Me.BtnFind.Text = "&Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'GboDetail
        '
        Me.GboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail.Controls.Add(Me.TabControl1)
        Me.GboDetail.Controls.Add(Me.SplitContainer1)
        Me.GboDetail.Controls.Add(Me.BtnCancel)
        Me.GboDetail.Controls.Add(Me.BtnOk)
        Me.GboDetail.Location = New System.Drawing.Point(10, 2)
        Me.GboDetail.Name = "GboDetail"
        Me.GboDetail.Size = New System.Drawing.Size(643, 422)
        Me.GboDetail.TabIndex = 0
        Me.GboDetail.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(5, 84)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(632, 297)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvWorkUnit)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(624, 271)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Production Rate"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvWorkUnit
        '
        Me.DgvWorkUnit.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWorkUnit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWorkUnit.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWorkUnit_Priority, Me.ColWorkUnit_Code, Me.ColWorkUnit_Desc, Me.ColWorkUnit_Rate, Me.ColWorkUnit_Efficiency, Me.ColWorkUnit_EffectiveRate, Me.ColWorkUnit_Remarks, Me.ColWorkUnit_PrimaryWorkUnit})
        Me.DgvWorkUnit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWorkUnit.Location = New System.Drawing.Point(2, 2)
        Me.DgvWorkUnit.MultiSelect = False
        Me.DgvWorkUnit.Name = "DgvWorkUnit"
        Me.DgvWorkUnit.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvWorkUnit.RowTemplate.Height = 23
        Me.DgvWorkUnit.Size = New System.Drawing.Size(620, 267)
        Me.DgvWorkUnit.StandardTab = True
        Me.DgvWorkUnit.TabIndex = 2
        Me.DgvWorkUnit.TabStop = False
        '
        'ColWorkUnit_Priority
        '
        Me.ColWorkUnit_Priority.HeaderText = "Priority"
        Me.ColWorkUnit_Priority.Name = "ColWorkUnit_Priority"
        Me.ColWorkUnit_Priority.ReadOnly = True
        Me.ColWorkUnit_Priority.Width = 63
        '
        'ColWorkUnit_Code
        '
        Me.ColWorkUnit_Code.HeaderText = "Work Unit Code"
        Me.ColWorkUnit_Code.Name = "ColWorkUnit_Code"
        Me.ColWorkUnit_Code.Width = 108
        '
        'ColWorkUnit_Desc
        '
        Me.ColWorkUnit_Desc.HeaderText = "Description"
        Me.ColWorkUnit_Desc.Name = "ColWorkUnit_Desc"
        Me.ColWorkUnit_Desc.ReadOnly = True
        Me.ColWorkUnit_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWorkUnit_Desc.Width = 66
        '
        'ColWorkUnit_Rate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWorkUnit_Rate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColWorkUnit_Rate.HeaderText = "Rate (Unit/Hour)"
        Me.ColWorkUnit_Rate.Name = "ColWorkUnit_Rate"
        Me.ColWorkUnit_Rate.Width = 111
        '
        'ColWorkUnit_Efficiency
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.NullValue = "0.00"
        Me.ColWorkUnit_Efficiency.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColWorkUnit_Efficiency.HeaderText = "Efficiency (%)"
        Me.ColWorkUnit_Efficiency.Name = "ColWorkUnit_Efficiency"
        Me.ColWorkUnit_Efficiency.Width = 95
        '
        'ColWorkUnit_EffectiveRate
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWorkUnit_EffectiveRate.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColWorkUnit_EffectiveRate.HeaderText = "Effective Rate"
        Me.ColWorkUnit_EffectiveRate.Name = "ColWorkUnit_EffectiveRate"
        Me.ColWorkUnit_EffectiveRate.ReadOnly = True
        '
        'ColWorkUnit_Remarks
        '
        Me.ColWorkUnit_Remarks.HeaderText = "Remarks"
        Me.ColWorkUnit_Remarks.MaxInputLength = 250
        Me.ColWorkUnit_Remarks.Name = "ColWorkUnit_Remarks"
        Me.ColWorkUnit_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWorkUnit_Remarks.Width = 55
        '
        'ColWorkUnit_PrimaryWorkUnit
        '
        Me.ColWorkUnit_PrimaryWorkUnit.HeaderText = "Primary Work Unit"
        Me.ColWorkUnit_PrimaryWorkUnit.Name = "ColWorkUnit_PrimaryWorkUnit"
        Me.ColWorkUnit_PrimaryWorkUnit.ReadOnly = True
        Me.ColWorkUnit_PrimaryWorkUnit.Visible = False
        Me.ColWorkUnit_PrimaryWorkUnit.Width = 117
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(624, 271)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Product Allocation"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(620, 267)
        Me.dgvItemProductGroup.TabIndex = 178
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 27
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 70
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 66
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 55
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(5, 12)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtProcessDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindProcess)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtProcessCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDocumentNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryProcess)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtDependency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label29)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboDependency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Size = New System.Drawing.Size(632, 67)
        Me.SplitContainer1.SplitterDistance = 313
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        '
        'txtProcessDesc
        '
        Me.txtProcessDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProcessDesc.Location = New System.Drawing.Point(77, 46)
        Me.txtProcessDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtProcessDesc.Name = "txtProcessDesc"
        Me.txtProcessDesc.ReadOnly = True
        Me.txtProcessDesc.Size = New System.Drawing.Size(233, 20)
        Me.txtProcessDesc.TabIndex = 6
        Me.txtProcessDesc.TabStop = False
        '
        'btnFindProcess
        '
        Me.btnFindProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProcess.Location = New System.Drawing.Point(265, 23)
        Me.btnFindProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindProcess.Name = "btnFindProcess"
        Me.btnFindProcess.Size = New System.Drawing.Size(43, 20)
        Me.btnFindProcess.TabIndex = 2
        Me.btnFindProcess.TabStop = False
        Me.btnFindProcess.Text = "Find"
        Me.btnFindProcess.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 25)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Process"
        '
        'txtProcessCode
        '
        Me.txtProcessCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProcessCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProcessCode.Location = New System.Drawing.Point(77, 23)
        Me.txtProcessCode.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtProcessCode.MaxLength = 250
        Me.txtProcessCode.Name = "txtProcessCode"
        Me.txtProcessCode.Size = New System.Drawing.Size(185, 20)
        Me.txtProcessCode.TabIndex = 1
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(205, 0)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(107, 20)
        Me.dtpDate.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 2)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Document No"
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReadOnly.Location = New System.Drawing.Point(77, 46)
        Me.txtReadOnly.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(31, 20)
        Me.txtReadOnly.TabIndex = 8
        Me.txtReadOnly.Visible = False
        '
        'TxtDocumentNo
        '
        Me.TxtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocumentNo.Location = New System.Drawing.Point(77, 0)
        Me.TxtDocumentNo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtDocumentNo.Name = "TxtDocumentNo"
        Me.TxtDocumentNo.ReadOnly = True
        Me.TxtDocumentNo.Size = New System.Drawing.Size(122, 20)
        Me.TxtDocumentNo.TabIndex = 0
        Me.TxtDocumentNo.TabStop = False
        '
        'txtPrimary
        '
        Me.txtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimary.Location = New System.Drawing.Point(77, -1)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtPrimary.MaxLength = 250
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(26, 20)
        Me.txtPrimary.TabIndex = 7
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'txtPrimaryProcess
        '
        Me.txtPrimaryProcess.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryProcess.Location = New System.Drawing.Point(77, 22)
        Me.txtPrimaryProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtPrimaryProcess.MaxLength = 250
        Me.txtPrimaryProcess.Name = "txtPrimaryProcess"
        Me.txtPrimaryProcess.Size = New System.Drawing.Size(26, 20)
        Me.txtPrimaryProcess.TabIndex = 9
        Me.txtPrimaryProcess.Text = "-1"
        Me.txtPrimaryProcess.Visible = False
        '
        'txtDependency
        '
        Me.txtDependency.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDependency.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDependency.Location = New System.Drawing.Point(237, 1)
        Me.txtDependency.MaxLength = 7
        Me.txtDependency.Name = "txtDependency"
        Me.txtDependency.Size = New System.Drawing.Size(77, 20)
        Me.txtDependency.TabIndex = 54
        Me.txtDependency.Text = "00.00"
        Me.txtDependency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(2, 2)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(68, 13)
        Me.Label29.TabIndex = 53
        Me.Label29.Text = "Dependency"
        '
        'cboDependency
        '
        Me.cboDependency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDependency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDependency.FormattingEnabled = True
        Me.cboDependency.Items.AddRange(New Object() {"End-Start", "Start-Start"})
        Me.cboDependency.Location = New System.Drawing.Point(74, 0)
        Me.cboDependency.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboDependency.Name = "cboDependency"
        Me.cboDependency.Size = New System.Drawing.Size(159, 21)
        Me.cboDependency.TabIndex = 52
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(74, 23)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(241, 41)
        Me.txtRemarks.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 25)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Remarks"
        '
        'btnApply
        '
        Me.btnApply.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnApply.Location = New System.Drawing.Point(428, 431)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(88, 26)
        Me.btnApply.TabIndex = 5
        Me.btnApply.Text = "&Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'BackgroundWorker2
        '
        '
        'BgwDelDgvWorkUnit
        '
        '
        'BackgroundWorker3
        '
        '
        'FrmMasterProductionRate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(670, 516)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnNew)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.GboDetail)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(680, 553)
        Me.Name = "FrmMasterProductionRate"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Rate Data"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.GboDetail.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvWorkUnit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents TslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents BtnOk As System.Windows.Forms.Button
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnNew As System.Windows.Forms.Button
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents GboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtProcessDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindProcess As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtProcessCode As System.Windows.Forms.TextBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents DgvWorkUnit As System.Windows.Forms.DataGridView
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryProcess As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwDelDgvWorkUnit As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColWorkUnit_Priority As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_Rate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_Efficiency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_EffectiveRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnit_PrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtDependency As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents cboDependency As System.Windows.Forms.ComboBox
End Class
