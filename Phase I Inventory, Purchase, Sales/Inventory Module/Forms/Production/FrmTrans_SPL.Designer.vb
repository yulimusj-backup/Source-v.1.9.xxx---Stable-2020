<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.dgvWorkshift = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colShiftCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEarlyOT = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colEarlyOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEarlyDelayDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEarlyBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDelayBeforeEarlyBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colOT = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDelayDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDelayBeforeBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryWorkshift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LlbWorkUnit = New System.Windows.Forms.LinkLabel
        Me.dgvWorkUnit = New System.Windows.Forms.DataGridView
        Me.colWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtJobDesc = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.txtDocumentNo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.txtReadonly = New System.Windows.Forms.TextBox
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllWokUnitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyToWorkUnitGroupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.dgvWorkshift, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvWorkUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(633, 567)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(107, 32)
        Me.btnPrint.TabIndex = 19
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(749, 567)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 20
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(404, 567)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 17
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(173, 567)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 15
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(288, 567)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 16
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(519, 567)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 18
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(11, 2)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(1009, 558)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.SplitContainer2)
        Me.gboItem.Location = New System.Drawing.Point(7, 101)
        Me.gboItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboItem.Size = New System.Drawing.Size(993, 411)
        Me.gboItem.TabIndex = 1
        Me.gboItem.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(4, 19)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.dgvWorkshift)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpFrom)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpTo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label5)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer2.Size = New System.Drawing.Size(985, 388)
        Me.SplitContainer2.SplitterDistance = 219
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 19
        '
        'dgvWorkshift
        '
        Me.dgvWorkshift.AllowUserToAddRows = False
        Me.dgvWorkshift.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWorkshift.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWorkshift.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWorkshift.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWorkshift.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colShiftCode, Me.colDesc, Me.colEarlyOT, Me.colEarlyOTDuration, Me.colEarlyDelayDuration, Me.ColEarlyBreak, Me.ColDelayBeforeEarlyBreak, Me.colOT, Me.colOTDuration, Me.colDelayDuration, Me.ColBreak, Me.ColDelayBeforeBreak, Me.colRemarks, Me.colPrimaryWorkshift})
        Me.dgvWorkshift.Location = New System.Drawing.Point(5, 36)
        Me.dgvWorkshift.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvWorkshift.Name = "dgvWorkshift"
        Me.dgvWorkshift.Size = New System.Drawing.Size(975, 171)
        Me.dgvWorkshift.TabIndex = 32
        '
        'colNo
        '
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 51
        '
        'colShiftCode
        '
        Me.colShiftCode.HeaderText = "Shift"
        Me.colShiftCode.Name = "colShiftCode"
        Me.colShiftCode.Width = 61
        '
        'colDesc
        '
        Me.colDesc.HeaderText = "Desc"
        Me.colDesc.Name = "colDesc"
        Me.colDesc.Width = 65
        '
        'colEarlyOT
        '
        Me.colEarlyOT.HeaderText = "Early OT"
        Me.colEarlyOT.Name = "colEarlyOT"
        Me.colEarlyOT.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colEarlyOT.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colEarlyOT.Width = 82
        '
        'colEarlyOTDuration
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colEarlyOTDuration.DefaultCellStyle = DataGridViewCellStyle1
        Me.colEarlyOTDuration.HeaderText = "Early Duration"
        Me.colEarlyOTDuration.Name = "colEarlyOTDuration"
        Me.colEarlyOTDuration.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colEarlyOTDuration.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colEarlyOTDuration.Width = 94
        '
        'colEarlyDelayDuration
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colEarlyDelayDuration.DefaultCellStyle = DataGridViewCellStyle2
        Me.colEarlyDelayDuration.HeaderText = "Early Delay"
        Me.colEarlyDelayDuration.Name = "colEarlyDelayDuration"
        Me.colEarlyDelayDuration.Width = 97
        '
        'ColEarlyBreak
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColEarlyBreak.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColEarlyBreak.HeaderText = "Early Break"
        Me.ColEarlyBreak.Name = "ColEarlyBreak"
        Me.ColEarlyBreak.Width = 97
        '
        'ColDelayBeforeEarlyBreak
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDelayBeforeEarlyBreak.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColDelayBeforeEarlyBreak.HeaderText = "Early Delay Before Break"
        Me.ColDelayBeforeEarlyBreak.Name = "ColDelayBeforeEarlyBreak"
        Me.ColDelayBeforeEarlyBreak.Width = 142
        '
        'colOT
        '
        Me.colOT.HeaderText = "OT"
        Me.colOT.Name = "colOT"
        Me.colOT.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colOT.Width = 34
        '
        'colOTDuration
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colOTDuration.DefaultCellStyle = DataGridViewCellStyle5
        Me.colOTDuration.HeaderText = "OT Duration"
        Me.colOTDuration.Name = "colOTDuration"
        Me.colOTDuration.Width = 102
        '
        'colDelayDuration
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDelayDuration.DefaultCellStyle = DataGridViewCellStyle6
        Me.colDelayDuration.HeaderText = "Delay"
        Me.colDelayDuration.Name = "colDelayDuration"
        Me.colDelayDuration.Width = 69
        '
        'ColBreak
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBreak.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColBreak.HeaderText = "Break"
        Me.ColBreak.Name = "ColBreak"
        Me.ColBreak.Width = 70
        '
        'ColDelayBeforeBreak
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDelayBeforeBreak.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDelayBeforeBreak.HeaderText = "Delay Before Break"
        Me.ColDelayBeforeBreak.Name = "ColDelayBeforeBreak"
        Me.ColDelayBeforeBreak.Width = 142
        '
        'colRemarks
        '
        Me.colRemarks.HeaderText = "Remarks"
        Me.colRemarks.Name = "colRemarks"
        Me.colRemarks.Width = 89
        '
        'colPrimaryWorkshift
        '
        Me.colPrimaryWorkshift.HeaderText = "Primary Workshift"
        Me.colPrimaryWorkshift.Name = "colPrimaryWorkshift"
        Me.colPrimaryWorkshift.ReadOnly = True
        Me.colPrimaryWorkshift.Visible = False
        Me.colPrimaryWorkshift.Width = 105
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(83, 4)
        Me.DtpFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(175, 22)
        Me.DtpFrom.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 9)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 17)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "From"
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(355, 4)
        Me.DtpTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(171, 22)
        Me.DtpTo.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(297, 9)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(25, 17)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "To"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LlbWorkUnit)
        Me.GroupBox1.Controls.Add(Me.dgvWorkUnit)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(985, 164)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Apply To"
        '
        'LlbWorkUnit
        '
        Me.LlbWorkUnit.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbWorkUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbWorkUnit.AutoSize = True
        Me.LlbWorkUnit.Location = New System.Drawing.Point(919, 0)
        Me.LlbWorkUnit.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LlbWorkUnit.Name = "LlbWorkUnit"
        Me.LlbWorkUnit.Size = New System.Drawing.Size(63, 17)
        Me.LlbWorkUnit.TabIndex = 1
        Me.LlbWorkUnit.TabStop = True
        Me.LlbWorkUnit.Text = "Advance"
        Me.LlbWorkUnit.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'dgvWorkUnit
        '
        Me.dgvWorkUnit.AllowUserToAddRows = False
        Me.dgvWorkUnit.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWorkUnit.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWorkUnit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWorkUnit.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWU_No, Me.colWU_Code, Me.colWU_Desc, Me.colWU_Primary})
        Me.dgvWorkUnit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvWorkUnit.Location = New System.Drawing.Point(4, 19)
        Me.dgvWorkUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvWorkUnit.Name = "dgvWorkUnit"
        Me.dgvWorkUnit.Size = New System.Drawing.Size(977, 141)
        Me.dgvWorkUnit.TabIndex = 0
        '
        'colWU_No
        '
        Me.colWU_No.Frozen = True
        Me.colWU_No.HeaderText = "No"
        Me.colWU_No.Name = "colWU_No"
        Me.colWU_No.ReadOnly = True
        Me.colWU_No.Width = 51
        '
        'colWU_Code
        '
        Me.colWU_Code.Frozen = True
        Me.colWU_Code.HeaderText = "Work Unit"
        Me.colWU_Code.MaxInputLength = 250
        Me.colWU_Code.Name = "colWU_Code"
        Me.colWU_Code.Width = 95
        '
        'colWU_Desc
        '
        Me.colWU_Desc.HeaderText = "Desc"
        Me.colWU_Desc.Name = "colWU_Desc"
        Me.colWU_Desc.ReadOnly = True
        Me.colWU_Desc.Width = 65
        '
        'colWU_Primary
        '
        Me.colWU_Primary.HeaderText = "Primary"
        Me.colWU_Primary.Name = "colWU_Primary"
        Me.colWU_Primary.ReadOnly = True
        Me.colWU_Primary.Visible = False
        Me.colWU_Primary.Width = 66
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 12)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtJobDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDocumentNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Size = New System.Drawing.Size(993, 102)
        Me.SplitContainer1.SplitterDistance = 531
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 99999999
        Me.SplitContainer1.TabStop = False
        '
        'txtJobDesc
        '
        Me.txtJobDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJobDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtJobDesc.Location = New System.Drawing.Point(132, 64)
        Me.txtJobDesc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtJobDesc.MaxLength = 250
        Me.txtJobDesc.Name = "txtJobDesc"
        Me.txtJobDesc.Size = New System.Drawing.Size(394, 22)
        Me.txtJobDesc.TabIndex = 20
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 69)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 17)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Reason"
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(85, 0)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(37, 22)
        Me.txtPrimary.TabIndex = 19
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 37)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 17)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Date"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(132, 32)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(394, 22)
        Me.dtpDate.TabIndex = 1
        '
        'txtDocumentNo
        '
        Me.txtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDocumentNo.BackColor = System.Drawing.SystemColors.Control
        Me.txtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDocumentNo.Location = New System.Drawing.Point(132, 0)
        Me.txtDocumentNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDocumentNo.MaxLength = 250
        Me.txtDocumentNo.Name = "txtDocumentNo"
        Me.txtDocumentNo.ReadOnly = True
        Me.txtDocumentNo.Size = New System.Drawing.Size(394, 22)
        Me.txtDocumentNo.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Document No"
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(77, 0)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(368, 88)
        Me.txtRemarks.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 5)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Remarks"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(508, 519)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(135, 32)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(365, 519)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(135, 32)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslModifiedDate, Me.tslCreatedDate, Me.tslModifiedBy})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 611)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1041, 25)
        Me.StatusStrip1.TabIndex = 22
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(255, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(255, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(255, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(255, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'txtReadonly
        '
        Me.txtReadonly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReadonly.Location = New System.Drawing.Point(981, 578)
        Me.txtReadonly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtReadonly.Name = "txtReadonly"
        Me.txtReadonly.ReadOnly = True
        Me.txtReadonly.Size = New System.Drawing.Size(37, 22)
        Me.txtReadonly.TabIndex = 20
        Me.txtReadonly.Visible = False
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllWokUnitToolStripMenuItem, Me.ApplyToWorkUnitGroupToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(251, 52)
        '
        'ApplyToAllWokUnitToolStripMenuItem
        '
        Me.ApplyToAllWokUnitToolStripMenuItem.Name = "ApplyToAllWokUnitToolStripMenuItem"
        Me.ApplyToAllWokUnitToolStripMenuItem.Size = New System.Drawing.Size(250, 24)
        Me.ApplyToAllWokUnitToolStripMenuItem.Text = "Apply to All Wok Unit"
        '
        'ApplyToWorkUnitGroupToolStripMenuItem
        '
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Name = "ApplyToWorkUnitGroupToolStripMenuItem"
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Size = New System.Drawing.Size(250, 24)
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Text = "Apply to Work Unit Group"
        '
        'BackgroundWorker1
        '
        '
        'FrmTrans_SPL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1041, 636)
        Me.Controls.Add(Me.txtReadonly)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(1049, 673)
        Me.Name = "FrmTrans_SPL"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Overtime Order"
        Me.gboDetail.ResumeLayout(False)
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.dgvWorkshift, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvWorkUnit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    'Friend WithEvents TGV_WO As AdvancedDataGridView.TreeGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents txtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtReadonly As System.Windows.Forms.TextBox
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvWorkUnit As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvWorkshift As System.Windows.Forms.DataGridView
    Friend WithEvents LlbWorkUnit As System.Windows.Forms.LinkLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyToAllWokUnitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToWorkUnitGroupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtJobDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents colWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colShiftCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEarlyOT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colEarlyOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEarlyDelayDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEarlyBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDelayBeforeEarlyBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colOT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDelayDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDelayBeforeBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryWorkshift As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
