<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMasterProduct_ChangeControl
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.chkModifyLocation = New System.Windows.Forms.CheckBox
        Me.chkModifyExpiry = New System.Windows.Forms.CheckBox
        Me.chkModifyBatch = New System.Windows.Forms.CheckBox
        Me.TxtLocationDefaultDesc = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLocationDefault = New System.Windows.Forms.TextBox
        Me.btnFindLocation = New System.Windows.Forms.Button
        Me.TxtLocationDefault = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.chkControlLocation = New System.Windows.Forms.CheckBox
        Me.dtpExpiryDefault = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkControlExpiry = New System.Windows.Forms.CheckBox
        Me.TxtBatchNoDefault = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkControlBatchNo = New System.Windows.Forms.CheckBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.LblProgress = New System.Windows.Forms.Label
        Me.Prg1 = New System.Windows.Forms.ProgressBar
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.TimerClose = New System.Windows.Forms.Timer(Me.components)
        Me.TimerNote = New System.Windows.Forms.Timer(Me.components)
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbCheckAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbUncheckAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvProduct)
        Me.SplitContainer1.Size = New System.Drawing.Size(982, 553)
        Me.SplitContainer1.SplitterDistance = 194
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkModifyLocation)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkModifyExpiry)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkModifyBatch)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtLocationDefaultDesc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryLocationDefault)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnFindLocation)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtLocationDefault)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkControlLocation)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpExpiryDefault)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkControlExpiry)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtBatchNoDefault)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkControlBatchNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnStart)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer2.Size = New System.Drawing.Size(982, 194)
        Me.SplitContainer2.SplitterDistance = 495
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'chkModifyLocation
        '
        Me.chkModifyLocation.AutoSize = True
        Me.chkModifyLocation.Location = New System.Drawing.Point(12, 111)
        Me.chkModifyLocation.Name = "chkModifyLocation"
        Me.chkModifyLocation.Size = New System.Drawing.Size(181, 21)
        Me.chkModifyLocation.TabIndex = 6
        Me.chkModifyLocation.Text = "Modifiy Location Control"
        Me.chkModifyLocation.UseVisualStyleBackColor = True
        '
        'chkModifyExpiry
        '
        Me.chkModifyExpiry.AutoSize = True
        Me.chkModifyExpiry.Location = New System.Drawing.Point(12, 57)
        Me.chkModifyExpiry.Name = "chkModifyExpiry"
        Me.chkModifyExpiry.Size = New System.Drawing.Size(199, 21)
        Me.chkModifyExpiry.TabIndex = 3
        Me.chkModifyExpiry.Text = "Modifiy Expiry Date Control"
        Me.chkModifyExpiry.UseVisualStyleBackColor = True
        '
        'chkModifyBatch
        '
        Me.chkModifyBatch.AutoSize = True
        Me.chkModifyBatch.Location = New System.Drawing.Point(12, 3)
        Me.chkModifyBatch.Name = "chkModifyBatch"
        Me.chkModifyBatch.Size = New System.Drawing.Size(185, 21)
        Me.chkModifyBatch.TabIndex = 0
        Me.chkModifyBatch.Text = "Modifiy Batch No Control"
        Me.chkModifyBatch.UseVisualStyleBackColor = True
        '
        'TxtLocationDefaultDesc
        '
        Me.TxtLocationDefaultDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocationDefaultDesc.Location = New System.Drawing.Point(276, 168)
        Me.TxtLocationDefaultDesc.Name = "TxtLocationDefaultDesc"
        Me.TxtLocationDefaultDesc.ReadOnly = True
        Me.TxtLocationDefaultDesc.Size = New System.Drawing.Size(216, 22)
        Me.TxtLocationDefaultDesc.TabIndex = 13
        Me.TxtLocationDefaultDesc.TabStop = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(174, 168)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(42, 22)
        Me.TxtReadOnly.TabIndex = 12
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtPrimaryLocationDefault
        '
        Me.TxtPrimaryLocationDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryLocationDefault.Location = New System.Drawing.Point(228, 141)
        Me.TxtPrimaryLocationDefault.Name = "TxtPrimaryLocationDefault"
        Me.TxtPrimaryLocationDefault.Size = New System.Drawing.Size(42, 22)
        Me.TxtPrimaryLocationDefault.TabIndex = 12
        Me.TxtPrimaryLocationDefault.TabStop = False
        Me.TxtPrimaryLocationDefault.Text = "-1"
        Me.TxtPrimaryLocationDefault.Visible = False
        '
        'btnFindLocation
        '
        Me.btnFindLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLocation.Location = New System.Drawing.Point(425, 139)
        Me.btnFindLocation.Name = "btnFindLocation"
        Me.btnFindLocation.Size = New System.Drawing.Size(67, 23)
        Me.btnFindLocation.TabIndex = 9
        Me.btnFindLocation.Text = "Find"
        Me.btnFindLocation.UseVisualStyleBackColor = True
        '
        'TxtLocationDefault
        '
        Me.TxtLocationDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocationDefault.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLocationDefault.Location = New System.Drawing.Point(276, 140)
        Me.TxtLocationDefault.Name = "TxtLocationDefault"
        Me.TxtLocationDefault.Size = New System.Drawing.Size(143, 22)
        Me.TxtLocationDefault.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(169, 141)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(101, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Default Value: "
        '
        'chkControlLocation
        '
        Me.chkControlLocation.AutoSize = True
        Me.chkControlLocation.Location = New System.Drawing.Point(12, 140)
        Me.chkControlLocation.Name = "chkControlLocation"
        Me.chkControlLocation.Size = New System.Drawing.Size(133, 21)
        Me.chkControlLocation.TabIndex = 7
        Me.chkControlLocation.Text = "Control Location"
        Me.chkControlLocation.UseVisualStyleBackColor = True
        '
        'dtpExpiryDefault
        '
        Me.dtpExpiryDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpExpiryDefault.CustomFormat = "dd MMM yyyy"
        Me.dtpExpiryDefault.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpExpiryDefault.Location = New System.Drawing.Point(276, 83)
        Me.dtpExpiryDefault.Name = "dtpExpiryDefault"
        Me.dtpExpiryDefault.Size = New System.Drawing.Size(216, 22)
        Me.dtpExpiryDefault.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(169, 85)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Default Value: "
        '
        'chkControlExpiry
        '
        Me.chkControlExpiry.AutoSize = True
        Me.chkControlExpiry.Location = New System.Drawing.Point(12, 84)
        Me.chkControlExpiry.Name = "chkControlExpiry"
        Me.chkControlExpiry.Size = New System.Drawing.Size(151, 21)
        Me.chkControlExpiry.TabIndex = 4
        Me.chkControlExpiry.Text = "Control Expiry Date"
        Me.chkControlExpiry.UseVisualStyleBackColor = True
        '
        'TxtBatchNoDefault
        '
        Me.TxtBatchNoDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBatchNoDefault.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBatchNoDefault.Location = New System.Drawing.Point(276, 28)
        Me.TxtBatchNoDefault.Name = "TxtBatchNoDefault"
        Me.TxtBatchNoDefault.Size = New System.Drawing.Size(216, 22)
        Me.TxtBatchNoDefault.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(169, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Default Value: "
        '
        'chkControlBatchNo
        '
        Me.chkControlBatchNo.AutoSize = True
        Me.chkControlBatchNo.Location = New System.Drawing.Point(12, 30)
        Me.chkControlBatchNo.Name = "chkControlBatchNo"
        Me.chkControlBatchNo.Size = New System.Drawing.Size(137, 21)
        Me.chkControlBatchNo.TabIndex = 1
        Me.chkControlBatchNo.Text = "Control Batch No"
        Me.chkControlBatchNo.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(93, 168)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 11
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(12, 168)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 10
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvItemProductGroup)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(483, 194)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Scope"
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(3, 18)
        Me.dgvItemProductGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(477, 173)
        Me.dgvItemProductGroup.TabIndex = 0
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.Width = 51
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 48
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.HeaderText = "Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.Width = 66
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        Me.ColProductGroup_Desc.Width = 104
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'Panel1
        '
        Me.Panel1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Panel1.Controls.Add(Me.LblProgress)
        Me.Panel1.Controls.Add(Me.Prg1)
        Me.Panel1.Location = New System.Drawing.Point(135, 147)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(712, 61)
        Me.Panel1.TabIndex = 3
        Me.Panel1.Visible = False
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(3, 29)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 1
        Me.LblProgress.Text = "Label4"
        '
        'Prg1
        '
        Me.Prg1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Prg1.Location = New System.Drawing.Point(3, 3)
        Me.Prg1.Name = "Prg1"
        Me.Prg1.Size = New System.Drawing.Size(706, 23)
        Me.Prg1.TabIndex = 0
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(816, 329)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCheckAll.TabIndex = 1
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(888, 329)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncheckAll.TabIndex = 2
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'dgvProduct
        '
        Me.dgvProduct.AllowUserToAddRows = False
        Me.dgvProduct.AllowUserToDeleteRows = False
        Me.dgvProduct.AllowUserToResizeColumns = False
        Me.dgvProduct.AllowUserToResizeRows = False
        Me.dgvProduct.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProduct.ColumnHeadersHeight = 25
        Me.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColBatchNo, Me.ColExpiry, Me.ColLocation, Me.ColModify, Me.ColPrimaryProduct})
        Me.dgvProduct.Location = New System.Drawing.Point(12, 3)
        Me.dgvProduct.Name = "dgvProduct"
        Me.dgvProduct.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvProduct.RowTemplate.Height = 24
        Me.dgvProduct.Size = New System.Drawing.Size(958, 323)
        Me.dgvProduct.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 88
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.Width = 104
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Control Batch"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.Width = 99
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Control Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.Width = 101
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Control Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 117
        '
        'ColModify
        '
        Me.ColModify.HeaderText = "Modify"
        Me.ColModify.Name = "ColModify"
        Me.ColModify.Width = 55
        '
        'ColPrimaryProduct
        '
        Me.ColPrimaryProduct.HeaderText = "Primary"
        Me.ColPrimaryProduct.Name = "ColPrimaryProduct"
        Me.ColPrimaryProduct.Visible = False
        Me.ColPrimaryProduct.Width = 81
        '
        'BackgroundWorker3
        '
        '
        'bgw1
        '
        Me.bgw1.WorkerReportsProgress = True
        '
        'TimerClose
        '
        Me.TimerClose.Interval = 500
        '
        'TimerNote
        '
        Me.TimerNote.Interval = 500
        '
        'frmMasterProduct_ChangeControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(982, 553)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "frmMasterProduct_ChangeControl"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change Control State"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryLocationDefault As System.Windows.Forms.TextBox
    Friend WithEvents btnFindLocation As System.Windows.Forms.Button
    Friend WithEvents TxtLocationDefault As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkControlLocation As System.Windows.Forms.CheckBox
    Friend WithEvents dtpExpiryDefault As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkControlExpiry As System.Windows.Forms.CheckBox
    Friend WithEvents TxtBatchNoDefault As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkControlBatchNo As System.Windows.Forms.CheckBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtLocationDefaultDesc As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents Prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents chkModifyLocation As System.Windows.Forms.CheckBox
    Friend WithEvents chkModifyExpiry As System.Windows.Forms.CheckBox
    Friend WithEvents chkModifyBatch As System.Windows.Forms.CheckBox
    Friend WithEvents TimerClose As System.Windows.Forms.Timer
    Friend WithEvents TimerNote As System.Windows.Forms.Timer
End Class
