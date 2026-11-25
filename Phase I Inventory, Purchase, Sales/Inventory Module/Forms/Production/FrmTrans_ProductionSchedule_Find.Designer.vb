<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Find
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtMO = New System.Windows.Forms.TextBox
        Me.txtJO = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFinishGood = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindFG = New System.Windows.Forms.Button
        Me.btnFindJO = New System.Windows.Forms.Button
        Me.btnFindMO = New System.Windows.Forms.Button
        Me.btnFindWU = New System.Windows.Forms.Button
        Me.txtWorkunit = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnFindProcess = New System.Windows.Forms.Button
        Me.txtProcess = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindWIP = New System.Windows.Forms.Button
        Me.txtWIP = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnFindWO = New System.Windows.Forms.Button
        Me.txtWO = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colGoto = New System.Windows.Forms.DataGridViewButtonColumn
        Me.colMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinishGood = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWIP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWorkunit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPlanStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPlanStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPlanEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPlanEndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryWorkuunit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(288, 76)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(72, 21)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "MO No"
        '
        'txtMO
        '
        Me.txtMO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMO.Location = New System.Drawing.Point(76, 5)
        Me.txtMO.MaxLength = 150
        Me.txtMO.Name = "txtMO"
        Me.txtMO.Size = New System.Drawing.Size(237, 20)
        Me.txtMO.TabIndex = 0
        '
        'txtJO
        '
        Me.txtJO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtJO.Location = New System.Drawing.Point(76, 29)
        Me.txtJO.MaxLength = 150
        Me.txtJO.Name = "txtJO"
        Me.txtJO.Size = New System.Drawing.Size(237, 20)
        Me.txtJO.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "JO No"
        '
        'txtFinishGood
        '
        Me.txtFinishGood.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFinishGood.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFinishGood.Location = New System.Drawing.Point(76, 53)
        Me.txtFinishGood.MaxLength = 150
        Me.txtFinishGood.Name = "txtFinishGood"
        Me.txtFinishGood.Size = New System.Drawing.Size(237, 20)
        Me.txtFinishGood.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Finish Good"
        '
        'btnFindFG
        '
        Me.btnFindFG.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindFG.Location = New System.Drawing.Point(319, 52)
        Me.btnFindFG.Name = "btnFindFG"
        Me.btnFindFG.Size = New System.Drawing.Size(46, 21)
        Me.btnFindFG.TabIndex = 5
        Me.btnFindFG.Text = "..."
        Me.btnFindFG.UseVisualStyleBackColor = True
        '
        'btnFindJO
        '
        Me.btnFindJO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindJO.Location = New System.Drawing.Point(319, 28)
        Me.btnFindJO.Name = "btnFindJO"
        Me.btnFindJO.Size = New System.Drawing.Size(46, 21)
        Me.btnFindJO.TabIndex = 3
        Me.btnFindJO.Text = "..."
        Me.btnFindJO.UseVisualStyleBackColor = True
        '
        'btnFindMO
        '
        Me.btnFindMO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindMO.Location = New System.Drawing.Point(319, 4)
        Me.btnFindMO.Name = "btnFindMO"
        Me.btnFindMO.Size = New System.Drawing.Size(46, 21)
        Me.btnFindMO.TabIndex = 1
        Me.btnFindMO.Text = "..."
        Me.btnFindMO.UseVisualStyleBackColor = True
        '
        'btnFindWU
        '
        Me.btnFindWU.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindWU.Location = New System.Drawing.Point(314, 52)
        Me.btnFindWU.Name = "btnFindWU"
        Me.btnFindWU.Size = New System.Drawing.Size(46, 21)
        Me.btnFindWU.TabIndex = 5
        Me.btnFindWU.Text = "..."
        Me.btnFindWU.UseVisualStyleBackColor = True
        '
        'txtWorkunit
        '
        Me.txtWorkunit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWorkunit.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtWorkunit.Location = New System.Drawing.Point(69, 53)
        Me.txtWorkunit.MaxLength = 150
        Me.txtWorkunit.Name = "txtWorkunit"
        Me.txtWorkunit.Size = New System.Drawing.Size(239, 20)
        Me.txtWorkunit.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Work Unit"
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.Location = New System.Drawing.Point(210, 76)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(72, 21)
        Me.btnFind.TabIndex = 6
        Me.btnFind.Text = "&Find All"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnFindProcess
        '
        Me.btnFindProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProcess.Location = New System.Drawing.Point(319, 76)
        Me.btnFindProcess.Name = "btnFindProcess"
        Me.btnFindProcess.Size = New System.Drawing.Size(46, 21)
        Me.btnFindProcess.TabIndex = 7
        Me.btnFindProcess.Text = "..."
        Me.btnFindProcess.UseVisualStyleBackColor = True
        '
        'txtProcess
        '
        Me.txtProcess.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProcess.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProcess.Location = New System.Drawing.Point(76, 77)
        Me.txtProcess.MaxLength = 150
        Me.txtProcess.Name = "txtProcess"
        Me.txtProcess.Size = New System.Drawing.Size(237, 20)
        Me.txtProcess.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Process"
        '
        'btnFindWIP
        '
        Me.btnFindWIP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindWIP.Location = New System.Drawing.Point(314, 28)
        Me.btnFindWIP.Name = "btnFindWIP"
        Me.btnFindWIP.Size = New System.Drawing.Size(46, 21)
        Me.btnFindWIP.TabIndex = 3
        Me.btnFindWIP.Text = "..."
        Me.btnFindWIP.UseVisualStyleBackColor = True
        '
        'txtWIP
        '
        Me.txtWIP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWIP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtWIP.Location = New System.Drawing.Point(69, 29)
        Me.txtWIP.MaxLength = 150
        Me.txtWIP.Name = "txtWIP"
        Me.txtWIP.Size = New System.Drawing.Size(239, 20)
        Me.txtWIP.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(28, 13)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "WIP"
        '
        'btnFindWO
        '
        Me.btnFindWO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindWO.Location = New System.Drawing.Point(314, 4)
        Me.btnFindWO.Name = "btnFindWO"
        Me.btnFindWO.Size = New System.Drawing.Size(46, 21)
        Me.btnFindWO.TabIndex = 1
        Me.btnFindWO.Text = "..."
        Me.btnFindWO.UseVisualStyleBackColor = True
        '
        'txtWO
        '
        Me.txtWO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtWO.Location = New System.Drawing.Point(69, 5)
        Me.txtWO.MaxLength = 150
        Me.txtWO.Name = "txtWO"
        Me.txtWO.Size = New System.Drawing.Size(239, 20)
        Me.txtWO.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "WO No"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 6)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtProcess)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtFinishGood)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtJO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtMO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindMO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindJO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindFG)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindProcess)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtWO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFind)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindWO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindWIP)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindWU)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtWorkunit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtWIP)
        Me.SplitContainer1.Size = New System.Drawing.Size(743, 100)
        Me.SplitContainer1.SplitterDistance = 368
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvItem.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colGoto, Me.colMO, Me.colJO, Me.ColFinishGood, Me.colProcess, Me.colWO, Me.colWIP, Me.colWorkunit, Me.colPlanStartDate, Me.colPlanStartTime, Me.colPlanEndDate, Me.colPlanEndTime, Me.colPrimaryMO, Me.colIdentifierJO, Me.colIdentifierWO, Me.colPrimaryProcess, Me.colPrimaryWorkuunit})
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgvItem.DefaultCellStyle = DataGridViewCellStyle4
        Me.DgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvItem.Location = New System.Drawing.Point(3, 16)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvItem.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvItem.Size = New System.Drawing.Size(734, 247)
        Me.DgvItem.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DgvItem)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 105)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(740, 266)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Result"
        '
        'colNo
        '
        Me.colNo.Frozen = True
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 46
        '
        'colGoto
        '
        Me.colGoto.Frozen = True
        Me.colGoto.HeaderText = "GoTo"
        Me.colGoto.Name = "colGoto"
        Me.colGoto.Width = 40
        '
        'colMO
        '
        Me.colMO.HeaderText = "MO"
        Me.colMO.Name = "colMO"
        Me.colMO.ReadOnly = True
        Me.colMO.Width = 49
        '
        'colJO
        '
        Me.colJO.HeaderText = "JO"
        Me.colJO.Name = "colJO"
        Me.colJO.ReadOnly = True
        Me.colJO.Width = 45
        '
        'ColFinishGood
        '
        Me.ColFinishGood.HeaderText = "Finish Good"
        Me.ColFinishGood.Name = "ColFinishGood"
        Me.ColFinishGood.ReadOnly = True
        Me.ColFinishGood.Width = 88
        '
        'colProcess
        '
        Me.colProcess.HeaderText = "Process"
        Me.colProcess.Name = "colProcess"
        Me.colProcess.ReadOnly = True
        Me.colProcess.Width = 70
        '
        'colWO
        '
        Me.colWO.HeaderText = "WO"
        Me.colWO.Name = "colWO"
        Me.colWO.ReadOnly = True
        Me.colWO.Width = 51
        '
        'colWIP
        '
        Me.colWIP.HeaderText = "WIP"
        Me.colWIP.Name = "colWIP"
        Me.colWIP.ReadOnly = True
        Me.colWIP.Width = 53
        '
        'colWorkunit
        '
        Me.colWorkunit.HeaderText = "Workunit"
        Me.colWorkunit.Name = "colWorkunit"
        Me.colWorkunit.ReadOnly = True
        Me.colWorkunit.Width = 75
        '
        'colPlanStartDate
        '
        Me.colPlanStartDate.HeaderText = "Start"
        Me.colPlanStartDate.Name = "colPlanStartDate"
        Me.colPlanStartDate.ReadOnly = True
        Me.colPlanStartDate.Width = 54
        '
        'colPlanStartTime
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPlanStartTime.DefaultCellStyle = DataGridViewCellStyle2
        Me.colPlanStartTime.HeaderText = ""
        Me.colPlanStartTime.Name = "colPlanStartTime"
        Me.colPlanStartTime.ReadOnly = True
        Me.colPlanStartTime.Width = 19
        '
        'colPlanEndDate
        '
        Me.colPlanEndDate.HeaderText = "End"
        Me.colPlanEndDate.Name = "colPlanEndDate"
        Me.colPlanEndDate.ReadOnly = True
        Me.colPlanEndDate.Width = 51
        '
        'colPlanEndTime
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPlanEndTime.DefaultCellStyle = DataGridViewCellStyle3
        Me.colPlanEndTime.HeaderText = ""
        Me.colPlanEndTime.Name = "colPlanEndTime"
        Me.colPlanEndTime.ReadOnly = True
        Me.colPlanEndTime.Width = 19
        '
        'colPrimaryMO
        '
        Me.colPrimaryMO.HeaderText = "Primary MO"
        Me.colPrimaryMO.Name = "colPrimaryMO"
        Me.colPrimaryMO.ReadOnly = True
        Me.colPrimaryMO.Visible = False
        Me.colPrimaryMO.Width = 86
        '
        'colIdentifierJO
        '
        Me.colIdentifierJO.HeaderText = "Identifier JO"
        Me.colIdentifierJO.Name = "colIdentifierJO"
        Me.colIdentifierJO.ReadOnly = True
        Me.colIdentifierJO.Visible = False
        Me.colIdentifierJO.Width = 88
        '
        'colIdentifierWO
        '
        Me.colIdentifierWO.HeaderText = "Identifier WO"
        Me.colIdentifierWO.Name = "colIdentifierWO"
        Me.colIdentifierWO.ReadOnly = True
        Me.colIdentifierWO.Visible = False
        Me.colIdentifierWO.Width = 94
        '
        'colPrimaryProcess
        '
        Me.colPrimaryProcess.HeaderText = "Primary Process"
        Me.colPrimaryProcess.Name = "colPrimaryProcess"
        Me.colPrimaryProcess.ReadOnly = True
        Me.colPrimaryProcess.Visible = False
        Me.colPrimaryProcess.Width = 107
        '
        'colPrimaryWorkuunit
        '
        Me.colPrimaryWorkuunit.HeaderText = "Primary Workunit"
        Me.colPrimaryWorkuunit.Name = "colPrimaryWorkuunit"
        Me.colPrimaryWorkuunit.ReadOnly = True
        Me.colPrimaryWorkuunit.Visible = False
        Me.colPrimaryWorkuunit.Width = 112
        '
        'FrmTrans_ProductionSchedule_Find
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(749, 379)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(649, 343)
        Me.Name = "FrmTrans_ProductionSchedule_Find"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Data"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtMO As System.Windows.Forms.TextBox
    Friend WithEvents txtJO As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFinishGood As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFindFG As System.Windows.Forms.Button
    Friend WithEvents btnFindJO As System.Windows.Forms.Button
    Friend WithEvents btnFindMO As System.Windows.Forms.Button
    Friend WithEvents btnFindWU As System.Windows.Forms.Button
    Friend WithEvents txtWorkunit As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnFindProcess As System.Windows.Forms.Button
    Friend WithEvents txtProcess As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindWIP As System.Windows.Forms.Button
    Friend WithEvents txtWIP As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFindWO As System.Windows.Forms.Button
    Friend WithEvents txtWO As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colGoto As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents colMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinishGood As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWIP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWorkunit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPlanStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPlanStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPlanEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPlanEndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryWorkuunit As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
