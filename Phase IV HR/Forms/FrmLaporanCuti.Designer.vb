<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanCuti
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ChkJobPosition = New System.Windows.Forms.CheckBox
        Me.Lbl_JP_UnChecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_JP_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbJobPosition = New System.Windows.Forms.CheckedListBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.ChkDepartemen = New System.Windows.Forms.CheckBox
        Me.Lbl_Dept_UnChecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Dept_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbDepartemen = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkLeaveData = New System.Windows.Forms.CheckBox
        Me.Lbl_Leave_UnChecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Leave_Checked = New System.Windows.Forms.LinkLabel
        Me.clbLeave = New System.Windows.Forms.CheckedListBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkEmployee = New System.Windows.Forms.CheckBox
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.CboFilterEmployee = New System.Windows.Forms.ComboBox
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(465, 98)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 102
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(537, 6)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(43, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(10, 37)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer1)
        Me.SplitContainer3.Size = New System.Drawing.Size(239, 474)
        Me.SplitContainer3.SplitterDistance = 232
        Me.SplitContainer3.TabIndex = 2
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox4)
        Me.SplitContainer2.Size = New System.Drawing.Size(239, 232)
        Me.SplitContainer2.SplitterDistance = 113
        Me.SplitContainer2.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ChkJobPosition)
        Me.GroupBox3.Controls.Add(Me.Lbl_JP_UnChecked)
        Me.GroupBox3.Controls.Add(Me.Lbl_JP_Checked)
        Me.GroupBox3.Controls.Add(Me.ClbJobPosition)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(239, 113)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'ChkJobPosition
        '
        Me.ChkJobPosition.AutoSize = True
        Me.ChkJobPosition.Location = New System.Drawing.Point(5, -1)
        Me.ChkJobPosition.Name = "ChkJobPosition"
        Me.ChkJobPosition.Size = New System.Drawing.Size(83, 17)
        Me.ChkJobPosition.TabIndex = 0
        Me.ChkJobPosition.Text = "Job Position"
        Me.ChkJobPosition.UseVisualStyleBackColor = True
        '
        'Lbl_JP_UnChecked
        '
        Me.Lbl_JP_UnChecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JP_UnChecked.AutoSize = True
        Me.Lbl_JP_UnChecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JP_UnChecked.Location = New System.Drawing.Point(168, 0)
        Me.Lbl_JP_UnChecked.Name = "Lbl_JP_UnChecked"
        Me.Lbl_JP_UnChecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_JP_UnChecked.TabIndex = 2
        Me.Lbl_JP_UnChecked.TabStop = True
        Me.Lbl_JP_UnChecked.Text = "Uncheck All"
        '
        'Lbl_JP_Checked
        '
        Me.Lbl_JP_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JP_Checked.AutoSize = True
        Me.Lbl_JP_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JP_Checked.Location = New System.Drawing.Point(110, 0)
        Me.Lbl_JP_Checked.Name = "Lbl_JP_Checked"
        Me.Lbl_JP_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_JP_Checked.TabIndex = 1
        Me.Lbl_JP_Checked.TabStop = True
        Me.Lbl_JP_Checked.Text = "Check All"
        '
        'ClbJobPosition
        '
        Me.ClbJobPosition.CheckOnClick = True
        Me.ClbJobPosition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbJobPosition.FormattingEnabled = True
        Me.ClbJobPosition.Location = New System.Drawing.Point(3, 16)
        Me.ClbJobPosition.MultiColumn = True
        Me.ClbJobPosition.Name = "ClbJobPosition"
        Me.ClbJobPosition.Size = New System.Drawing.Size(233, 94)
        Me.ClbJobPosition.TabIndex = 3
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.ChkDepartemen)
        Me.GroupBox4.Controls.Add(Me.Lbl_Dept_UnChecked)
        Me.GroupBox4.Controls.Add(Me.Lbl_Dept_Checked)
        Me.GroupBox4.Controls.Add(Me.ClbDepartemen)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(239, 115)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'ChkDepartemen
        '
        Me.ChkDepartemen.AutoSize = True
        Me.ChkDepartemen.Location = New System.Drawing.Point(6, -1)
        Me.ChkDepartemen.Name = "ChkDepartemen"
        Me.ChkDepartemen.Size = New System.Drawing.Size(81, 17)
        Me.ChkDepartemen.TabIndex = 0
        Me.ChkDepartemen.Text = "Department"
        Me.ChkDepartemen.UseVisualStyleBackColor = True
        '
        'Lbl_Dept_UnChecked
        '
        Me.Lbl_Dept_UnChecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_UnChecked.AutoSize = True
        Me.Lbl_Dept_UnChecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_UnChecked.Location = New System.Drawing.Point(168, 0)
        Me.Lbl_Dept_UnChecked.Name = "Lbl_Dept_UnChecked"
        Me.Lbl_Dept_UnChecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Dept_UnChecked.TabIndex = 2
        Me.Lbl_Dept_UnChecked.TabStop = True
        Me.Lbl_Dept_UnChecked.Text = "Uncheck All"
        '
        'Lbl_Dept_Checked
        '
        Me.Lbl_Dept_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_Checked.AutoSize = True
        Me.Lbl_Dept_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_Checked.Location = New System.Drawing.Point(110, 0)
        Me.Lbl_Dept_Checked.Name = "Lbl_Dept_Checked"
        Me.Lbl_Dept_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Dept_Checked.TabIndex = 1
        Me.Lbl_Dept_Checked.TabStop = True
        Me.Lbl_Dept_Checked.Text = "Check All"
        '
        'ClbDepartemen
        '
        Me.ClbDepartemen.CheckOnClick = True
        Me.ClbDepartemen.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbDepartemen.FormattingEnabled = True
        Me.ClbDepartemen.Location = New System.Drawing.Point(3, 16)
        Me.ClbDepartemen.MultiColumn = True
        Me.ClbDepartemen.Name = "ClbDepartemen"
        Me.ClbDepartemen.Size = New System.Drawing.Size(233, 94)
        Me.ClbDepartemen.TabIndex = 3
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(239, 238)
        Me.SplitContainer1.SplitterDistance = 115
        Me.SplitContainer1.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ChkLeaveData)
        Me.GroupBox1.Controls.Add(Me.Lbl_Leave_UnChecked)
        Me.GroupBox1.Controls.Add(Me.Lbl_Leave_Checked)
        Me.GroupBox1.Controls.Add(Me.clbLeave)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(239, 115)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ChkLeaveData
        '
        Me.ChkLeaveData.AutoSize = True
        Me.ChkLeaveData.Location = New System.Drawing.Point(6, -1)
        Me.ChkLeaveData.Name = "ChkLeaveData"
        Me.ChkLeaveData.Size = New System.Drawing.Size(82, 17)
        Me.ChkLeaveData.TabIndex = 0
        Me.ChkLeaveData.Text = "Leave Data"
        Me.ChkLeaveData.UseVisualStyleBackColor = True
        '
        'Lbl_Leave_UnChecked
        '
        Me.Lbl_Leave_UnChecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Leave_UnChecked.AutoSize = True
        Me.Lbl_Leave_UnChecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Leave_UnChecked.Location = New System.Drawing.Point(168, 0)
        Me.Lbl_Leave_UnChecked.Name = "Lbl_Leave_UnChecked"
        Me.Lbl_Leave_UnChecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Leave_UnChecked.TabIndex = 2
        Me.Lbl_Leave_UnChecked.TabStop = True
        Me.Lbl_Leave_UnChecked.Text = "Uncheck All"
        '
        'Lbl_Leave_Checked
        '
        Me.Lbl_Leave_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Leave_Checked.AutoSize = True
        Me.Lbl_Leave_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Leave_Checked.Location = New System.Drawing.Point(110, 0)
        Me.Lbl_Leave_Checked.Name = "Lbl_Leave_Checked"
        Me.Lbl_Leave_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Leave_Checked.TabIndex = 1
        Me.Lbl_Leave_Checked.TabStop = True
        Me.Lbl_Leave_Checked.Text = "Check All"
        '
        'clbLeave
        '
        Me.clbLeave.CheckOnClick = True
        Me.clbLeave.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbLeave.FormattingEnabled = True
        Me.clbLeave.Location = New System.Drawing.Point(3, 16)
        Me.clbLeave.MultiColumn = True
        Me.clbLeave.Name = "clbLeave"
        Me.clbLeave.Size = New System.Drawing.Size(233, 94)
        Me.clbLeave.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvEmployee)
        Me.GroupBox2.Controls.Add(Me.ChkEmployee)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(239, 113)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'dgvEmployee
        '
        Me.dgvEmployee.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode})
        Me.dgvEmployee.Location = New System.Drawing.Point(6, 18)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(227, 89)
        Me.dgvEmployee.TabIndex = 1
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
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.Width = 50
        '
        'ColPrimaryCode
        '
        Me.ColPrimaryCode.HeaderText = "Primary"
        Me.ColPrimaryCode.Name = "ColPrimaryCode"
        Me.ColPrimaryCode.Visible = False
        Me.ColPrimaryCode.Width = 66
        '
        'ChkEmployee
        '
        Me.ChkEmployee.AutoSize = True
        Me.ChkEmployee.Location = New System.Drawing.Point(6, -1)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(72, 17)
        Me.ChkEmployee.TabIndex = 0
        Me.ChkEmployee.Text = "Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Recap", "Detail"})
        Me.cboFormat.Location = New System.Drawing.Point(57, 8)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(143, 21)
        Me.cboFormat.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 107
        Me.Label1.Text = "Format"
        '
        'DtpPeriod
        '
        Me.DtpPeriod.CustomFormat = "dd MMM yyyy"
        Me.DtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPeriod.Location = New System.Drawing.Point(295, 9)
        Me.DtpPeriod.Name = "DtpPeriod"
        Me.DtpPeriod.Size = New System.Drawing.Size(136, 20)
        Me.DtpPeriod.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(252, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 109
        Me.Label2.Text = "Period"
        '
        'BackgroundWorker1
        '
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(255, 36)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(505, 475)
        Me.crv1.TabIndex = 4
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'CboFilterEmployee
        '
        Me.CboFilterEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFilterEmployee.FormattingEnabled = True
        Me.CboFilterEmployee.Items.AddRange(New Object() {"All", "Active", "Not Active"})
        Me.CboFilterEmployee.Location = New System.Drawing.Point(442, 7)
        Me.CboFilterEmployee.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.CboFilterEmployee.Name = "CboFilterEmployee"
        Me.CboFilterEmployee.Size = New System.Drawing.Size(81, 21)
        Me.CboFilterEmployee.TabIndex = 110
        '
        'FrmLaporanCuti
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(772, 519)
        Me.Controls.Add(Me.CboFilterEmployee)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DtpPeriod)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboFormat)
        Me.Controls.Add(Me.SplitContainer3)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.MinimizeBox = False
        Me.Name = "FrmLaporanCuti"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave Report"
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkJobPosition As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_JP_UnChecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_JP_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbJobPosition As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDepartemen As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_Dept_UnChecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Dept_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbDepartemen As System.Windows.Forms.CheckedListBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkLeaveData As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_Leave_UnChecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Leave_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents clbLeave As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ChkEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CboFilterEmployee As System.Windows.Forms.ComboBox
End Class
