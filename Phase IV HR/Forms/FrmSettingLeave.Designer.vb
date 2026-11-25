<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSettingLeave
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
        Me.dgvLeaveCode = New System.Windows.Forms.DataGridView
        Me.BtnOk = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.dtpPeriode = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeaveCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColPrimaryLeave = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvLeaveCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvLeaveCode
        '
        Me.dgvLeaveCode.AllowUserToAddRows = False
        Me.dgvLeaveCode.AllowUserToDeleteRows = False
        Me.dgvLeaveCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvLeaveCode.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvLeaveCode.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvLeaveCode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLeaveCode.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColLeaveCode, Me.ColStatus, Me.ColPrimaryLeave})
        Me.dgvLeaveCode.Location = New System.Drawing.Point(13, 38)
        Me.dgvLeaveCode.Name = "dgvLeaveCode"
        Me.dgvLeaveCode.Size = New System.Drawing.Size(331, 266)
        Me.dgvLeaveCode.TabIndex = 0
        '
        'BtnOk
        '
        Me.BtnOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOk.Location = New System.Drawing.Point(188, 310)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(75, 23)
        Me.BtnOk.TabIndex = 1
        Me.BtnOk.Text = "Ok"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.Location = New System.Drawing.Point(269, 310)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(75, 23)
        Me.BtnCancel.TabIndex = 2
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'dtpPeriode
        '
        Me.dtpPeriode.CustomFormat = "yyyy"
        Me.dtpPeriode.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriode.Location = New System.Drawing.Point(83, 6)
        Me.dtpPeriode.Name = "dtpPeriode"
        Me.dtpPeriode.ShowUpDown = True
        Me.dtpPeriode.Size = New System.Drawing.Size(71, 20)
        Me.dtpPeriode.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Periode:"
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColLeaveCode
        '
        Me.ColLeaveCode.HeaderText = "Leave Code"
        Me.ColLeaveCode.Name = "ColLeaveCode"
        Me.ColLeaveCode.ReadOnly = True
        Me.ColLeaveCode.Width = 90
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Items.AddRange(New Object() {"General", "Special"})
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColStatus.Width = 62
        '
        'ColPrimaryLeave
        '
        Me.ColPrimaryLeave.HeaderText = "Primary Leave"
        Me.ColPrimaryLeave.Name = "ColPrimaryLeave"
        Me.ColPrimaryLeave.Visible = False
        Me.ColPrimaryLeave.Width = 99
        '
        'FrmSettingLeave
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(356, 345)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpPeriode)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOk)
        Me.Controls.Add(Me.dgvLeaveCode)
        Me.Name = "FrmSettingLeave"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave Setting"
        CType(Me.dgvLeaveCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvLeaveCode As System.Windows.Forms.DataGridView
    Friend WithEvents BtnOk As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents dtpPeriode As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColPrimaryLeave As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
