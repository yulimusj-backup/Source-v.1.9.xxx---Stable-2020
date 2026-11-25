<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_EditOT2
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
        Me.lblFirmPlanDate = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvFilterWU = New System.Windows.Forms.DataGridView
        Me.ColFilterWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterWU_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColFilterWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpEndTime = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpStartTime = New System.Windows.Forms.DateTimePicker
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvFilterWU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblFirmPlanDate
        '
        Me.lblFirmPlanDate.AutoSize = True
        Me.lblFirmPlanDate.Location = New System.Drawing.Point(87, 9)
        Me.lblFirmPlanDate.Name = "lblFirmPlanDate"
        Me.lblFirmPlanDate.Size = New System.Drawing.Size(101, 17)
        Me.lblFirmPlanDate.TabIndex = 0
        Me.lblFirmPlanDate.Text = "Firm Plan Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 17)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "Firm Plan"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvFilterWU)
        Me.GroupBox1.Location = New System.Drawing.Point(252, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(387, 242)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Work Unit"
        '
        'dgvFilterWU
        '
        Me.dgvFilterWU.AllowUserToResizeColumns = False
        Me.dgvFilterWU.AllowUserToResizeRows = False
        Me.dgvFilterWU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFilterWU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvFilterWU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilterWU_No, Me.ColFilterWU_Level, Me.ColFilterWU_Code, Me.ColFilterWU_Desc, Me.ColFilterWU_Primary})
        Me.dgvFilterWU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFilterWU.Location = New System.Drawing.Point(3, 18)
        Me.dgvFilterWU.Name = "dgvFilterWU"
        Me.dgvFilterWU.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvFilterWU.RowTemplate.Height = 24
        Me.dgvFilterWU.Size = New System.Drawing.Size(381, 221)
        Me.dgvFilterWU.TabIndex = 8
        '
        'ColFilterWU_No
        '
        Me.ColFilterWU_No.HeaderText = "No"
        Me.ColFilterWU_No.Name = "ColFilterWU_No"
        Me.ColFilterWU_No.Width = 51
        '
        'ColFilterWU_Level
        '
        Me.ColFilterWU_Level.HeaderText = "Level"
        Me.ColFilterWU_Level.Items.AddRange(New Object() {"Work Unit Group", "Work Unit"})
        Me.ColFilterWU_Level.Name = "ColFilterWU_Level"
        Me.ColFilterWU_Level.Width = 48
        '
        'ColFilterWU_Code
        '
        Me.ColFilterWU_Code.HeaderText = "Code"
        Me.ColFilterWU_Code.Name = "ColFilterWU_Code"
        Me.ColFilterWU_Code.Width = 66
        '
        'ColFilterWU_Desc
        '
        Me.ColFilterWU_Desc.HeaderText = "Description"
        Me.ColFilterWU_Desc.Name = "ColFilterWU_Desc"
        Me.ColFilterWU_Desc.Width = 104
        '
        'ColFilterWU_Primary
        '
        Me.ColFilterWU_Primary.HeaderText = "Primary"
        Me.ColFilterWU_Primary.Name = "ColFilterWU_Primary"
        Me.ColFilterWU_Primary.Visible = False
        Me.ColFilterWU_Primary.Width = 81
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 17)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "End Date"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(90, 57)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(156, 22)
        Me.dtpEndDate.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 17)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Start Date"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(90, 29)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(156, 22)
        Me.dtpStartDate.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 17)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "End Time"
        '
        'dtpEndTime
        '
        Me.dtpEndTime.CustomFormat = "HH.mm"
        Me.dtpEndTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndTime.Location = New System.Drawing.Point(90, 113)
        Me.dtpEndTime.Name = "dtpEndTime"
        Me.dtpEndTime.Size = New System.Drawing.Size(156, 22)
        Me.dtpEndTime.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 17)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Start Time"
        '
        'dtpStartTime
        '
        Me.dtpStartTime.CustomFormat = "HH.mm"
        Me.dtpStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartTime.Location = New System.Drawing.Point(90, 85)
        Me.dtpStartTime.Name = "dtpStartTime"
        Me.dtpStartTime.Size = New System.Drawing.Size(156, 22)
        Me.dtpStartTime.TabIndex = 3
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(90, 141)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 25)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "Ok"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(171, 141)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'BackgroundWorker3
        '
        '
        'FrmTrans_ProductionSchedule_EditOT2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(651, 263)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpEndTime)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dtpStartTime)
        Me.Controls.Add(Me.lblFirmPlanDate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpStartDate)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(659, 300)
        Me.Name = "FrmTrans_ProductionSchedule_EditOT2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Overtime"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvFilterWU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFirmPlanDate As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvFilterWU As System.Windows.Forms.DataGridView
    Friend WithEvents ColFilterWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterWU_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColFilterWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpEndTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpStartTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
End Class
