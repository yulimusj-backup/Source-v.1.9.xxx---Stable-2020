<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch2_Detail
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.DgvWUAllocate_Process = New System.Windows.Forms.DataGridView
        Me.llbColumnsSetting = New System.Windows.Forms.LinkLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvDetailWH = New System.Windows.Forms.DataGridView
        Me.ColDetail_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_TipeWH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_WH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.bgw2 = New System.ComponentModel.BackgroundWorker
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalWorkHour = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimal = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvDetailWH, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(340, 114)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'DgvWUAllocate_Process
        '
        Me.DgvWUAllocate_Process.AllowUserToAddRows = False
        Me.DgvWUAllocate_Process.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_Process.AllowUserToResizeColumns = False
        Me.DgvWUAllocate_Process.AllowUserToResizeRows = False
        Me.DgvWUAllocate_Process.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_Process.ColumnHeadersHeight = 25
        Me.DgvWUAllocate_Process.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_Process.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColJONo, Me.ColWONumber, Me.ColProcessCode, Me.ColCodeNo, Me.ColDesc, Me.ColPrimaryProcess, Me.ColRate, Me.ColQty, Me.ColUnitCode, Me.ColTotalWorkHour, Me.ColStartDate, Me.ColStartTime, Me.ColEndDate, Me.ColEndTime, Me.ColIdentifierWO, Me.ColDecimal})
        Me.DgvWUAllocate_Process.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_Process.Location = New System.Drawing.Point(3, 2)
        Me.DgvWUAllocate_Process.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DgvWUAllocate_Process.MultiSelect = False
        Me.DgvWUAllocate_Process.Name = "DgvWUAllocate_Process"
        Me.DgvWUAllocate_Process.ReadOnly = True
        Me.DgvWUAllocate_Process.RowHeadersVisible = False
        Me.DgvWUAllocate_Process.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvWUAllocate_Process.RowTemplate.Height = 24
        Me.DgvWUAllocate_Process.Size = New System.Drawing.Size(763, 326)
        Me.DgvWUAllocate_Process.TabIndex = 2
        '
        'llbColumnsSetting
        '
        Me.llbColumnsSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbColumnsSetting.AutoSize = True
        Me.llbColumnsSetting.Location = New System.Drawing.Point(5, 402)
        Me.llbColumnsSetting.Name = "llbColumnsSetting"
        Me.llbColumnsSetting.Size = New System.Drawing.Size(117, 17)
        Me.llbColumnsSetting.TabIndex = 3
        Me.llbColumnsSetting.TabStop = True
        Me.llbColumnsSetting.Text = "Columns Settings"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Period"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(67, 12)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(153, 22)
        Me.dtpFrom.TabIndex = 5
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(257, 12)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(153, 22)
        Me.dtpTo.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(227, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "To"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 39)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(777, 359)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvWUAllocate_Process)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Size = New System.Drawing.Size(769, 330)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Work Queue"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvDetailWH)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Size = New System.Drawing.Size(769, 330)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Detail Work Hour"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvDetailWH
        '
        Me.dgvDetailWH.AllowUserToAddRows = False
        Me.dgvDetailWH.AllowUserToDeleteRows = False
        Me.dgvDetailWH.AllowUserToResizeColumns = False
        Me.dgvDetailWH.AllowUserToResizeRows = False
        Me.dgvDetailWH.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetailWH.ColumnHeadersHeight = 25
        Me.dgvDetailWH.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDetailWH.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDetail_No, Me.ColDetail_TipeWH, Me.ColDetail_StartDate, Me.ColDetail_StartTime, Me.ColDetail_EndDate, Me.ColDetail_EndTime, Me.ColDetail_WH})
        Me.dgvDetailWH.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDetailWH.Location = New System.Drawing.Point(3, 2)
        Me.dgvDetailWH.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvDetailWH.Name = "dgvDetailWH"
        Me.dgvDetailWH.ReadOnly = True
        Me.dgvDetailWH.RowHeadersVisible = False
        Me.dgvDetailWH.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDetailWH.RowTemplate.Height = 24
        Me.dgvDetailWH.Size = New System.Drawing.Size(763, 326)
        Me.dgvDetailWH.TabIndex = 0
        '
        'ColDetail_No
        '
        Me.ColDetail_No.HeaderText = "No"
        Me.ColDetail_No.Name = "ColDetail_No"
        Me.ColDetail_No.ReadOnly = True
        Me.ColDetail_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_No.Width = 32
        '
        'ColDetail_TipeWH
        '
        Me.ColDetail_TipeWH.HeaderText = "Type"
        Me.ColDetail_TipeWH.Name = "ColDetail_TipeWH"
        Me.ColDetail_TipeWH.ReadOnly = True
        Me.ColDetail_TipeWH.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_TipeWH.Width = 46
        '
        'ColDetail_StartDate
        '
        Me.ColDetail_StartDate.HeaderText = "Start"
        Me.ColDetail_StartDate.Name = "ColDetail_StartDate"
        Me.ColDetail_StartDate.ReadOnly = True
        Me.ColDetail_StartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_StartDate.Width = 44
        '
        'ColDetail_StartTime
        '
        Me.ColDetail_StartTime.HeaderText = ""
        Me.ColDetail_StartTime.Name = "ColDetail_StartTime"
        Me.ColDetail_StartTime.ReadOnly = True
        Me.ColDetail_StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_StartTime.Width = 5
        '
        'ColDetail_EndDate
        '
        Me.ColDetail_EndDate.HeaderText = "End"
        Me.ColDetail_EndDate.Name = "ColDetail_EndDate"
        Me.ColDetail_EndDate.ReadOnly = True
        Me.ColDetail_EndDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_EndDate.Width = 39
        '
        'ColDetail_EndTime
        '
        Me.ColDetail_EndTime.HeaderText = ""
        Me.ColDetail_EndTime.Name = "ColDetail_EndTime"
        Me.ColDetail_EndTime.ReadOnly = True
        Me.ColDetail_EndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_EndTime.Width = 5
        '
        'ColDetail_WH
        '
        Me.ColDetail_WH.HeaderText = "W/H"
        Me.ColDetail_WH.Name = "ColDetail_WH"
        Me.ColDetail_WH.ReadOnly = True
        Me.ColDetail_WH.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_WH.Width = 41
        '
        'bgw1
        '
        '
        'bgw2
        '
        '
        'ColJONo
        '
        Me.ColJONo.HeaderText = "JO Number"
        Me.ColJONo.MinimumWidth = 95
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.ReadOnly = True
        Me.ColJONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJONo.Width = 95
        '
        'ColWONumber
        '
        Me.ColWONumber.HeaderText = "WO No"
        Me.ColWONumber.Name = "ColWONumber"
        Me.ColWONumber.ReadOnly = True
        Me.ColWONumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWONumber.Width = 60
        '
        'ColProcessCode
        '
        Me.ColProcessCode.HeaderText = "Process"
        Me.ColProcessCode.Name = "ColProcessCode"
        Me.ColProcessCode.ReadOnly = True
        Me.ColProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProcessCode.Width = 65
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 69
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 85
        '
        'ColPrimaryProcess
        '
        Me.ColPrimaryProcess.HeaderText = "Primary Process"
        Me.ColPrimaryProcess.Name = "ColPrimaryProcess"
        Me.ColPrimaryProcess.ReadOnly = True
        Me.ColPrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryProcess.Visible = False
        Me.ColPrimaryProcess.Width = 117
        '
        'ColRate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColRate.HeaderText = "Rate (Unit/Hour)"
        Me.ColRate.Name = "ColRate"
        Me.ColRate.ReadOnly = True
        Me.ColRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRate.Width = 118
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty Produce"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 93
        '
        'ColUnitCode
        '
        Me.ColUnitCode.HeaderText = "Unit"
        Me.ColUnitCode.Name = "ColUnitCode"
        Me.ColUnitCode.ReadOnly = True
        Me.ColUnitCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitCode.Width = 39
        '
        'ColTotalWorkHour
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalWorkHour.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTotalWorkHour.HeaderText = "W/H"
        Me.ColTotalWorkHour.Name = "ColTotalWorkHour"
        Me.ColTotalWorkHour.ReadOnly = True
        Me.ColTotalWorkHour.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTotalWorkHour.Width = 41
        '
        'ColStartDate
        '
        Me.ColStartDate.HeaderText = "Start"
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.ReadOnly = True
        Me.ColStartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStartDate.Width = 44
        '
        'ColStartTime
        '
        Me.ColStartTime.HeaderText = ""
        Me.ColStartTime.Name = "ColStartTime"
        Me.ColStartTime.ReadOnly = True
        Me.ColStartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStartTime.Width = 5
        '
        'ColEndDate
        '
        Me.ColEndDate.HeaderText = "End"
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.ReadOnly = True
        Me.ColEndDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEndDate.Width = 39
        '
        'ColEndTime
        '
        Me.ColEndTime.HeaderText = ""
        Me.ColEndTime.Name = "ColEndTime"
        Me.ColEndTime.ReadOnly = True
        Me.ColEndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEndTime.Width = 5
        '
        'ColIdentifierWO
        '
        Me.ColIdentifierWO.HeaderText = "Identifier WO"
        Me.ColIdentifierWO.Name = "ColIdentifierWO"
        Me.ColIdentifierWO.ReadOnly = True
        Me.ColIdentifierWO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIdentifierWO.Visible = False
        Me.ColIdentifierWO.Width = 96
        '
        'ColDecimal
        '
        Me.ColDecimal.HeaderText = "Decimal"
        Me.ColDecimal.Name = "ColDecimal"
        Me.ColDecimal.ReadOnly = True
        Me.ColDecimal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDecimal.Visible = False
        Me.ColDecimal.Width = 64
        '
        'FrmTrans_ProductionSchedule_NewSch2_Detail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(804, 430)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.llbColumnsSetting)
        Me.Controls.Add(Me.btnClose)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(812, 467)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch2_Detail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Unit Detail Schedule"
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvDetailWH, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents DgvWUAllocate_Process As System.Windows.Forms.DataGridView
    Friend WithEvents llbColumnsSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvDetailWH As System.Windows.Forms.DataGridView
    Friend WithEvents ColDetail_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_TipeWH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_WH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgw2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalWorkHour As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimal As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
