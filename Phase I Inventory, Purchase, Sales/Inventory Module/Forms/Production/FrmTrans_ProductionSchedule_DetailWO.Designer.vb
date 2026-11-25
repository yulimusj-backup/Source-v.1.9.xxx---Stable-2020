<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_DetailWO
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtJONumber = New System.Windows.Forms.TextBox
        Me.TxtMONumber = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
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
        Me.ColWOTbl_Modify = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColWOTbl_PrimaryWorkunit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.llbColumnSetting = New System.Windows.Forms.LinkLabel
        CType(Me.dgvItemWO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 12)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "JO Number"
        '
        'TxtJONumber
        '
        Me.TxtJONumber.Location = New System.Drawing.Point(74, 10)
        Me.TxtJONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtJONumber.Name = "TxtJONumber"
        Me.TxtJONumber.ReadOnly = True
        Me.TxtJONumber.Size = New System.Drawing.Size(134, 20)
        Me.TxtJONumber.TabIndex = 1
        '
        'TxtMONumber
        '
        Me.TxtMONumber.Location = New System.Drawing.Point(300, 10)
        Me.TxtMONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtMONumber.Name = "TxtMONumber"
        Me.TxtMONumber.ReadOnly = True
        Me.TxtMONumber.Size = New System.Drawing.Size(134, 20)
        Me.TxtMONumber.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(232, 12)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "MO Number"
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
        Me.dgvItemWO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWOTbl_No, Me.ColWOTbl_Progress, Me.ColWOTbl_WONO, Me.ColWOTbl_CodeNo, Me.ColWOTbl_Desc, Me.ColWOTbl_ProcessCode, Me.ColWOTbl_ProcessDesc, Me.ColWOTbl_WorkUnit, Me.ColWOTbl_PrimaryProcess, Me.ColWOTbl_IdentifierWO, Me.ColWOTbl_IdentifierJO, Me.ColWOTbl_StartDate, Me.ColWOTbl_StartTime, Me.ColWOTbl_EndDate, Me.ColWOTbl_EndTime, Me.ColWOTbl_Remarks, Me.ColWOTbl_Modify, Me.ColWOTbl_PrimaryWorkunit})
        Me.dgvItemWO.Location = New System.Drawing.Point(9, 32)
        Me.dgvItemWO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItemWO.Name = "dgvItemWO"
        Me.dgvItemWO.RowHeadersVisible = False
        Me.dgvItemWO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemWO.RowTemplate.Height = 24
        Me.dgvItemWO.Size = New System.Drawing.Size(817, 322)
        Me.dgvItemWO.TabIndex = 22
        '
        'ColWOTbl_No
        '
        Me.ColWOTbl_No.HeaderText = "No"
        Me.ColWOTbl_No.Name = "ColWOTbl_No"
        Me.ColWOTbl_No.Width = 46
        '
        'ColWOTbl_Progress
        '
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
        Me.ColWOTbl_ProcessDesc.HeaderText = "Process Dsecription"
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
        'ColWOTbl_Modify
        '
        Me.ColWOTbl_Modify.HeaderText = "Modify"
        Me.ColWOTbl_Modify.Name = "ColWOTbl_Modify"
        Me.ColWOTbl_Modify.Width = 44
        '
        'ColWOTbl_PrimaryWorkunit
        '
        Me.ColWOTbl_PrimaryWorkunit.HeaderText = "Primary Workunit"
        Me.ColWOTbl_PrimaryWorkunit.Name = "ColWOTbl_PrimaryWorkunit"
        Me.ColWOTbl_PrimaryWorkunit.ReadOnly = True
        Me.ColWOTbl_PrimaryWorkunit.Visible = False
        Me.ColWOTbl_PrimaryWorkunit.Width = 112
        '
        'llbColumnSetting
        '
        Me.llbColumnSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbColumnSetting.AutoSize = True
        Me.llbColumnSetting.Location = New System.Drawing.Point(7, 355)
        Me.llbColumnSetting.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbColumnSetting.Name = "llbColumnSetting"
        Me.llbColumnSetting.Size = New System.Drawing.Size(88, 13)
        Me.llbColumnSetting.TabIndex = 23
        Me.llbColumnSetting.TabStop = True
        Me.llbColumnSetting.Text = "Columns Settings"
        '
        'FrmTrans_ProductionSchedule_DetailWO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(837, 376)
        Me.Controls.Add(Me.llbColumnSetting)
        Me.Controls.Add(Me.dgvItemWO)
        Me.Controls.Add(Me.TxtMONumber)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtJONumber)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(853, 414)
        Me.Name = "FrmTrans_ProductionSchedule_DetailWO"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Schedule"
        CType(Me.dgvItemWO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtJONumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtMONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dgvItemWO As System.Windows.Forms.DataGridView
    Friend WithEvents llbColumnSetting As System.Windows.Forms.LinkLabel
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
    Friend WithEvents ColWOTbl_Modify As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColWOTbl_PrimaryWorkunit As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
