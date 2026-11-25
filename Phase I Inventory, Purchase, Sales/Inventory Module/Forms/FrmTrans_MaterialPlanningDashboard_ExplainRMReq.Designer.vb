<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainRMReq
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
        Me.dgvDataRMReq = New System.Windows.Forms.DataGridView
        Me.ColRMReq_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_MONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_MODate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_JONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_WONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_Process = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_BalanceQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_SpareQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_TotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_UnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_IsWIP = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRMReq_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_RequiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.chkShowWONo = New System.Windows.Forms.CheckBox
        Me.llbSetReqDate = New System.Windows.Forms.LinkLabel
        CType(Me.dgvDataRMReq, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDataRMReq
        '
        Me.dgvDataRMReq.AllowUserToAddRows = False
        Me.dgvDataRMReq.AllowUserToDeleteRows = False
        Me.dgvDataRMReq.AllowUserToResizeColumns = False
        Me.dgvDataRMReq.AllowUserToResizeRows = False
        Me.dgvDataRMReq.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDataRMReq.ColumnHeadersHeight = 25
        Me.dgvDataRMReq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDataRMReq.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRMReq_No, Me.ColRMReq_MONo, Me.ColRMReq_MODate, Me.ColRMReq_JONo, Me.ColRMReq_WONo, Me.ColRMReq_Process, Me.ColRMReq_BalanceQty, Me.ColRMReq_SpareQty, Me.ColRMReq_TotalQty, Me.ColRMReq_UnitCode, Me.ColRMReq_IsWIP, Me.ColRMReq_IdentifierJO, Me.ColRMReq_IdentifierWO, Me.ColRMReq_RequiredDate})
        Me.dgvDataRMReq.Location = New System.Drawing.Point(13, 13)
        Me.dgvDataRMReq.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvDataRMReq.Name = "dgvDataRMReq"
        Me.dgvDataRMReq.RowHeadersVisible = False
        Me.dgvDataRMReq.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDataRMReq.RowTemplate.Height = 24
        Me.dgvDataRMReq.Size = New System.Drawing.Size(916, 438)
        Me.dgvDataRMReq.TabIndex = 8
        '
        'ColRMReq_No
        '
        Me.ColRMReq_No.HeaderText = "No"
        Me.ColRMReq_No.Name = "ColRMReq_No"
        Me.ColRMReq_No.ReadOnly = True
        Me.ColRMReq_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRMReq_No.Width = 32
        '
        'ColRMReq_MONo
        '
        Me.ColRMReq_MONo.HeaderText = "MO No"
        Me.ColRMReq_MONo.Name = "ColRMReq_MONo"
        Me.ColRMReq_MONo.ReadOnly = True
        Me.ColRMReq_MONo.Width = 77
        '
        'ColRMReq_MODate
        '
        Me.ColRMReq_MODate.HeaderText = "MO Date"
        Me.ColRMReq_MODate.Name = "ColRMReq_MODate"
        Me.ColRMReq_MODate.ReadOnly = True
        Me.ColRMReq_MODate.Width = 89
        '
        'ColRMReq_JONo
        '
        Me.ColRMReq_JONo.HeaderText = "JO No"
        Me.ColRMReq_JONo.Name = "ColRMReq_JONo"
        Me.ColRMReq_JONo.ReadOnly = True
        Me.ColRMReq_JONo.Width = 73
        '
        'ColRMReq_WONo
        '
        Me.ColRMReq_WONo.HeaderText = "WO No"
        Me.ColRMReq_WONo.Name = "ColRMReq_WONo"
        Me.ColRMReq_WONo.ReadOnly = True
        Me.ColRMReq_WONo.Width = 79
        '
        'ColRMReq_Process
        '
        Me.ColRMReq_Process.HeaderText = "Process"
        Me.ColRMReq_Process.Name = "ColRMReq_Process"
        Me.ColRMReq_Process.ReadOnly = True
        Me.ColRMReq_Process.Width = 84
        '
        'ColRMReq_BalanceQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMReq_BalanceQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColRMReq_BalanceQty.HeaderText = "Balance Qty"
        Me.ColRMReq_BalanceQty.Name = "ColRMReq_BalanceQty"
        Me.ColRMReq_BalanceQty.Width = 110
        '
        'ColRMReq_SpareQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMReq_SpareQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColRMReq_SpareQty.HeaderText = "Spare Qty"
        Me.ColRMReq_SpareQty.Name = "ColRMReq_SpareQty"
        Me.ColRMReq_SpareQty.Width = 97
        '
        'ColRMReq_TotalQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMReq_TotalQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRMReq_TotalQty.HeaderText = "Total Qty"
        Me.ColRMReq_TotalQty.Name = "ColRMReq_TotalQty"
        Me.ColRMReq_TotalQty.Width = 91
        '
        'ColRMReq_UnitCode
        '
        Me.ColRMReq_UnitCode.HeaderText = "Unit Code"
        Me.ColRMReq_UnitCode.Name = "ColRMReq_UnitCode"
        Me.ColRMReq_UnitCode.ReadOnly = True
        Me.ColRMReq_UnitCode.Width = 95
        '
        'ColRMReq_IsWIP
        '
        Me.ColRMReq_IsWIP.HeaderText = "Is WIP"
        Me.ColRMReq_IsWIP.Name = "ColRMReq_IsWIP"
        Me.ColRMReq_IsWIP.ReadOnly = True
        Me.ColRMReq_IsWIP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMReq_IsWIP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMReq_IsWIP.Width = 72
        '
        'ColRMReq_IdentifierJO
        '
        Me.ColRMReq_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColRMReq_IdentifierJO.Name = "ColRMReq_IdentifierJO"
        Me.ColRMReq_IdentifierJO.ReadOnly = True
        Me.ColRMReq_IdentifierJO.Visible = False
        Me.ColRMReq_IdentifierJO.Width = 109
        '
        'ColRMReq_IdentifierWO
        '
        Me.ColRMReq_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColRMReq_IdentifierWO.Name = "ColRMReq_IdentifierWO"
        Me.ColRMReq_IdentifierWO.ReadOnly = True
        Me.ColRMReq_IdentifierWO.Visible = False
        Me.ColRMReq_IdentifierWO.Width = 115
        '
        'ColRMReq_RequiredDate
        '
        Me.ColRMReq_RequiredDate.HeaderText = "Required Date"
        Me.ColRMReq_RequiredDate.Name = "ColRMReq_RequiredDate"
        Me.ColRMReq_RequiredDate.Width = 125
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(560, 164)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'chkShowWONo
        '
        Me.chkShowWONo.AutoSize = True
        Me.chkShowWONo.Location = New System.Drawing.Point(12, 458)
        Me.chkShowWONo.Name = "chkShowWONo"
        Me.chkShowWONo.Size = New System.Drawing.Size(114, 21)
        Me.chkShowWONo.TabIndex = 10
        Me.chkShowWONo.Text = "Show WO No"
        Me.chkShowWONo.UseVisualStyleBackColor = True
        '
        'llbSetReqDate
        '
        Me.llbSetReqDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbSetReqDate.AutoSize = True
        Me.llbSetReqDate.Location = New System.Drawing.Point(132, 460)
        Me.llbSetReqDate.Name = "llbSetReqDate"
        Me.llbSetReqDate.Size = New System.Drawing.Size(167, 17)
        Me.llbSetReqDate.TabIndex = 11
        Me.llbSetReqDate.TabStop = True
        Me.llbSetReqDate.Text = "Set Required Date By JO"
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainRMReq
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(942, 491)
        Me.Controls.Add(Me.llbSetReqDate)
        Me.Controls.Add(Me.chkShowWONo)
        Me.Controls.Add(Me.dgvDataRMReq)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(950, 528)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainRMReq"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Raw Material Request Detail"
        CType(Me.dgvDataRMReq, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvDataRMReq As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents chkShowWONo As System.Windows.Forms.CheckBox
    Friend WithEvents ColRMReq_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_MONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_MODate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_JONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_WONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_Process As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_BalanceQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_SpareQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_TotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_UnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_IsWIP As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRMReq_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_RequiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbSetReqDate As System.Windows.Forms.LinkLabel
End Class
