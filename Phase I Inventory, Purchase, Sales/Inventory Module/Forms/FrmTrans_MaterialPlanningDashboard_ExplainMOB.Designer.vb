<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainMOB
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
        Me.dgvDataMOB = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMODate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProses = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTargetFinishDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProgress = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.dgvDataMOB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDataMOB
        '
        Me.dgvDataMOB.AllowUserToAddRows = False
        Me.dgvDataMOB.AllowUserToDeleteRows = False
        Me.dgvDataMOB.AllowUserToResizeColumns = False
        Me.dgvDataMOB.AllowUserToResizeRows = False
        Me.dgvDataMOB.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDataMOB.ColumnHeadersHeight = 25
        Me.dgvDataMOB.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDataMOB.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColMONo, Me.ColMODate, Me.ColJONo, Me.ColWONo, Me.ColProses, Me.ColBalanceQty, Me.ColUnitCode, Me.ColTargetFinishDate, Me.ColProgress, Me.ColIdentifierJO, Me.ColIdentifierWO})
        Me.dgvDataMOB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDataMOB.Location = New System.Drawing.Point(0, 0)
        Me.dgvDataMOB.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvDataMOB.Name = "dgvDataMOB"
        Me.dgvDataMOB.ReadOnly = True
        Me.dgvDataMOB.RowHeadersVisible = False
        Me.dgvDataMOB.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDataMOB.RowTemplate.Height = 24
        Me.dgvDataMOB.Size = New System.Drawing.Size(942, 491)
        Me.dgvDataMOB.TabIndex = 6
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColMONo
        '
        Me.ColMONo.HeaderText = "MO No"
        Me.ColMONo.Name = "ColMONo"
        Me.ColMONo.ReadOnly = True
        Me.ColMONo.Width = 77
        '
        'ColMODate
        '
        Me.ColMODate.HeaderText = "MO Date"
        Me.ColMODate.Name = "ColMODate"
        Me.ColMODate.ReadOnly = True
        Me.ColMODate.Width = 89
        '
        'ColJONo
        '
        Me.ColJONo.HeaderText = "JO No"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.ReadOnly = True
        Me.ColJONo.Width = 73
        '
        'ColWONo
        '
        Me.ColWONo.HeaderText = "WO No"
        Me.ColWONo.Name = "ColWONo"
        Me.ColWONo.ReadOnly = True
        Me.ColWONo.Width = 79
        '
        'ColProses
        '
        Me.ColProses.HeaderText = "Process"
        Me.ColProses.Name = "ColProses"
        Me.ColProses.ReadOnly = True
        Me.ColProses.Width = 84
        '
        'ColBalanceQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBalanceQty.HeaderText = "Balance Qty"
        Me.ColBalanceQty.Name = "ColBalanceQty"
        Me.ColBalanceQty.ReadOnly = True
        Me.ColBalanceQty.Width = 110
        '
        'ColUnitCode
        '
        Me.ColUnitCode.HeaderText = "Unit Code"
        Me.ColUnitCode.Name = "ColUnitCode"
        Me.ColUnitCode.ReadOnly = True
        Me.ColUnitCode.Width = 95
        '
        'ColTargetFinishDate
        '
        Me.ColTargetFinishDate.HeaderText = "Target Finish Date"
        Me.ColTargetFinishDate.Name = "ColTargetFinishDate"
        Me.ColTargetFinishDate.ReadOnly = True
        Me.ColTargetFinishDate.Visible = False
        Me.ColTargetFinishDate.Width = 150
        '
        'ColProgress
        '
        Me.ColProgress.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColProgress.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColProgress.HeaderText = "Progress"
        Me.ColProgress.Name = "ColProgress"
        Me.ColProgress.ReadOnly = True
        Me.ColProgress.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColProgress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColProgress.Visible = False
        Me.ColProgress.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColProgress.Width = 90
        '
        'ColIdentifierJO
        '
        Me.ColIdentifierJO.HeaderText = "Identifier JO"
        Me.ColIdentifierJO.Name = "ColIdentifierJO"
        Me.ColIdentifierJO.ReadOnly = True
        Me.ColIdentifierJO.Visible = False
        Me.ColIdentifierJO.Width = 109
        '
        'ColIdentifierWO
        '
        Me.ColIdentifierWO.HeaderText = "IdentifierWO"
        Me.ColIdentifierWO.Name = "ColIdentifierWO"
        Me.ColIdentifierWO.ReadOnly = True
        Me.ColIdentifierWO.Visible = False
        Me.ColIdentifierWO.Width = 111
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(560, 164)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainMOB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(942, 491)
        Me.Controls.Add(Me.dgvDataMOB)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(950, 528)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainMOB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manufacturing Order Balance Detail"
        CType(Me.dgvDataMOB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDataMOB As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMODate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTargetFinishDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProgress As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
