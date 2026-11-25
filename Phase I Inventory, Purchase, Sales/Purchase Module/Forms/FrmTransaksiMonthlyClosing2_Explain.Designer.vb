<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing2_Explain
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
        Me.TxtExplain = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.dgvItemCost_DetailTransaction = New System.Windows.Forms.DataGridView
        Me.ColCostCodeNoD_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_AllocatedAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_CostFactor = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostCodeNoD_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label7 = New System.Windows.Forms.Label
        Me.LblUnits = New System.Windows.Forms.Label
        Me.llbAlocatedOverhead = New System.Windows.Forms.LinkLabel
        Me.LblTotalQty = New System.Windows.Forms.Label
        Me.LblTotalOverheadAmount = New System.Windows.Forms.Label
        Me.LblCostGroup = New System.Windows.Forms.Label
        Me.LblDocumentNo = New System.Windows.Forms.Label
        Me.LblProcess = New System.Windows.Forms.Label
        Me.LblTotalQty_Bundling = New System.Windows.Forms.Label
        Me.LblUnits_Bundling = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        CType(Me.dgvItemCost_DetailTransaction, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtExplain
        '
        Me.TxtExplain.BackColor = System.Drawing.SystemColors.Control
        Me.TxtExplain.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtExplain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtExplain.Location = New System.Drawing.Point(3, 18)
        Me.TxtExplain.Multiline = True
        Me.TxtExplain.Name = "TxtExplain"
        Me.TxtExplain.Size = New System.Drawing.Size(668, 160)
        Me.TxtExplain.TabIndex = 0
        Me.TxtExplain.Text = "TxtExplain"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Cost Group"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(344, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Total Overhead"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(344, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Total Qty"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 17)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Document No"
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(151, 68)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(162, 22)
        Me.dtpDate.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(344, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(133, 17)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Overhead Allocated"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 71)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(38, 17)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Date"
        '
        'dgvItemCost_DetailTransaction
        '
        Me.dgvItemCost_DetailTransaction.AllowUserToAddRows = False
        Me.dgvItemCost_DetailTransaction.AllowUserToDeleteRows = False
        Me.dgvItemCost_DetailTransaction.AllowUserToResizeColumns = False
        Me.dgvItemCost_DetailTransaction.AllowUserToResizeRows = False
        Me.dgvItemCost_DetailTransaction.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemCost_DetailTransaction.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemCost_DetailTransaction.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemCost_DetailTransaction.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCostCodeNoD_CodeNo, Me.ColCostCodeNoD_Desc, Me.ColCostCodeNoD_Qty, Me.ColCostCodeNoD_Unit, Me.ColCostCodeNoD_AllocatedAmount, Me.ColCostCodeNoD_CostFactor, Me.ColCostCodeNoD_PrimaryCodeNo, Me.ColCostCodeNoD_PrimaryUnit, Me.ColCostCodeNoD_Identifier, Me.ColCostCodeNoD_PrimaryProcess})
        Me.dgvItemCost_DetailTransaction.Location = New System.Drawing.Point(12, 123)
        Me.dgvItemCost_DetailTransaction.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItemCost_DetailTransaction.MultiSelect = False
        Me.dgvItemCost_DetailTransaction.Name = "dgvItemCost_DetailTransaction"
        Me.dgvItemCost_DetailTransaction.ReadOnly = True
        Me.dgvItemCost_DetailTransaction.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemCost_DetailTransaction.RowTemplate.Height = 24
        Me.dgvItemCost_DetailTransaction.Size = New System.Drawing.Size(674, 218)
        Me.dgvItemCost_DetailTransaction.TabIndex = 146
        '
        'ColCostCodeNoD_CodeNo
        '
        Me.ColCostCodeNoD_CodeNo.HeaderText = "Code No"
        Me.ColCostCodeNoD_CodeNo.Name = "ColCostCodeNoD_CodeNo"
        Me.ColCostCodeNoD_CodeNo.ReadOnly = True
        Me.ColCostCodeNoD_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_CodeNo.Width = 69
        '
        'ColCostCodeNoD_Desc
        '
        Me.ColCostCodeNoD_Desc.HeaderText = "Description"
        Me.ColCostCodeNoD_Desc.Name = "ColCostCodeNoD_Desc"
        Me.ColCostCodeNoD_Desc.ReadOnly = True
        Me.ColCostCodeNoD_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_Desc.Width = 85
        '
        'ColCostCodeNoD_Qty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostCodeNoD_Qty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColCostCodeNoD_Qty.HeaderText = "Qty"
        Me.ColCostCodeNoD_Qty.Name = "ColCostCodeNoD_Qty"
        Me.ColCostCodeNoD_Qty.ReadOnly = True
        Me.ColCostCodeNoD_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_Qty.Width = 36
        '
        'ColCostCodeNoD_Unit
        '
        Me.ColCostCodeNoD_Unit.HeaderText = "Unit"
        Me.ColCostCodeNoD_Unit.Name = "ColCostCodeNoD_Unit"
        Me.ColCostCodeNoD_Unit.ReadOnly = True
        Me.ColCostCodeNoD_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_Unit.Width = 39
        '
        'ColCostCodeNoD_AllocatedAmount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostCodeNoD_AllocatedAmount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColCostCodeNoD_AllocatedAmount.HeaderText = "Allocated Amount"
        Me.ColCostCodeNoD_AllocatedAmount.Name = "ColCostCodeNoD_AllocatedAmount"
        Me.ColCostCodeNoD_AllocatedAmount.ReadOnly = True
        Me.ColCostCodeNoD_AllocatedAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_AllocatedAmount.Width = 124
        '
        'ColCostCodeNoD_CostFactor
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostCodeNoD_CostFactor.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColCostCodeNoD_CostFactor.HeaderText = "Cost Factor (%)"
        Me.ColCostCodeNoD_CostFactor.Name = "ColCostCodeNoD_CostFactor"
        Me.ColCostCodeNoD_CostFactor.ReadOnly = True
        Me.ColCostCodeNoD_CostFactor.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_CostFactor.Width = 112
        '
        'ColCostCodeNoD_PrimaryCodeNo
        '
        Me.ColCostCodeNoD_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColCostCodeNoD_PrimaryCodeNo.Name = "ColCostCodeNoD_PrimaryCodeNo"
        Me.ColCostCodeNoD_PrimaryCodeNo.ReadOnly = True
        Me.ColCostCodeNoD_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_PrimaryCodeNo.Visible = False
        Me.ColCostCodeNoD_PrimaryCodeNo.Width = 121
        '
        'ColCostCodeNoD_PrimaryUnit
        '
        Me.ColCostCodeNoD_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColCostCodeNoD_PrimaryUnit.Name = "ColCostCodeNoD_PrimaryUnit"
        Me.ColCostCodeNoD_PrimaryUnit.ReadOnly = True
        Me.ColCostCodeNoD_PrimaryUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_PrimaryUnit.Visible = False
        Me.ColCostCodeNoD_PrimaryUnit.Width = 91
        '
        'ColCostCodeNoD_Identifier
        '
        Me.ColCostCodeNoD_Identifier.HeaderText = "Identifier Group"
        Me.ColCostCodeNoD_Identifier.Name = "ColCostCodeNoD_Identifier"
        Me.ColCostCodeNoD_Identifier.ReadOnly = True
        Me.ColCostCodeNoD_Identifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostCodeNoD_Identifier.Visible = False
        Me.ColCostCodeNoD_Identifier.Width = 112
        '
        'ColCostCodeNoD_PrimaryProcess
        '
        Me.ColCostCodeNoD_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColCostCodeNoD_PrimaryProcess.Name = "ColCostCodeNoD_PrimaryProcess"
        Me.ColCostCodeNoD_PrimaryProcess.ReadOnly = True
        Me.ColCostCodeNoD_PrimaryProcess.Visible = False
        Me.ColCostCodeNoD_PrimaryProcess.Width = 136
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 99)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(59, 17)
        Me.Label7.TabIndex = 147
        Me.Label7.Text = "Process"
        '
        'LblUnits
        '
        Me.LblUnits.AutoSize = True
        Me.LblUnits.Location = New System.Drawing.Point(620, 43)
        Me.LblUnits.Name = "LblUnits"
        Me.LblUnits.Size = New System.Drawing.Size(40, 17)
        Me.LblUnits.TabIndex = 149
        Me.LblUnits.Text = "Units"
        '
        'llbAlocatedOverhead
        '
        Me.llbAlocatedOverhead.AutoSize = True
        Me.llbAlocatedOverhead.Location = New System.Drawing.Point(598, 99)
        Me.llbAlocatedOverhead.Name = "llbAlocatedOverhead"
        Me.llbAlocatedOverhead.Size = New System.Drawing.Size(16, 17)
        Me.llbAlocatedOverhead.TabIndex = 150
        Me.llbAlocatedOverhead.TabStop = True
        Me.llbAlocatedOverhead.Text = "0"
        '
        'LblTotalQty
        '
        Me.LblTotalQty.AutoSize = True
        Me.LblTotalQty.Location = New System.Drawing.Point(598, 43)
        Me.LblTotalQty.Name = "LblTotalQty"
        Me.LblTotalQty.Size = New System.Drawing.Size(16, 17)
        Me.LblTotalQty.TabIndex = 151
        Me.LblTotalQty.Text = "0"
        '
        'LblTotalOverheadAmount
        '
        Me.LblTotalOverheadAmount.AutoSize = True
        Me.LblTotalOverheadAmount.Location = New System.Drawing.Point(598, 15)
        Me.LblTotalOverheadAmount.Name = "LblTotalOverheadAmount"
        Me.LblTotalOverheadAmount.Size = New System.Drawing.Size(16, 17)
        Me.LblTotalOverheadAmount.TabIndex = 152
        Me.LblTotalOverheadAmount.Text = "0"
        '
        'LblCostGroup
        '
        Me.LblCostGroup.AutoSize = True
        Me.LblCostGroup.Location = New System.Drawing.Point(151, 15)
        Me.LblCostGroup.Name = "LblCostGroup"
        Me.LblCostGroup.Size = New System.Drawing.Size(95, 17)
        Me.LblCostGroup.TabIndex = 153
        Me.LblCostGroup.Text = "LblCostGroup"
        '
        'LblDocumentNo
        '
        Me.LblDocumentNo.AutoSize = True
        Me.LblDocumentNo.Location = New System.Drawing.Point(151, 43)
        Me.LblDocumentNo.Name = "LblDocumentNo"
        Me.LblDocumentNo.Size = New System.Drawing.Size(109, 17)
        Me.LblDocumentNo.TabIndex = 154
        Me.LblDocumentNo.Text = "LblDocumentNo"
        '
        'LblProcess
        '
        Me.LblProcess.AutoSize = True
        Me.LblProcess.Location = New System.Drawing.Point(151, 99)
        Me.LblProcess.Name = "LblProcess"
        Me.LblProcess.Size = New System.Drawing.Size(78, 17)
        Me.LblProcess.TabIndex = 155
        Me.LblProcess.Text = "LblProcess"
        '
        'LblTotalQty_Bundling
        '
        Me.LblTotalQty_Bundling.AutoSize = True
        Me.LblTotalQty_Bundling.Location = New System.Drawing.Point(598, 71)
        Me.LblTotalQty_Bundling.Name = "LblTotalQty_Bundling"
        Me.LblTotalQty_Bundling.Size = New System.Drawing.Size(16, 17)
        Me.LblTotalQty_Bundling.TabIndex = 158
        Me.LblTotalQty_Bundling.Text = "0"
        '
        'LblUnits_Bundling
        '
        Me.LblUnits_Bundling.AutoSize = True
        Me.LblUnits_Bundling.Location = New System.Drawing.Point(620, 71)
        Me.LblUnits_Bundling.Name = "LblUnits_Bundling"
        Me.LblUnits_Bundling.Size = New System.Drawing.Size(40, 17)
        Me.LblUnits_Bundling.TabIndex = 157
        Me.LblUnits_Bundling.Text = "Units"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(344, 71)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(140, 17)
        Me.Label10.TabIndex = 156
        Me.Label10.Text = "Total Qty In Bundling"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtExplain)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 346)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(674, 181)
        Me.GroupBox1.TabIndex = 159
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Explanation"
        '
        'FrmTransaksiMonthlyClosing2_Explain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(698, 539)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.LblTotalQty_Bundling)
        Me.Controls.Add(Me.LblUnits_Bundling)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.LblProcess)
        Me.Controls.Add(Me.LblDocumentNo)
        Me.Controls.Add(Me.LblCostGroup)
        Me.Controls.Add(Me.LblTotalOverheadAmount)
        Me.Controls.Add(Me.LblTotalQty)
        Me.Controls.Add(Me.llbAlocatedOverhead)
        Me.Controls.Add(Me.LblUnits)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.dgvItemCost_DetailTransaction)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dtpDate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(706, 576)
        Me.Name = "FrmTransaksiMonthlyClosing2_Explain"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Calculation Explanation"
        CType(Me.dgvItemCost_DetailTransaction, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtExplain As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dgvItemCost_DetailTransaction As System.Windows.Forms.DataGridView
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents LblUnits As System.Windows.Forms.Label
    Friend WithEvents llbAlocatedOverhead As System.Windows.Forms.LinkLabel
    Friend WithEvents LblTotalQty As System.Windows.Forms.Label
    Friend WithEvents LblTotalOverheadAmount As System.Windows.Forms.Label
    Friend WithEvents LblCostGroup As System.Windows.Forms.Label
    Friend WithEvents LblDocumentNo As System.Windows.Forms.Label
    Friend WithEvents LblProcess As System.Windows.Forms.Label
    Friend WithEvents LblTotalQty_Bundling As System.Windows.Forms.Label
    Friend WithEvents LblUnits_Bundling As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ColCostCodeNoD_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_AllocatedAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_CostFactor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostCodeNoD_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
