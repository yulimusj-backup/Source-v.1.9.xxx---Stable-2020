<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_EditQty
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.lblTotalQtyProduce = New System.Windows.Forms.Label
        Me.DgvWUAllocate_WU = New System.Windows.Forms.DataGridView
        Me.CoAll_WU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CoAll_WU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_ProdRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Efficiency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PlanRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_ProduceQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.lblDeviasi = New System.Windows.Forms.Label
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTotalQtyProduce
        '
        Me.lblTotalQtyProduce.AutoSize = True
        Me.lblTotalQtyProduce.Location = New System.Drawing.Point(12, 9)
        Me.lblTotalQtyProduce.Name = "lblTotalQtyProduce"
        Me.lblTotalQtyProduce.Size = New System.Drawing.Size(135, 17)
        Me.lblTotalQtyProduce.TabIndex = 0
        Me.lblTotalQtyProduce.Text = "Total Qty Produce : "
        '
        'DgvWUAllocate_WU
        '
        Me.DgvWUAllocate_WU.AllowUserToAddRows = False
        Me.DgvWUAllocate_WU.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_WU.AllowUserToResizeRows = False
        Me.DgvWUAllocate_WU.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvWUAllocate_WU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_WU.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWUAllocate_WU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_WU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CoAll_WU_No, Me.CoAll_WU_Allocate, Me.CoAll_WU_Code, Me.CoAll_WU_Desc, Me.CoAll_WU_ProdRate, Me.CoAll_WU_Efficiency, Me.CoAll_WU_PlanRate, Me.CoAll_WU_PrimaryWU, Me.CoAll_WU_ProduceQty})
        Me.DgvWUAllocate_WU.Location = New System.Drawing.Point(13, 30)
        Me.DgvWUAllocate_WU.Margin = New System.Windows.Forms.Padding(4)
        Me.DgvWUAllocate_WU.MultiSelect = False
        Me.DgvWUAllocate_WU.Name = "DgvWUAllocate_WU"
        Me.DgvWUAllocate_WU.RowHeadersVisible = False
        Me.DgvWUAllocate_WU.RowTemplate.Height = 24
        Me.DgvWUAllocate_WU.Size = New System.Drawing.Size(623, 191)
        Me.DgvWUAllocate_WU.TabIndex = 4
        '
        'CoAll_WU_No
        '
        Me.CoAll_WU_No.HeaderText = "Priority"
        Me.CoAll_WU_No.Name = "CoAll_WU_No"
        Me.CoAll_WU_No.ReadOnly = True
        Me.CoAll_WU_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_No.Width = 58
        '
        'CoAll_WU_Allocate
        '
        Me.CoAll_WU_Allocate.HeaderText = "Use"
        Me.CoAll_WU_Allocate.Name = "CoAll_WU_Allocate"
        Me.CoAll_WU_Allocate.Width = 39
        '
        'CoAll_WU_Code
        '
        Me.CoAll_WU_Code.HeaderText = "Code"
        Me.CoAll_WU_Code.Name = "CoAll_WU_Code"
        Me.CoAll_WU_Code.ReadOnly = True
        Me.CoAll_WU_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Code.Width = 47
        '
        'CoAll_WU_Desc
        '
        Me.CoAll_WU_Desc.HeaderText = "Description"
        Me.CoAll_WU_Desc.Name = "CoAll_WU_Desc"
        Me.CoAll_WU_Desc.ReadOnly = True
        Me.CoAll_WU_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Desc.Width = 85
        '
        'CoAll_WU_ProdRate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_ProdRate.DefaultCellStyle = DataGridViewCellStyle1
        Me.CoAll_WU_ProdRate.HeaderText = "Rate (Unit/Hour)"
        Me.CoAll_WU_ProdRate.Name = "CoAll_WU_ProdRate"
        Me.CoAll_WU_ProdRate.ReadOnly = True
        Me.CoAll_WU_ProdRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_ProdRate.Visible = False
        Me.CoAll_WU_ProdRate.Width = 118
        '
        'CoAll_WU_Efficiency
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_Efficiency.DefaultCellStyle = DataGridViewCellStyle2
        Me.CoAll_WU_Efficiency.HeaderText = "Efficiency (%)"
        Me.CoAll_WU_Efficiency.Name = "CoAll_WU_Efficiency"
        Me.CoAll_WU_Efficiency.ReadOnly = True
        Me.CoAll_WU_Efficiency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Efficiency.Visible = False
        '
        'CoAll_WU_PlanRate
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_PlanRate.DefaultCellStyle = DataGridViewCellStyle3
        Me.CoAll_WU_PlanRate.HeaderText = "Plan Rate (Unit/Hour)"
        Me.CoAll_WU_PlanRate.Name = "CoAll_WU_PlanRate"
        Me.CoAll_WU_PlanRate.ReadOnly = True
        Me.CoAll_WU_PlanRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_PlanRate.Width = 150
        '
        'CoAll_WU_PrimaryWU
        '
        Me.CoAll_WU_PrimaryWU.HeaderText = "Primary"
        Me.CoAll_WU_PrimaryWU.Name = "CoAll_WU_PrimaryWU"
        Me.CoAll_WU_PrimaryWU.ReadOnly = True
        Me.CoAll_WU_PrimaryWU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_PrimaryWU.Visible = False
        Me.CoAll_WU_PrimaryWU.Width = 62
        '
        'CoAll_WU_ProduceQty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_ProduceQty.DefaultCellStyle = DataGridViewCellStyle4
        Me.CoAll_WU_ProduceQty.HeaderText = "Qty"
        Me.CoAll_WU_ProduceQty.Name = "CoAll_WU_ProduceQty"
        Me.CoAll_WU_ProduceQty.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CoAll_WU_ProduceQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_ProduceQty.Width = 36
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(561, 228)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(480, 228)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "Ok"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'lblDeviasi
        '
        Me.lblDeviasi.AutoSize = True
        Me.lblDeviasi.Location = New System.Drawing.Point(10, 231)
        Me.lblDeviasi.Name = "lblDeviasi"
        Me.lblDeviasi.Size = New System.Drawing.Size(79, 17)
        Me.lblDeviasi.TabIndex = 7
        Me.lblDeviasi.Text = "Deviation : "
        '
        'FrmTrans_ProductionSchedule_EditQty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(649, 263)
        Me.Controls.Add(Me.lblDeviasi)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.DgvWUAllocate_WU)
        Me.Controls.Add(Me.lblTotalQtyProduce)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(657, 300)
        Me.Name = "FrmTrans_ProductionSchedule_EditQty"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Qty"
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblTotalQtyProduce As System.Windows.Forms.Label
    Friend WithEvents DgvWUAllocate_WU As System.Windows.Forms.DataGridView
    Friend WithEvents CoAll_WU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CoAll_WU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_ProdRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Efficiency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PlanRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_ProduceQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lblDeviasi As System.Windows.Forms.Label
End Class
