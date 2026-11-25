<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Data1_Exception
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
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.Label1 = New System.Windows.Forms.Label
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGrossMargin = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AllowUserToResizeColumns = False
        Me.DataGridView2.AllowUserToResizeRows = False
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView2.ColumnHeadersHeight = 25
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColUnit, Me.ColTotalQty, Me.ColSalesAmount, Me.ColGrossMargin})
        Me.DataGridView2.Location = New System.Drawing.Point(13, 39)
        Me.DataGridView2.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView2.RowTemplate.Height = 24
        Me.DataGridView2.Size = New System.Drawing.Size(866, 561)
        Me.DataGridView2.TabIndex = 145
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(412, 17)
        Me.Label1.TabIndex = 146
        Me.Label1.Text = "The following code no has negative qty, please be aware of this."
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 88
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.Width = 104
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.Width = 58
        '
        'ColTotalQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColTotalQty.HeaderText = "Total Qty"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.ReadOnly = True
        Me.ColTotalQty.Width = 91
        '
        'ColSalesAmount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesAmount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSalesAmount.HeaderText = "Nett Sales Amount"
        Me.ColSalesAmount.Name = "ColSalesAmount"
        Me.ColSalesAmount.ReadOnly = True
        Me.ColSalesAmount.Width = 150
        '
        'ColGrossMargin
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColGrossMargin.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColGrossMargin.HeaderText = "Gross Margin"
        Me.ColGrossMargin.Name = "ColGrossMargin"
        Me.ColGrossMargin.ReadOnly = True
        Me.ColGrossMargin.Width = 118
        '
        'FrmTrans_MaterialPlanningDashboard_Data1_Exception
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 613)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 650)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Data1_Exception"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exception"
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGrossMargin As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
