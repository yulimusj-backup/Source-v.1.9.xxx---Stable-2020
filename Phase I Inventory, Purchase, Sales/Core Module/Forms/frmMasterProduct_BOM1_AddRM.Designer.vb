<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMasterProduct_BOM1_AddRM
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
        Me.dgvItemBOM = New System.Windows.Forms.DataGridView
        Me.ColBOM_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_GoodPercentage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Tolerance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvItemBOM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItemBOM
        '
        Me.dgvItemBOM.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemBOM.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemBOM.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBOM_CodeNo, Me.ColBOM_Desc, Me.ColBOM_Qty, Me.ColBOM_Unit, Me.ColBOM_GoodPercentage, Me.ColBOM_Tolerance, Me.ColBOM_Remarks, Me.ColBOM_PrimaryProduct})
        Me.dgvItemBOM.Location = New System.Drawing.Point(3, 4)
        Me.dgvItemBOM.MultiSelect = False
        Me.dgvItemBOM.Name = "dgvItemBOM"
        Me.dgvItemBOM.RowHeadersWidth = 25
        Me.dgvItemBOM.RowTemplate.Height = 24
        Me.dgvItemBOM.Size = New System.Drawing.Size(627, 275)
        Me.dgvItemBOM.StandardTab = True
        Me.dgvItemBOM.TabIndex = 0
        '
        'ColBOM_CodeNo
        '
        Me.ColBOM_CodeNo.Frozen = True
        Me.ColBOM_CodeNo.HeaderText = "Code No"
        Me.ColBOM_CodeNo.Name = "ColBOM_CodeNo"
        Me.ColBOM_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_CodeNo.Width = 55
        '
        'ColBOM_Desc
        '
        Me.ColBOM_Desc.HeaderText = "Desc"
        Me.ColBOM_Desc.Name = "ColBOM_Desc"
        Me.ColBOM_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Desc.Width = 38
        '
        'ColBOM_Qty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.NullValue = "1"
        Me.ColBOM_Qty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBOM_Qty.HeaderText = "Qty"
        Me.ColBOM_Qty.Name = "ColBOM_Qty"
        Me.ColBOM_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Qty.Width = 29
        '
        'ColBOM_Unit
        '
        Me.ColBOM_Unit.HeaderText = "Unit"
        Me.ColBOM_Unit.Name = "ColBOM_Unit"
        Me.ColBOM_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Unit.Width = 32
        '
        'ColBOM_GoodPercentage
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.Format = "N2"
        DataGridViewCellStyle2.NullValue = "100.00"
        Me.ColBOM_GoodPercentage.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColBOM_GoodPercentage.HeaderText = "Good Percentage (%)"
        Me.ColBOM_GoodPercentage.Name = "ColBOM_GoodPercentage"
        Me.ColBOM_GoodPercentage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_GoodPercentage.Width = 114
        '
        'ColBOM_Tolerance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0.00"
        Me.ColBOM_Tolerance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBOM_Tolerance.HeaderText = "Tolerance (%)"
        Me.ColBOM_Tolerance.Name = "ColBOM_Tolerance"
        Me.ColBOM_Tolerance.Width = 97
        '
        'ColBOM_Remarks
        '
        Me.ColBOM_Remarks.HeaderText = "Remarks"
        Me.ColBOM_Remarks.Name = "ColBOM_Remarks"
        Me.ColBOM_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Remarks.Width = 55
        '
        'ColBOM_PrimaryProduct
        '
        Me.ColBOM_PrimaryProduct.HeaderText = "Primary Product"
        Me.ColBOM_PrimaryProduct.Name = "ColBOM_PrimaryProduct"
        Me.ColBOM_PrimaryProduct.Visible = False
        Me.ColBOM_PrimaryProduct.Width = 106
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(565, 285)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(65, 21)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(494, 285)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(65, 21)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'frmMasterProduct_BOM1_AddRM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(634, 311)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.dgvItemBOM)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(650, 350)
        Me.Name = "frmMasterProduct_BOM1_AddRM"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Raw Material"
        CType(Me.dgvItemBOM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents dgvItemBOM As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents ColBOM_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_GoodPercentage As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_Tolerance As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColBOM_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
