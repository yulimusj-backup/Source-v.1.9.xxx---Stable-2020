<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakStockUnderMinimumQuantity
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
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.BtnCetak = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BgwProduct = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(3, 122)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.ShowRefreshButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(728, 336)
        Me.CrystalReportViewer1.TabIndex = 1
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvProduct)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.BtnCetak)
        Me.GroupBox1.Location = New System.Drawing.Point(3, -3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(728, 119)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'dgvProduct
        '
        Me.dgvProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProduct.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProduct_No, Me.colProduct_GroupLevel, Me.colProduct_Code, Me.colProduct_Desc, Me.colProduct_Primary})
        Me.dgvProduct.Location = New System.Drawing.Point(122, 7)
        Me.dgvProduct.Name = "dgvProduct"
        Me.dgvProduct.RowTemplate.Height = 24
        Me.dgvProduct.Size = New System.Drawing.Size(600, 110)
        Me.dgvProduct.TabIndex = 7
        '
        'colProduct_No
        '
        Me.colProduct_No.HeaderText = "No"
        Me.colProduct_No.Name = "colProduct_No"
        Me.colProduct_No.Width = 46
        '
        'colProduct_GroupLevel
        '
        Me.colProduct_GroupLevel.HeaderText = "Group Level"
        Me.colProduct_GroupLevel.Items.AddRange(New Object() {"Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.colProduct_GroupLevel.Name = "colProduct_GroupLevel"
        Me.colProduct_GroupLevel.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colProduct_GroupLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colProduct_GroupLevel.Width = 90
        '
        'colProduct_Code
        '
        Me.colProduct_Code.HeaderText = "Code"
        Me.colProduct_Code.Name = "colProduct_Code"
        Me.colProduct_Code.Width = 57
        '
        'colProduct_Desc
        '
        Me.colProduct_Desc.HeaderText = "Desc"
        Me.colProduct_Desc.Name = "colProduct_Desc"
        Me.colProduct_Desc.ReadOnly = True
        Me.colProduct_Desc.Width = 57
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 66
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(9, 38)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(99, 17)
        Me.RadioButton2.TabIndex = 6
        Me.RadioButton2.Text = "Filtered Product"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(9, 15)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(76, 17)
        Me.RadioButton1.TabIndex = 5
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "All Product"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'BtnCetak
        '
        Me.BtnCetak.Location = New System.Drawing.Point(47, 61)
        Me.BtnCetak.Name = "BtnCetak"
        Me.BtnCetak.Size = New System.Drawing.Size(69, 21)
        Me.BtnCetak.TabIndex = 4
        Me.BtnCetak.Text = "&Preview"
        Me.BtnCetak.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(477, 170)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(48, 20)
        Me.BtnClose.TabIndex = 14
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BgwProduct
        '
        '
        'FrmCetakStockUnderMinimumQuantity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(734, 461)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(750, 500)
        Me.Name = "FrmCetakStockUnderMinimumQuantity"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Under Minimum Quantity Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnCetak As System.Windows.Forms.Button
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwProduct As System.ComponentModel.BackgroundWorker
End Class
