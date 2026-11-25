<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmEditPriceList
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.txtPrimaryBrand = New System.Windows.Forms.TextBox
        Me.txtPrimaryProductGroup = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnCategory = New System.Windows.Forms.Button
        Me.txtCategory = New System.Windows.Forms.TextBox
        Me.btnBrand = New System.Windows.Forms.Button
        Me.txtBrand = New System.Windows.Forms.TextBox
        Me.btnProductGroup = New System.Windows.Forms.Button
        Me.txtProductGroup = New System.Windows.Forms.TextBox
        Me.rboCategory = New System.Windows.Forms.RadioButton
        Me.rboBrand = New System.Windows.Forms.RadioButton
        Me.rboProductGroup = New System.Windows.Forms.RadioButton
        Me.rboAll = New System.Windows.Forms.RadioButton
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.colPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryCategory)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryBrand)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryProductGroup)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryCurrency)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Controls.Add(Me.btnFilter)
        Me.GroupBox1.Controls.Add(Me.btnCategory)
        Me.GroupBox1.Controls.Add(Me.txtCategory)
        Me.GroupBox1.Controls.Add(Me.btnBrand)
        Me.GroupBox1.Controls.Add(Me.txtBrand)
        Me.GroupBox1.Controls.Add(Me.btnProductGroup)
        Me.GroupBox1.Controls.Add(Me.txtProductGroup)
        Me.GroupBox1.Controls.Add(Me.rboCategory)
        Me.GroupBox1.Controls.Add(Me.rboBrand)
        Me.GroupBox1.Controls.Add(Me.rboProductGroup)
        Me.GroupBox1.Controls.Add(Me.rboAll)
        Me.GroupBox1.Controls.Add(Me.btnFindCurrency)
        Me.GroupBox1.Controls.Add(Me.txtCurrencyCode)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(899, 94)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(816, 40)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(74, 22)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(494, 67)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(40, 20)
        Me.txtPrimaryCategory.TabIndex = 22
        Me.txtPrimaryCategory.Visible = False
        '
        'txtPrimaryBrand
        '
        Me.txtPrimaryBrand.Location = New System.Drawing.Point(494, 40)
        Me.txtPrimaryBrand.Name = "txtPrimaryBrand"
        Me.txtPrimaryBrand.Size = New System.Drawing.Size(40, 20)
        Me.txtPrimaryBrand.TabIndex = 21
        Me.txtPrimaryBrand.Visible = False
        '
        'txtPrimaryProductGroup
        '
        Me.txtPrimaryProductGroup.Location = New System.Drawing.Point(109, 63)
        Me.txtPrimaryProductGroup.Name = "txtPrimaryProductGroup"
        Me.txtPrimaryProductGroup.Size = New System.Drawing.Size(40, 20)
        Me.txtPrimaryProductGroup.TabIndex = 20
        Me.txtPrimaryProductGroup.Visible = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(109, 16)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(40, 20)
        Me.txtPrimaryCurrency.TabIndex = 19
        Me.txtPrimaryCurrency.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(736, 66)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(74, 22)
        Me.btnSave.TabIndex = 13
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(736, 40)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(74, 22)
        Me.btnFilter.TabIndex = 12
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnCategory
        '
        Me.btnCategory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCategory.Location = New System.Drawing.Point(698, 66)
        Me.btnCategory.Name = "btnCategory"
        Me.btnCategory.Size = New System.Drawing.Size(32, 20)
        Me.btnCategory.TabIndex = 11
        Me.btnCategory.TabStop = False
        Me.btnCategory.Text = "...."
        Me.btnCategory.UseVisualStyleBackColor = True
        '
        'txtCategory
        '
        Me.txtCategory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCategory.Location = New System.Drawing.Point(494, 67)
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.ReadOnly = True
        Me.txtCategory.Size = New System.Drawing.Size(198, 20)
        Me.txtCategory.TabIndex = 10
        '
        'btnBrand
        '
        Me.btnBrand.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrand.Location = New System.Drawing.Point(698, 40)
        Me.btnBrand.Name = "btnBrand"
        Me.btnBrand.Size = New System.Drawing.Size(32, 20)
        Me.btnBrand.TabIndex = 8
        Me.btnBrand.TabStop = False
        Me.btnBrand.Text = "...."
        Me.btnBrand.UseVisualStyleBackColor = True
        '
        'txtBrand
        '
        Me.txtBrand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBrand.Location = New System.Drawing.Point(494, 41)
        Me.txtBrand.Name = "txtBrand"
        Me.txtBrand.ReadOnly = True
        Me.txtBrand.Size = New System.Drawing.Size(198, 20)
        Me.txtBrand.TabIndex = 7
        '
        'btnProductGroup
        '
        Me.btnProductGroup.Location = New System.Drawing.Point(313, 65)
        Me.btnProductGroup.Name = "btnProductGroup"
        Me.btnProductGroup.Size = New System.Drawing.Size(32, 20)
        Me.btnProductGroup.TabIndex = 5
        Me.btnProductGroup.TabStop = False
        Me.btnProductGroup.Text = "...."
        Me.btnProductGroup.UseVisualStyleBackColor = True
        '
        'txtProductGroup
        '
        Me.txtProductGroup.Location = New System.Drawing.Point(109, 65)
        Me.txtProductGroup.Name = "txtProductGroup"
        Me.txtProductGroup.ReadOnly = True
        Me.txtProductGroup.Size = New System.Drawing.Size(198, 20)
        Me.txtProductGroup.TabIndex = 4
        '
        'rboCategory
        '
        Me.rboCategory.AutoSize = True
        Me.rboCategory.Location = New System.Drawing.Point(421, 68)
        Me.rboCategory.Name = "rboCategory"
        Me.rboCategory.Size = New System.Drawing.Size(67, 17)
        Me.rboCategory.TabIndex = 9
        Me.rboCategory.TabStop = True
        Me.rboCategory.Text = "Category"
        Me.rboCategory.UseVisualStyleBackColor = True
        '
        'rboBrand
        '
        Me.rboBrand.AutoSize = True
        Me.rboBrand.Location = New System.Drawing.Point(421, 43)
        Me.rboBrand.Name = "rboBrand"
        Me.rboBrand.Size = New System.Drawing.Size(53, 17)
        Me.rboBrand.TabIndex = 6
        Me.rboBrand.TabStop = True
        Me.rboBrand.Text = "Brand"
        Me.rboBrand.UseVisualStyleBackColor = True
        '
        'rboProductGroup
        '
        Me.rboProductGroup.AutoSize = True
        Me.rboProductGroup.Location = New System.Drawing.Point(9, 66)
        Me.rboProductGroup.Name = "rboProductGroup"
        Me.rboProductGroup.Size = New System.Drawing.Size(94, 17)
        Me.rboProductGroup.TabIndex = 3
        Me.rboProductGroup.TabStop = True
        Me.rboProductGroup.Text = "Product Group"
        Me.rboProductGroup.UseVisualStyleBackColor = True
        '
        'rboAll
        '
        Me.rboAll.AutoSize = True
        Me.rboAll.Checked = True
        Me.rboAll.Location = New System.Drawing.Point(9, 43)
        Me.rboAll.Name = "rboAll"
        Me.rboAll.Size = New System.Drawing.Size(36, 17)
        Me.rboAll.TabIndex = 2
        Me.rboAll.TabStop = True
        Me.rboAll.Text = "All"
        Me.rboAll.UseVisualStyleBackColor = True
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(313, 16)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(32, 20)
        Me.btnFindCurrency.TabIndex = 1
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "...."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.Location = New System.Drawing.Point(109, 16)
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(198, 20)
        Me.txtCurrencyCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Currency"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 112)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(899, 316)
        Me.DataGridView1.TabIndex = 1
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(918, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AllowUserToResizeRows = False
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colPrimaryProduct, Me.colUnitPrice, Me.colDiscount})
        Me.DataGridView2.Location = New System.Drawing.Point(409, 164)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(500, 220)
        Me.DataGridView2.TabIndex = 3
        Me.DataGridView2.Visible = False
        '
        'colPrimaryProduct
        '
        Me.colPrimaryProduct.HeaderText = "PrimaryProduct"
        Me.colPrimaryProduct.Name = "colPrimaryProduct"
        '
        'colUnitPrice
        '
        Me.colUnitPrice.HeaderText = "UnitPrice"
        Me.colUnitPrice.Name = "colUnitPrice"
        '
        'colDiscount
        '
        Me.colDiscount.HeaderText = "Discount"
        Me.colDiscount.Name = "colDiscount"
        '
        'FrmEditPriceList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(918, 453)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(699, 489)
        Me.Name = "FrmEditPriceList"
        Me.ShowInTaskbar = False
        Me.Text = "Edit Price List"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnCategory As System.Windows.Forms.Button
    Friend WithEvents txtCategory As System.Windows.Forms.TextBox
    Friend WithEvents btnBrand As System.Windows.Forms.Button
    Friend WithEvents txtBrand As System.Windows.Forms.TextBox
    Friend WithEvents btnProductGroup As System.Windows.Forms.Button
    Friend WithEvents txtProductGroup As System.Windows.Forms.TextBox
    Friend WithEvents rboCategory As System.Windows.Forms.RadioButton
    Friend WithEvents rboBrand As System.Windows.Forms.RadioButton
    Friend WithEvents rboProductGroup As System.Windows.Forms.RadioButton
    Friend WithEvents rboAll As System.Windows.Forms.RadioButton
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryBrand As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryProductGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents colPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
