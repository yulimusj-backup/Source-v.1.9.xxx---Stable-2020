<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_ProcessOrder
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
        Me.DgvProsesOrder = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProcessName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUrutan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnReorderWO = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        CType(Me.DgvProsesOrder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvProsesOrder
        '
        Me.DgvProsesOrder.AllowUserToAddRows = False
        Me.DgvProsesOrder.AllowUserToDeleteRows = False
        Me.DgvProsesOrder.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvProsesOrder.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvProsesOrder.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvProsesOrder.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvProsesOrder.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colProcessCode, Me.colProcessName, Me.colPrimary, Me.colPrimaryProduct, Me.colUrutan})
        Me.DgvProsesOrder.Location = New System.Drawing.Point(13, 15)
        Me.DgvProsesOrder.Margin = New System.Windows.Forms.Padding(4)
        Me.DgvProsesOrder.MultiSelect = False
        Me.DgvProsesOrder.Name = "DgvProsesOrder"
        Me.DgvProsesOrder.ReadOnly = True
        Me.DgvProsesOrder.Size = New System.Drawing.Size(739, 335)
        Me.DgvProsesOrder.TabIndex = 0
        '
        'colNo
        '
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 51
        '
        'colProcessCode
        '
        Me.colProcessCode.HeaderText = "Process Code"
        Me.colProcessCode.Name = "colProcessCode"
        Me.colProcessCode.ReadOnly = True
        Me.colProcessCode.Width = 121
        '
        'colProcessName
        '
        Me.colProcessName.HeaderText = "Process Name"
        Me.colProcessName.Name = "colProcessName"
        Me.colProcessName.ReadOnly = True
        Me.colProcessName.Width = 125
        '
        'colPrimary
        '
        Me.colPrimary.HeaderText = "Primary"
        Me.colPrimary.Name = "colPrimary"
        Me.colPrimary.ReadOnly = True
        Me.colPrimary.Width = 81
        '
        'colPrimaryProduct
        '
        Me.colPrimaryProduct.HeaderText = "Primary Product"
        Me.colPrimaryProduct.Name = "colPrimaryProduct"
        Me.colPrimaryProduct.ReadOnly = True
        Me.colPrimaryProduct.Width = 134
        '
        'colUrutan
        '
        Me.colUrutan.HeaderText = "Urutan"
        Me.colUrutan.Name = "colUrutan"
        Me.colUrutan.ReadOnly = True
        Me.colUrutan.Width = 76
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(16, 357)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 32
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(61, 357)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 33
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(649, 361)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(100, 28)
        Me.BtnOK.TabIndex = 34
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnReorderWO
        '
        Me.BtnReorderWO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnReorderWO.Location = New System.Drawing.Point(541, 361)
        Me.BtnReorderWO.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnReorderWO.Name = "BtnReorderWO"
        Me.BtnReorderWO.Size = New System.Drawing.Size(100, 28)
        Me.BtnReorderWO.TabIndex = 35
        Me.BtnReorderWO.Text = "Reorder WO"
        Me.BtnReorderWO.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(652, 321)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 28)
        Me.BtnClose.TabIndex = 36
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_ManufacturingOrder_ProcessOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(767, 406)
        Me.Controls.Add(Me.BtnReorderWO)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.btnUP)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.DgvProsesOrder)
        Me.Controls.Add(Me.BtnClose)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(495, 443)
        Me.Name = "FrmTrans_ManufacturingOrder_ProcessOrder"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Process Order"
        CType(Me.DgvProsesOrder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvProsesOrder As System.Windows.Forms.DataGridView
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnReorderWO As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcessName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUrutan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
