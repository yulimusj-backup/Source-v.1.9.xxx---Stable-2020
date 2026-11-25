<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_MP_AddMO
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
        Me.btnOK = New System.Windows.Forms.Button
        Me.rboSetDeliveryDate2 = New System.Windows.Forms.RadioButton
        Me.btnCancel = New System.Windows.Forms.Button
        Me.rboNoDeliveryDate = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvCodeNo = New System.Windows.Forms.DataGridView
        Me.ColHeader_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader_QtyOrder = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader_QtyMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader_DefPrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemMO = New System.Windows.Forms.DataGridView
        Me.ColItem_DocType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_RequiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyOrderBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyMo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_DocTypInt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_DefPrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_IsWIP = New System.Windows.Forms.DataGridViewCheckBoxColumn
        CType(Me.dgvCodeNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemMO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(601, 543)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 25)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "Ok"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'rboSetDeliveryDate2
        '
        Me.rboSetDeliveryDate2.AutoSize = True
        Me.rboSetDeliveryDate2.Location = New System.Drawing.Point(11, 58)
        Me.rboSetDeliveryDate2.Name = "rboSetDeliveryDate2"
        Me.rboSetDeliveryDate2.Size = New System.Drawing.Size(319, 21)
        Me.rboSetDeliveryDate2.TabIndex = 10
        Me.rboSetDeliveryDate2.Text = "Differentiate MO for each reference document"
        Me.rboSetDeliveryDate2.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(682, 543)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'rboNoDeliveryDate
        '
        Me.rboNoDeliveryDate.AutoSize = True
        Me.rboNoDeliveryDate.Checked = True
        Me.rboNoDeliveryDate.Location = New System.Drawing.Point(11, 31)
        Me.rboNoDeliveryDate.Name = "rboNoDeliveryDate"
        Me.rboNoDeliveryDate.Size = New System.Drawing.Size(220, 21)
        Me.rboNoDeliveryDate.TabIndex = 7
        Me.rboNoDeliveryDate.TabStop = True
        Me.rboNoDeliveryDate.Text = "Create 1 MO for each code no"
        Me.rboNoDeliveryDate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(244, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Please specify how to specify MO Qty"
        '
        'dgvCodeNo
        '
        Me.dgvCodeNo.AllowUserToAddRows = False
        Me.dgvCodeNo.AllowUserToDeleteRows = False
        Me.dgvCodeNo.AllowUserToResizeColumns = False
        Me.dgvCodeNo.AllowUserToResizeRows = False
        Me.dgvCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCodeNo.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCodeNo.ColumnHeadersHeight = 25
        Me.dgvCodeNo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCodeNo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColHeader_CodeNo, Me.ColHeader_Desc, Me.ColHeader_QtyOrder, Me.ColHeader_QtyMO, Me.ColHeader_Unit, Me.ColHeader_PrimaryCodeNo, Me.ColHeader_DefPrec})
        Me.dgvCodeNo.Location = New System.Drawing.Point(11, 85)
        Me.dgvCodeNo.MultiSelect = False
        Me.dgvCodeNo.Name = "dgvCodeNo"
        Me.dgvCodeNo.ReadOnly = True
        Me.dgvCodeNo.RowHeadersVisible = False
        Me.dgvCodeNo.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCodeNo.RowTemplate.Height = 24
        Me.dgvCodeNo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCodeNo.Size = New System.Drawing.Size(746, 150)
        Me.dgvCodeNo.TabIndex = 12
        '
        'ColHeader_CodeNo
        '
        Me.ColHeader_CodeNo.HeaderText = "Code No"
        Me.ColHeader_CodeNo.Name = "ColHeader_CodeNo"
        Me.ColHeader_CodeNo.ReadOnly = True
        Me.ColHeader_CodeNo.Width = 88
        '
        'ColHeader_Desc
        '
        Me.ColHeader_Desc.HeaderText = "Description"
        Me.ColHeader_Desc.Name = "ColHeader_Desc"
        Me.ColHeader_Desc.ReadOnly = True
        Me.ColHeader_Desc.Width = 104
        '
        'ColHeader_QtyOrder
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColHeader_QtyOrder.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColHeader_QtyOrder.HeaderText = "Qty Order"
        Me.ColHeader_QtyOrder.Name = "ColHeader_QtyOrder"
        Me.ColHeader_QtyOrder.ReadOnly = True
        Me.ColHeader_QtyOrder.Width = 96
        '
        'ColHeader_QtyMO
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColHeader_QtyMO.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColHeader_QtyMO.HeaderText = "Qty MO"
        Me.ColHeader_QtyMO.Name = "ColHeader_QtyMO"
        Me.ColHeader_QtyMO.ReadOnly = True
        Me.ColHeader_QtyMO.Width = 81
        '
        'ColHeader_Unit
        '
        Me.ColHeader_Unit.HeaderText = "Unit"
        Me.ColHeader_Unit.Name = "ColHeader_Unit"
        Me.ColHeader_Unit.ReadOnly = True
        Me.ColHeader_Unit.Width = 58
        '
        'ColHeader_PrimaryCodeNo
        '
        Me.ColHeader_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColHeader_PrimaryCodeNo.Name = "ColHeader_PrimaryCodeNo"
        Me.ColHeader_PrimaryCodeNo.ReadOnly = True
        Me.ColHeader_PrimaryCodeNo.Visible = False
        Me.ColHeader_PrimaryCodeNo.Width = 140
        '
        'ColHeader_DefPrec
        '
        Me.ColHeader_DefPrec.HeaderText = "Def Prec"
        Me.ColHeader_DefPrec.Name = "ColHeader_DefPrec"
        Me.ColHeader_DefPrec.ReadOnly = True
        Me.ColHeader_DefPrec.Visible = False
        Me.ColHeader_DefPrec.Width = 88
        '
        'dgvItemMO
        '
        Me.dgvItemMO.AllowUserToAddRows = False
        Me.dgvItemMO.AllowUserToDeleteRows = False
        Me.dgvItemMO.AllowUserToResizeColumns = False
        Me.dgvItemMO.AllowUserToResizeRows = False
        Me.dgvItemMO.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemMO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemMO.ColumnHeadersHeight = 25
        Me.dgvItemMO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemMO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_DocType, Me.ColItem_DocNo, Me.ColItem_RequiredDate, Me.ColItem_QtyOrderBalance, Me.ColItem_QtyMo, Me.ColItem_Unit, Me.ColItem_DocTypInt, Me.ColItem_DefPrec, Me.ColItem_PrimaryCodeNo, Me.ColItem_IsWIP})
        Me.dgvItemMO.Location = New System.Drawing.Point(11, 241)
        Me.dgvItemMO.MultiSelect = False
        Me.dgvItemMO.Name = "dgvItemMO"
        Me.dgvItemMO.RowHeadersVisible = False
        Me.dgvItemMO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemMO.RowTemplate.Height = 24
        Me.dgvItemMO.Size = New System.Drawing.Size(746, 296)
        Me.dgvItemMO.TabIndex = 13
        '
        'ColItem_DocType
        '
        Me.ColItem_DocType.HeaderText = "Doc Type"
        Me.ColItem_DocType.Name = "ColItem_DocType"
        Me.ColItem_DocType.Width = 94
        '
        'ColItem_DocNo
        '
        Me.ColItem_DocNo.HeaderText = "Document No"
        Me.ColItem_DocNo.Name = "ColItem_DocNo"
        Me.ColItem_DocNo.Width = 119
        '
        'ColItem_RequiredDate
        '
        Me.ColItem_RequiredDate.HeaderText = "Required Date"
        Me.ColItem_RequiredDate.Name = "ColItem_RequiredDate"
        Me.ColItem_RequiredDate.Width = 125
        '
        'ColItem_QtyOrderBalance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyOrderBalance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColItem_QtyOrderBalance.HeaderText = "Qty Order Balance"
        Me.ColItem_QtyOrderBalance.Name = "ColItem_QtyOrderBalance"
        Me.ColItem_QtyOrderBalance.Width = 151
        '
        'ColItem_QtyMo
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyMo.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColItem_QtyMo.HeaderText = "Qty MO"
        Me.ColItem_QtyMo.Name = "ColItem_QtyMo"
        Me.ColItem_QtyMo.Width = 81
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Unit"
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.Width = 58
        '
        'ColItem_DocTypInt
        '
        Me.ColItem_DocTypInt.HeaderText = "Doc Type Int"
        Me.ColItem_DocTypInt.Name = "ColItem_DocTypInt"
        Me.ColItem_DocTypInt.Visible = False
        Me.ColItem_DocTypInt.Width = 113
        '
        'ColItem_DefPrec
        '
        Me.ColItem_DefPrec.HeaderText = "DefPrec"
        Me.ColItem_DefPrec.Name = "ColItem_DefPrec"
        Me.ColItem_DefPrec.Visible = False
        Me.ColItem_DefPrec.Width = 84
        '
        'ColItem_PrimaryCodeNo
        '
        Me.ColItem_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColItem_PrimaryCodeNo.Name = "ColItem_PrimaryCodeNo"
        Me.ColItem_PrimaryCodeNo.Visible = False
        Me.ColItem_PrimaryCodeNo.Width = 140
        '
        'ColItem_IsWIP
        '
        Me.ColItem_IsWIP.HeaderText = "Is WIP"
        Me.ColItem_IsWIP.Name = "ColItem_IsWIP"
        Me.ColItem_IsWIP.Width = 53
        '
        'FrmTrans_MaterialPlanningDashboard_MP_AddMO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 580)
        Me.Controls.Add(Me.dgvItemMO)
        Me.Controls.Add(Me.dgvCodeNo)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.rboSetDeliveryDate2)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.rboNoDeliveryDate)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(713, 617)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_MP_AddMO"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add MO Mode"
        CType(Me.dgvCodeNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemMO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents rboSetDeliveryDate2 As System.Windows.Forms.RadioButton
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents rboNoDeliveryDate As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvCodeNo As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItemMO As System.Windows.Forms.DataGridView
    Friend WithEvents ColHeader_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader_QtyOrder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader_QtyMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader_DefPrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DocType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_RequiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyOrderBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyMo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DocTypInt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DefPrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_IsWIP As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
