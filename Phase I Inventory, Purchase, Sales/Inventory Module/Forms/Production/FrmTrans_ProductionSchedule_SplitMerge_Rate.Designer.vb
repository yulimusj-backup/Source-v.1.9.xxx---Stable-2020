<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_SplitMerge_Rate
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.Label1 = New System.Windows.Forms.Label
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.btnOK = New System.Windows.Forms.Button
        Me.ColMoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colQtyOri = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate_Ori = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWH_Ori = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRateEfficiency_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRateEffective_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWH_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHide_Rate_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHide_Rate_Efficiency_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHide_Rate_Effective_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHide_WH_New = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimalPoint = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMoNo, Me.ColJONo, Me.ColProcessCode, Me.ColWONo, Me.ColCodeNo, Me.ColDesc, Me.colQtyOri, Me.ColQty, Me.ColUnit, Me.ColRate_Ori, Me.ColWH_Ori, Me.ColRate_New, Me.ColRateEfficiency_New, Me.ColRateEffective_New, Me.ColWH_New, Me.ColPrimaryMO, Me.ColIdentifierJO, Me.ColPrimaryProcess, Me.ColIdentifierWO, Me.ColPrimaryProduct, Me.ColPrimaryUnit, Me.ColHide_Rate_New, Me.ColHide_Rate_Efficiency_New, Me.ColHide_Rate_Effective_New, Me.ColHide_WH_New, Me.ColDecimalPoint})
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvItem.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvItem.Location = New System.Drawing.Point(6, 88)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(859, 418)
        Me.dgvItem.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(439, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "The following work unit item has different rate on target work unit, please decid" & _
            "e what to do:"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(11, 24)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(188, 17)
        Me.RadioButton1.TabIndex = 14
        Me.RadioButton1.Text = "Keep using current production rate"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Location = New System.Drawing.Point(11, 46)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(310, 17)
        Me.RadioButton2.TabIndex = 15
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Update all using standard production rate on target work unit"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(11, 67)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(261, 17)
        Me.RadioButton3.TabIndex = 16
        Me.RadioButton3.Text = "Let me decide for each work order item individualy"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(811, 64)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 21)
        Me.btnOK.TabIndex = 24
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ColMoNo
        '
        Me.ColMoNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ColMoNo.Frozen = True
        Me.ColMoNo.HeaderText = "Mo No"
        Me.ColMoNo.Name = "ColMoNo"
        Me.ColMoNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMoNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMoNo.Width = 70
        '
        'ColJONo
        '
        Me.ColJONo.Frozen = True
        Me.ColJONo.HeaderText = "JO No"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColJONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJONo.Width = 43
        '
        'ColProcessCode
        '
        Me.ColProcessCode.Frozen = True
        Me.ColProcessCode.HeaderText = "Process"
        Me.ColProcessCode.Name = "ColProcessCode"
        Me.ColProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProcessCode.Width = 51
        '
        'ColWONo
        '
        Me.ColWONo.Frozen = True
        Me.ColWONo.HeaderText = "WO No"
        Me.ColWONo.Name = "ColWONo"
        Me.ColWONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWONo.Width = 49
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 55
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'colQtyOri
        '
        Me.colQtyOri.HeaderText = "Qty Ori"
        Me.colQtyOri.Name = "colQtyOri"
        Me.colQtyOri.Visible = False
        Me.colQtyOri.Width = 64
        '
        'ColQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 29
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 32
        '
        'ColRate_Ori
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate_Ori.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColRate_Ori.HeaderText = "Rate"
        Me.ColRate_Ori.Name = "ColRate_Ori"
        Me.ColRate_Ori.ReadOnly = True
        Me.ColRate_Ori.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRate_Ori.Width = 36
        '
        'ColWH_Ori
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWH_Ori.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColWH_Ori.HeaderText = "W/H"
        Me.ColWH_Ori.Name = "ColWH_Ori"
        Me.ColWH_Ori.ReadOnly = True
        Me.ColWH_Ori.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWH_Ori.Width = 37
        '
        'ColRate_New
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate_New.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColRate_New.HeaderText = "Rate New"
        Me.ColRate_New.Name = "ColRate_New"
        Me.ColRate_New.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRate_New.Width = 61
        '
        'ColRateEfficiency_New
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRateEfficiency_New.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRateEfficiency_New.HeaderText = "Rate Efficiency New"
        Me.ColRateEfficiency_New.Name = "ColRateEfficiency_New"
        Me.ColRateEfficiency_New.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRateEfficiency_New.Width = 110
        '
        'ColRateEffective_New
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRateEffective_New.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColRateEffective_New.HeaderText = "Rate Effective New"
        Me.ColRateEffective_New.Name = "ColRateEffective_New"
        Me.ColRateEffective_New.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRateEffective_New.Width = 106
        '
        'ColWH_New
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWH_New.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColWH_New.HeaderText = "W/H New"
        Me.ColWH_New.Name = "ColWH_New"
        Me.ColWH_New.ReadOnly = True
        Me.ColWH_New.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWH_New.Width = 62
        '
        'ColPrimaryMO
        '
        Me.ColPrimaryMO.HeaderText = "Primary MO"
        Me.ColPrimaryMO.Name = "ColPrimaryMO"
        Me.ColPrimaryMO.ReadOnly = True
        Me.ColPrimaryMO.Visible = False
        Me.ColPrimaryMO.Width = 86
        '
        'ColIdentifierJO
        '
        Me.ColIdentifierJO.HeaderText = "IdentifierJO"
        Me.ColIdentifierJO.Name = "ColIdentifierJO"
        Me.ColIdentifierJO.Visible = False
        Me.ColIdentifierJO.Width = 85
        '
        'ColPrimaryProcess
        '
        Me.ColPrimaryProcess.HeaderText = "Primary Process"
        Me.ColPrimaryProcess.Name = "ColPrimaryProcess"
        Me.ColPrimaryProcess.ReadOnly = True
        Me.ColPrimaryProcess.Visible = False
        Me.ColPrimaryProcess.Width = 107
        '
        'ColIdentifierWO
        '
        Me.ColIdentifierWO.HeaderText = "IdentifierWO"
        Me.ColIdentifierWO.Name = "ColIdentifierWO"
        Me.ColIdentifierWO.Visible = False
        Me.ColIdentifierWO.Width = 91
        '
        'ColPrimaryProduct
        '
        Me.ColPrimaryProduct.HeaderText = "Primary Product"
        Me.ColPrimaryProduct.Name = "ColPrimaryProduct"
        Me.ColPrimaryProduct.ReadOnly = True
        Me.ColPrimaryProduct.Visible = False
        Me.ColPrimaryProduct.Width = 106
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 88
        '
        'ColHide_Rate_New
        '
        Me.ColHide_Rate_New.HeaderText = "Hide Rate New"
        Me.ColHide_Rate_New.Name = "ColHide_Rate_New"
        Me.ColHide_Rate_New.ReadOnly = True
        Me.ColHide_Rate_New.Visible = False
        Me.ColHide_Rate_New.Width = 105
        '
        'ColHide_Rate_Efficiency_New
        '
        Me.ColHide_Rate_Efficiency_New.HeaderText = "Hide Rate Efficiency New"
        Me.ColHide_Rate_Efficiency_New.Name = "ColHide_Rate_Efficiency_New"
        Me.ColHide_Rate_Efficiency_New.ReadOnly = True
        Me.ColHide_Rate_Efficiency_New.Visible = False
        Me.ColHide_Rate_Efficiency_New.Width = 154
        '
        'ColHide_Rate_Effective_New
        '
        Me.ColHide_Rate_Effective_New.HeaderText = "Hide Rate Effective"
        Me.ColHide_Rate_Effective_New.Name = "ColHide_Rate_Effective_New"
        Me.ColHide_Rate_Effective_New.ReadOnly = True
        Me.ColHide_Rate_Effective_New.Visible = False
        Me.ColHide_Rate_Effective_New.Width = 125
        '
        'ColHide_WH_New
        '
        Me.ColHide_WH_New.HeaderText = "Hide W/H New"
        Me.ColHide_WH_New.Name = "ColHide_WH_New"
        Me.ColHide_WH_New.ReadOnly = True
        Me.ColHide_WH_New.Visible = False
        Me.ColHide_WH_New.Width = 106
        '
        'ColDecimalPoint
        '
        Me.ColDecimalPoint.HeaderText = "Decimal Point"
        Me.ColDecimalPoint.Name = "ColDecimalPoint"
        Me.ColDecimalPoint.ReadOnly = True
        Me.ColDecimalPoint.Visible = False
        Me.ColDecimalPoint.Width = 97
        '
        'FrmTrans_ProductionSchedule_Split_Rate
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(871, 513)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.RadioButton3)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmTrans_ProductionSchedule_Split_Rate"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Rate"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColMoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colQtyOri As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate_Ori As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWH_Ori As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRateEfficiency_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRateEffective_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWH_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHide_Rate_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHide_Rate_Efficiency_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHide_Rate_Effective_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHide_WH_New As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimalPoint As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
