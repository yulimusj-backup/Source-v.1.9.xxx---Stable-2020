<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainSOBDelivery
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtOrderQtyMulti = New System.Windows.Forms.TextBox
        Me.TxtOrderQty = New System.Windows.Forms.TextBox
        Me.LblOrderUnitMulti = New System.Windows.Forms.Label
        Me.LblUnitCode = New System.Windows.Forms.Label
        Me.LblUnitCodeOrderBalance = New System.Windows.Forms.Label
        Me.LblOrderBalanceUnitMulti = New System.Windows.Forms.Label
        Me.TxtSOB = New System.Windows.Forms.TextBox
        Me.TxtSOBMulti = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.LblRatio = New System.Windows.Forms.Label
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyBalanceMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPlusPlus = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPriorityLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColPrimaryPriority = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyBalanceDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatio = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(582, 466)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'dgvDetail
        '
        Me.dgvDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.ColumnHeadersHeight = 25
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDeliveryDate, Me.ColQtyMulti, Me.ColQtyBalanceMulti, Me.ColUnitMulti, Me.ColPlusPlus, Me.ColPriorityLevel, Me.ColPrimaryPriority, Me.ColQty, Me.ColQtyBalanceDefault, Me.ColUnit, Me.ColRatio, Me.ColRemarks})
        Me.dgvDetail.Location = New System.Drawing.Point(12, 86)
        Me.dgvDetail.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.RowTemplate.Height = 24
        Me.dgvDetail.Size = New System.Drawing.Size(655, 372)
        Me.dgvDetail.TabIndex = 6
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(500, 466)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Order Qty"
        '
        'TxtOrderQtyMulti
        '
        Me.TxtOrderQtyMulti.Location = New System.Drawing.Point(157, 12)
        Me.TxtOrderQtyMulti.Name = "TxtOrderQtyMulti"
        Me.TxtOrderQtyMulti.ReadOnly = True
        Me.TxtOrderQtyMulti.Size = New System.Drawing.Size(109, 22)
        Me.TxtOrderQtyMulti.TabIndex = 8
        Me.TxtOrderQtyMulti.Text = "0"
        Me.TxtOrderQtyMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtOrderQty
        '
        Me.TxtOrderQty.Location = New System.Drawing.Point(369, 12)
        Me.TxtOrderQty.Name = "TxtOrderQty"
        Me.TxtOrderQty.ReadOnly = True
        Me.TxtOrderQty.Size = New System.Drawing.Size(109, 22)
        Me.TxtOrderQty.TabIndex = 9
        Me.TxtOrderQty.Text = "0"
        Me.TxtOrderQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblOrderUnitMulti
        '
        Me.LblOrderUnitMulti.AutoSize = True
        Me.LblOrderUnitMulti.Location = New System.Drawing.Point(272, 15)
        Me.LblOrderUnitMulti.Name = "LblOrderUnitMulti"
        Me.LblOrderUnitMulti.Size = New System.Drawing.Size(82, 17)
        Me.LblOrderUnitMulti.TabIndex = 10
        Me.LblOrderUnitMulti.Text = "Unit Multi = "
        '
        'LblUnitCode
        '
        Me.LblUnitCode.AutoSize = True
        Me.LblUnitCode.Location = New System.Drawing.Point(484, 15)
        Me.LblUnitCode.Name = "LblUnitCode"
        Me.LblUnitCode.Size = New System.Drawing.Size(33, 17)
        Me.LblUnitCode.TabIndex = 11
        Me.LblUnitCode.Text = "Unit"
        '
        'LblUnitCodeOrderBalance
        '
        Me.LblUnitCodeOrderBalance.AutoSize = True
        Me.LblUnitCodeOrderBalance.Location = New System.Drawing.Point(484, 43)
        Me.LblUnitCodeOrderBalance.Name = "LblUnitCodeOrderBalance"
        Me.LblUnitCodeOrderBalance.Size = New System.Drawing.Size(33, 17)
        Me.LblUnitCodeOrderBalance.TabIndex = 16
        Me.LblUnitCodeOrderBalance.Text = "Unit"
        '
        'LblOrderBalanceUnitMulti
        '
        Me.LblOrderBalanceUnitMulti.AutoSize = True
        Me.LblOrderBalanceUnitMulti.Location = New System.Drawing.Point(272, 43)
        Me.LblOrderBalanceUnitMulti.Name = "LblOrderBalanceUnitMulti"
        Me.LblOrderBalanceUnitMulti.Size = New System.Drawing.Size(82, 17)
        Me.LblOrderBalanceUnitMulti.TabIndex = 15
        Me.LblOrderBalanceUnitMulti.Text = "Unit Multi = "
        '
        'TxtSOB
        '
        Me.TxtSOB.Location = New System.Drawing.Point(369, 40)
        Me.TxtSOB.Name = "TxtSOB"
        Me.TxtSOB.ReadOnly = True
        Me.TxtSOB.Size = New System.Drawing.Size(109, 22)
        Me.TxtSOB.TabIndex = 14
        Me.TxtSOB.Text = "0"
        Me.TxtSOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtSOBMulti
        '
        Me.TxtSOBMulti.Location = New System.Drawing.Point(157, 40)
        Me.TxtSOBMulti.Name = "TxtSOBMulti"
        Me.TxtSOBMulti.ReadOnly = True
        Me.TxtSOBMulti.Size = New System.Drawing.Size(109, 22)
        Me.TxtSOBMulti.TabIndex = 13
        Me.TxtSOBMulti.Text = "0"
        Me.TxtSOBMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(139, 17)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Sales Order Balance"
        '
        'LblRatio
        '
        Me.LblRatio.AutoSize = True
        Me.LblRatio.Location = New System.Drawing.Point(154, 67)
        Me.LblRatio.Name = "LblRatio"
        Me.LblRatio.Size = New System.Drawing.Size(60, 17)
        Me.LblRatio.TabIndex = 17
        Me.LblRatio.Text = "LblRatio"
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDeliveryDate.Width = 99
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyMulti.Width = 36
        '
        'ColQtyBalanceMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyBalanceMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyBalanceMulti.HeaderText = "Qty Balance"
        Me.ColQtyBalanceMulti.Name = "ColQtyBalanceMulti"
        Me.ColQtyBalanceMulti.Width = 110
        '
        'ColUnitMulti
        '
        Me.ColUnitMulti.HeaderText = "Unit"
        Me.ColUnitMulti.Name = "ColUnitMulti"
        Me.ColUnitMulti.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitMulti.Width = 39
        '
        'ColPlusPlus
        '
        Me.ColPlusPlus.HeaderText = "Set"
        Me.ColPlusPlus.Name = "ColPlusPlus"
        Me.ColPlusPlus.Text = ""
        Me.ColPlusPlus.Width = 35
        '
        'ColPriorityLevel
        '
        Me.ColPriorityLevel.HeaderText = "Priority"
        Me.ColPriorityLevel.Name = "ColPriorityLevel"
        Me.ColPriorityLevel.Width = 58
        '
        'ColPrimaryPriority
        '
        Me.ColPrimaryPriority.HeaderText = "Primary Priority"
        Me.ColPrimaryPriority.Name = "ColPrimaryPriority"
        Me.ColPrimaryPriority.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimaryPriority.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryPriority.Visible = False
        Me.ColPrimaryPriority.Width = 110
        '
        'ColQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQty.HeaderText = "Qty Default"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Visible = False
        Me.ColQty.Width = 85
        '
        'ColQtyBalanceDefault
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyBalanceDefault.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQtyBalanceDefault.HeaderText = "Qty Balance"
        Me.ColQtyBalanceDefault.Name = "ColQtyBalanceDefault"
        Me.ColQtyBalanceDefault.Width = 110
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Visible = False
        Me.ColUnit.Width = 88
        '
        'ColRatio
        '
        Me.ColRatio.HeaderText = "Ratio"
        Me.ColRatio.Name = "ColRatio"
        Me.ColRatio.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatio.Visible = False
        Me.ColRatio.Width = 47
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainSOBDelivery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 502)
        Me.Controls.Add(Me.LblRatio)
        Me.Controls.Add(Me.LblUnitCodeOrderBalance)
        Me.Controls.Add(Me.LblOrderBalanceUnitMulti)
        Me.Controls.Add(Me.TxtSOB)
        Me.Controls.Add(Me.TxtSOBMulti)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LblUnitCode)
        Me.Controls.Add(Me.LblOrderUnitMulti)
        Me.Controls.Add(Me.TxtOrderQty)
        Me.Controls.Add(Me.TxtOrderQtyMulti)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.btnOK)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(687, 539)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainSOBDelivery"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Delivery Detail"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtOrderQtyMulti As System.Windows.Forms.TextBox
    Friend WithEvents TxtOrderQty As System.Windows.Forms.TextBox
    Friend WithEvents LblOrderUnitMulti As System.Windows.Forms.Label
    Friend WithEvents LblUnitCode As System.Windows.Forms.Label
    Friend WithEvents LblUnitCodeOrderBalance As System.Windows.Forms.Label
    Friend WithEvents LblOrderBalanceUnitMulti As System.Windows.Forms.Label
    Friend WithEvents TxtSOB As System.Windows.Forms.TextBox
    Friend WithEvents TxtSOBMulti As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblRatio As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyBalanceMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPlusPlus As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPriorityLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColPrimaryPriority As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyBalanceDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
