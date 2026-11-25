<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainSOB
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
        Me.dgvData = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustPONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOrderQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSpareQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryDateCont = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOrderQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AllowUserToResizeColumns = False
        Me.dgvData.AllowUserToResizeRows = False
        Me.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvData.ColumnHeadersHeight = 25
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColSONumber, Me.ColDate, Me.ColCustomer, Me.ColCustPONumber, Me.ColOrderQty, Me.ColSpareQty, Me.ColTotalQty, Me.ColUnit, Me.ColPrimarySO, Me.ColDeliveryDateCont, Me.ColDeliveryDate, Me.ColIdentifier, Me.ColPrimaryUnit, Me.ColPrimaryUnitMulti, Me.ColRatioDefault, Me.ColRatioReference, Me.ColOrderQtyMulti, Me.ColPrimaryCust})
        Me.dgvData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvData.Location = New System.Drawing.Point(0, 0)
        Me.dgvData.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.RowHeadersVisible = False
        Me.dgvData.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.Size = New System.Drawing.Size(891, 491)
        Me.dgvData.TabIndex = 4
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColSONumber
        '
        Me.ColSONumber.HeaderText = "SO Number"
        Me.ColSONumber.Name = "ColSONumber"
        Me.ColSONumber.ReadOnly = True
        Me.ColSONumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSONumber.Width = 88
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 44
        '
        'ColCustomer
        '
        Me.ColCustomer.HeaderText = "Customer"
        Me.ColCustomer.Name = "ColCustomer"
        Me.ColCustomer.ReadOnly = True
        Me.ColCustomer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustomer.Width = 74
        '
        'ColCustPONumber
        '
        Me.ColCustPONumber.HeaderText = "Customer PO Number"
        Me.ColCustPONumber.Name = "ColCustPONumber"
        Me.ColCustPONumber.ReadOnly = True
        Me.ColCustPONumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustPONumber.Width = 152
        '
        'ColOrderQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOrderQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColOrderQty.HeaderText = "Balance Qty"
        Me.ColOrderQty.Name = "ColOrderQty"
        Me.ColOrderQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOrderQty.Width = 91
        '
        'ColSpareQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSpareQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSpareQty.HeaderText = "Spare Qty"
        Me.ColSpareQty.Name = "ColSpareQty"
        Me.ColSpareQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSpareQty.Width = 78
        '
        'ColTotalQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTotalQty.HeaderText = "Total Qty"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTotalQty.Width = 72
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 39
        '
        'ColPrimarySO
        '
        Me.ColPrimarySO.HeaderText = "Primary SO"
        Me.ColPrimarySO.Name = "ColPrimarySO"
        Me.ColPrimarySO.ReadOnly = True
        Me.ColPrimarySO.Visible = False
        Me.ColPrimarySO.Width = 105
        '
        'ColDeliveryDateCont
        '
        Me.ColDeliveryDateCont.HeaderText = "Delivery Date Control"
        Me.ColDeliveryDateCont.Name = "ColDeliveryDateCont"
        Me.ColDeliveryDateCont.Width = 148
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.Width = 99
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier SO"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 111
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 110
        '
        'ColPrimaryUnitMulti
        '
        Me.ColPrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColPrimaryUnitMulti.Name = "ColPrimaryUnitMulti"
        Me.ColPrimaryUnitMulti.Visible = False
        Me.ColPrimaryUnitMulti.Width = 143
        '
        'ColRatioDefault
        '
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 115
        '
        'ColRatioReference
        '
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 136
        '
        'ColOrderQtyMulti
        '
        Me.ColOrderQtyMulti.HeaderText = "Qty Order Multi"
        Me.ColOrderQtyMulti.Name = "ColOrderQtyMulti"
        Me.ColOrderQtyMulti.Visible = False
        Me.ColOrderQtyMulti.Width = 129
        '
        'ColPrimaryCust
        '
        Me.ColPrimaryCust.HeaderText = "Primary Cust"
        Me.ColPrimaryCust.Name = "ColPrimaryCust"
        Me.ColPrimaryCust.Visible = False
        Me.ColPrimaryCust.Width = 113
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(560, 164)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainSOB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(891, 491)
        Me.Controls.Add(Me.dgvData)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(653, 528)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainSOB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Order Balance Detail"
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustPONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOrderQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSpareQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryDateCont As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOrderQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
