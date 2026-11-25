<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiAOP_AdjustDPP
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.ColItemAdjustmentTimestamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_AdjustUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_AdjustmentReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Adjustment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitsMulti_Item = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault_Item = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference_Item = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_AtAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryGRN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShowDetail = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColATCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc_Item = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitsCodeMulti_Item = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti_Item = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStockOutNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGRNNO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppliersProductDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppliersCodeNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItem = New System.Windows.Forms.DataGridView
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(773, 518)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(106, 32)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(659, 518)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(106, 32)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ColItemAdjustmentTimestamp
        '
        Me.ColItemAdjustmentTimestamp.HeaderText = "Adjust Timestamp"
        Me.ColItemAdjustmentTimestamp.Name = "ColItemAdjustmentTimestamp"
        Me.ColItemAdjustmentTimestamp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItemAdjustmentTimestamp.Width = 126
        '
        'ColItem_AdjustUser
        '
        Me.ColItem_AdjustUser.HeaderText = "Adjust User"
        Me.ColItem_AdjustUser.Name = "ColItem_AdjustUser"
        Me.ColItem_AdjustUser.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_AdjustUser.Visible = False
        Me.ColItem_AdjustUser.Width = 87
        '
        'ColItem_AdjustmentReason
        '
        Me.ColItem_AdjustmentReason.HeaderText = "Adjust Reason"
        Me.ColItem_AdjustmentReason.Name = "ColItem_AdjustmentReason"
        Me.ColItem_AdjustmentReason.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_AdjustmentReason.Visible = False
        Me.ColItem_AdjustmentReason.Width = 106
        '
        'ColItem_Adjustment
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_Adjustment.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItem_Adjustment.HeaderText = "Adjustment"
        Me.ColItem_Adjustment.Name = "ColItem_Adjustment"
        Me.ColItem_Adjustment.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Adjustment.Visible = False
        Me.ColItem_Adjustment.Width = 84
        '
        'ColSecUnit
        '
        Me.ColSecUnit.HeaderText = "@ Unit"
        Me.ColSecUnit.Name = "ColSecUnit"
        Me.ColSecUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSecUnit.Width = 57
        '
        'ColSecQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSecQty.HeaderText = "@"
        Me.ColSecQty.Name = "ColSecQty"
        Me.ColSecQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSecQty.Width = 28
        '
        'ColMasterUnits
        '
        Me.ColMasterUnits.HeaderText = "Master Unit"
        Me.ColMasterUnits.Name = "ColMasterUnits"
        Me.ColMasterUnits.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMasterUnits.Visible = False
        Me.ColMasterUnits.Width = 86
        '
        'ColQtyPerMaster
        '
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 109
        '
        'ColQty
        '
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Visible = False
        Me.ColQty.Width = 36
        '
        'ColQtyMaster
        '
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 83
        '
        'ColPrimaryUnitsMulti_Item
        '
        Me.ColPrimaryUnitsMulti_Item.HeaderText = "Primary Units Multi"
        Me.ColPrimaryUnitsMulti_Item.Name = "ColPrimaryUnitsMulti_Item"
        Me.ColPrimaryUnitsMulti_Item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnitsMulti_Item.Visible = False
        Me.ColPrimaryUnitsMulti_Item.Width = 131
        '
        'ColRatioDefault_Item
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioDefault_Item.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRatioDefault_Item.HeaderText = "Ratio Default"
        Me.ColRatioDefault_Item.Name = "ColRatioDefault_Item"
        Me.ColRatioDefault_Item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatioDefault_Item.Visible = False
        Me.ColRatioDefault_Item.Width = 96
        '
        'ColRatioReference_Item
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioReference_Item.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColRatioReference_Item.HeaderText = "Ratio Reference"
        Me.ColRatioReference_Item.Name = "ColRatioReference_Item"
        Me.ColRatioReference_Item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatioReference_Item.Visible = False
        Me.ColRatioReference_Item.Width = 117
        '
        'ColDiscFromPerct
        '
        Me.ColDiscFromPerct.HeaderText = "DiscFromPerct"
        Me.ColDiscFromPerct.Name = "ColDiscFromPerct"
        Me.ColDiscFromPerct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscFromPerct.Visible = False
        Me.ColDiscFromPerct.Width = 106
        '
        'Col_AtAmount
        '
        Me.Col_AtAmount.HeaderText = "At Amount"
        Me.Col_AtAmount.Name = "Col_AtAmount"
        Me.Col_AtAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_AtAmount.Visible = False
        Me.Col_AtAmount.Width = 79
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "PrimaryCurrency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 119
        '
        'ColPrimarySecUnit
        '
        Me.ColPrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColPrimarySecUnit.Name = "ColPrimarySecUnit"
        Me.ColPrimarySecUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimarySecUnit.Visible = False
        Me.ColPrimarySecUnit.Width = 119
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Units"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 98
        '
        'ColPrimaryMasterUnits
        '
        Me.ColPrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColPrimaryMasterUnits.Name = "ColPrimaryMasterUnits"
        Me.ColPrimaryMasterUnits.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryMasterUnits.Visible = False
        Me.ColPrimaryMasterUnits.Width = 138
        '
        'ColPrimaryGRN
        '
        Me.ColPrimaryGRN.HeaderText = "PrimaryGRN"
        Me.ColPrimaryGRN.Name = "ColPrimaryGRN"
        Me.ColPrimaryGRN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryGRN.Visible = False
        Me.ColPrimaryGRN.Width = 93
        '
        'ColPrimaryPO
        '
        Me.ColPrimaryPO.HeaderText = "Primary PO"
        Me.ColPrimaryPO.Name = "ColPrimaryPO"
        Me.ColPrimaryPO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryPO.Visible = False
        Me.ColPrimaryPO.Width = 86
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 113
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'ColShowDetail
        '
        Me.ColShowDetail.HeaderText = "Detail Calculation"
        Me.ColShowDetail.Name = "ColShowDetail"
        Me.ColShowDetail.Width = 123
        '
        'ColATCost
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColATCost.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColATCost.HeaderText = "At Cost (IDR)"
        Me.ColATCost.Name = "ColATCost"
        Me.ColATCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColATCost.Width = 96
        '
        'ColAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAmount.Width = 62
        '
        'ColNettPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNettPrice.Width = 76
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDiscAmount.HeaderText = "Disc. Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscAmount.Width = 97
        '
        'ColDiscount
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscount.Width = 95
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 75
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        Me.ColCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCurrency.Width = 71
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBatchNo.Width = 72
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 95
        '
        'ColTotalQty
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColTotalQty.HeaderText = "Qty Default"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTotalQty.Width = 85
        '
        'ColRatioDesc_Item
        '
        Me.ColRatioDesc_Item.HeaderText = "Ratio"
        Me.ColRatioDesc_Item.Name = "ColRatioDesc_Item"
        Me.ColRatioDesc_Item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatioDesc_Item.Width = 47
        '
        'ColUnitsCodeMulti_Item
        '
        Me.ColUnitsCodeMulti_Item.HeaderText = "Units Code"
        Me.ColUnitsCodeMulti_Item.Name = "ColUnitsCodeMulti_Item"
        Me.ColUnitsCodeMulti_Item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitsCodeMulti_Item.Width = 83
        '
        'ColQtyMulti_Item
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti_Item.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColQtyMulti_Item.HeaderText = "Qty"
        Me.ColQtyMulti_Item.Name = "ColQtyMulti_Item"
        Me.ColQtyMulti_Item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyMulti_Item.Width = 36
        '
        'ColSuppDONo
        '
        Me.ColSuppDONo.HeaderText = "Supplier DO No"
        Me.ColSuppDONo.Name = "ColSuppDONo"
        Me.ColSuppDONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSuppDONo.Width = 113
        '
        'ColStockOutNo
        '
        Me.ColStockOutNo.HeaderText = "Stock In No"
        Me.ColStockOutNo.Name = "ColStockOutNo"
        Me.ColStockOutNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStockOutNo.Width = 86
        '
        'ColGRNNO
        '
        Me.ColGRNNO.HeaderText = "GRN No"
        Me.ColGRNNO.Name = "ColGRNNO"
        Me.ColGRNNO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColGRNNO.Visible = False
        Me.ColGRNNO.Width = 67
        '
        'ColPONO
        '
        Me.ColPONO.HeaderText = "PO. No"
        Me.ColPONO.Name = "ColPONO"
        Me.ColPONO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPONO.Width = 60
        '
        'ColSuppliersProductDesc
        '
        Me.ColSuppliersProductDesc.HeaderText = "Supplier's Product Desc"
        Me.ColSuppliersProductDesc.Name = "ColSuppliersProductDesc"
        Me.ColSuppliersProductDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSuppliersProductDesc.Visible = False
        Me.ColSuppliersProductDesc.Width = 165
        '
        'ColSuppliersCodeNumber
        '
        Me.ColSuppliersCodeNumber.HeaderText = "Supplier's Code Number"
        Me.ColSuppliersCodeNumber.Name = "ColSuppliersCodeNumber"
        Me.ColSuppliersCodeNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSuppliersCodeNumber.Visible = False
        Me.ColSuppliersCodeNumber.Width = 167
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescription.Width = 85
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 69
        '
        'ColItem_No
        '
        Me.ColItem_No.Frozen = True
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 32
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle13.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle13
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColCodeNo, Me.ColDescription, Me.ColSuppliersCodeNumber, Me.ColSuppliersProductDesc, Me.ColPONO, Me.ColGRNNO, Me.ColStockOutNo, Me.ColSuppDONo, Me.ColQtyMulti_Item, Me.ColUnitsCodeMulti_Item, Me.ColRatioDesc_Item, Me.ColTotalQty, Me.ColUnit, Me.ColBatchNo, Me.ColCurrency, Me.ColUnitPrice, Me.ColDiscount, Me.ColDiscAmount, Me.ColNettPrice, Me.ColAmount, Me.ColATCost, Me.ColShowDetail, Me.ColRemarks, Me.ColPrimaryCodeNo, Me.ColPrimaryPO, Me.ColPrimaryGRN, Me.ColPrimaryMasterUnits, Me.ColPrimaryUnit, Me.ColPrimarySecUnit, Me.ColPrimaryCurrency, Me.Col_AtAmount, Me.ColDiscFromPerct, Me.ColRatioReference_Item, Me.ColRatioDefault_Item, Me.ColPrimaryUnitsMulti_Item, Me.ColQtyMaster, Me.ColQty, Me.ColQtyPerMaster, Me.ColMasterUnits, Me.ColSecQty, Me.ColSecUnit, Me.ColItem_Adjustment, Me.ColItem_AdjustmentReason, Me.ColItem_AdjustUser, Me.ColItemAdjustmentTimestamp})
        Me.dgvItem.Location = New System.Drawing.Point(13, 13)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(866, 497)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 2
        '
        'FrmTransaksiAOP_AdjustDPP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 563)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgvItem)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "FrmTransaksiAOP_AdjustDPP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Price Adjustment"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColItemAdjustmentTimestamp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_AdjustUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_AdjustmentReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Adjustment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitsMulti_Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault_Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference_Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_AtAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryGRN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShowDetail As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColATCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc_Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitsCodeMulti_Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti_Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStockOutNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGRNNO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppliersProductDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppliersCodeNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
End Class
